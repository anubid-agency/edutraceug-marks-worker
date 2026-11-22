// edutraceug-marks-worker (D1)
//
// Records assessment marks, computes term reports, generates printable report cards.
// Uses Cloudflare D1 (binding: DB) for all storage. Firebase ID token verification unchanged.
//
// Tables owned by this worker (write):
//   marks, mark_sheets, report_holds, report_notes,
//   subject_summaries, term_summaries, audit_log
//
// Tables owned by other workers (read; a few columns written where noted):
//   roster (roster-worker)            — record_type, class_name, stream, first_name, last_name, name_key, gender, subjects_json
//   terms (teachers-worker)           — name, year, start_date, end_date, next_term_start, is_current, publish_rank
//   assessment_types (teachers-worker)— name, order_, archived
//   term_assessment_settings          — weight, max_score per (term, type)
//   school_settings                   — scoring_mode, cbc_levels_json, grading_scale_json,
//                                       points_bands_json, subsidiary_pass_mark, publish_position,
//                                       card_motto, card_website, stamp_url, teacher_access_mode
//   teacher_settings (teachers-worker)— can_enter_marks, class_teacher_class, class_teacher_stream
//   teacher_assignments               — class_name, stream, subject
//   schools (schools-worker)          — name, location, phone, email, logo_url, theme_color
//
// Secrets:
//   STATE_SIGNING_KEY   (HMAC key for ambiguous-upload state tokens)
//   CLOUDINARY_URL
//   CLOUDMERSIVE_API_KEY
//
// Project: edutrace-ug (JWKS-verified Firebase ID tokens)

import * as XLSX from "xlsx";

// ================================================================
// Constants
// ================================================================

const FIREBASE_PROJECT_ID = "edutrace-ug";

const MARK_STATUSES = ["entered", "absent", "excused", "na"];
const SHEET_STATUSES = ["open", "submitted", "locked"];
const SCORING_MODES = ["percentage", "cbc"];

const DEFAULT_MAX_SCORES_BY_NAME = {
  BOT: 20, CA: 20, AOI: 20, COURSEWORK: 20, MIDTERM: 20, ENDTERM: 100,
};
const FALLBACK_MAX_SCORE = 100;

const DEFAULT_GRADING_SCALE = {
  bands: [
    { min: 80, max: 100, grade: "A", points: 1, descriptor: "Excellent" },
    { min: 70, max: 80, grade: "B", points: 2, descriptor: "Very good" },
    { min: 60, max: 70, grade: "C", points: 3, descriptor: "Good" },
    { min: 50, max: 60, grade: "D", points: 4, descriptor: "Fair" },
    { min: 40, max: 50, grade: "E", points: 5, descriptor: "Weak" },
    { min: 0, max: 40, grade: "F", points: 6, descriptor: "Fail" },
  ],
};

const DEFAULT_POINTS_BANDS = [
  { min: 80, grade: "A" },
  { min: 70, grade: "B" },
  { min: 60, grade: "C" },
  { min: 50, grade: "D" },
  { min: 40, grade: "E" },
  { min: 30, grade: "O" },
  { min: 0,  grade: "F" },
];
const DEFAULT_SUBSIDIARY_PASS_MARK = 50;

const PRINCIPAL_POINTS = { A: 6, B: 5, C: 4, D: 3, E: 2, O: 1, F: 0 };

const MAX_FILE_SIZE_MB = 25;
const BATCH_SIZE = 100;
const MAX_GRID_ROWS = 500;
const MAX_ROSTER_ROWS = 20000;
const MAX_SHEETS = 2000;
const MAX_HOLDS = 2000;
const MAX_TYPES = 100;

const SCAN_URL = "https://api.cloudmersive.com/virus/scan/file";
const SCAN_TIMEOUT_MS = 60_000;

const ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"];
const ACCEPTED_MIME = new Set([
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  "application/vnd.ms-excel",
  "application/vnd.ms-excel.sheet.macroenabled.12",
  "text/csv",
  "application/csv",
  "text/plain",
  "application/octet-stream",
  "",
]);

const ALIASES = {
  rosterId: ["rosterid", "id", "studentid", "regno", "regnumber", "registration", "lin"],
  firstName: ["firstname", "givenname", "forename", "first", "fname", "christianname"],
  lastName: ["lastname", "surname", "familyname", "secondname", "last", "lname"],
  name: ["name", "studentname", "student", "fullname", "pupil", "pupilname"],
  class: ["class", "classname", "grade", "form", "level"],
  stream: ["stream", "section", "arm", "division"],
};

const ABSENT_TOKENS = new Set(["abs", "absent", "a"]);
const EXCUSED_TOKENS = new Set(["exc", "excused", "e"]);
const NA_TOKENS = new Set(["n/a", "na", "-", "--", "—"]);

// ================================================================
// Errors
// ================================================================

class HttpError extends Error {
  constructor(status, message, extra = {}) {
    super(message);
    this.status = status;
    this.extra = extra;
  }
}

// ================================================================
// HTTP helpers
// ================================================================

const CORS_HEADERS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, PATCH, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Authorization, Content-Type",
  "Access-Control-Max-Age": "86400",
};

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json", "Cache-Control": "no-store", ...CORS_HEADERS },
  });
}
function html(body, status = 200) {
  return new Response(body, {
    status,
    headers: { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store", ...CORS_HEADERS },
  });
}
function binary(body, contentType, extraHeaders = {}) {
  return new Response(body, {
    status: 200,
    headers: { "Content-Type": contentType, "Cache-Control": "no-store", ...CORS_HEADERS, ...extraHeaders },
  });
}

async function safeJson(request) {
  try {
    const body = await request.json();
    if (!body || typeof body !== "object" || Array.isArray(body)) throw new Error("not object");
    return body;
  } catch {
    throw new HttpError(400, "Invalid JSON body");
  }
}

// ================================================================
// Base64 / UTF-8
// ================================================================

function utf8ToBytes(str) { return new TextEncoder().encode(str); }
function bytesToBase64Url(bytes) {
  const arr = new Uint8Array(bytes);
  let bin = "";
  const CHUNK = 0x8000;
  for (let i = 0; i < arr.length; i += CHUNK) bin += String.fromCharCode.apply(null, arr.subarray(i, i + CHUNK));
  return btoa(bin).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}
function b64urlToBytes(s) {
  s = String(s).replace(/-/g, "+").replace(/_/g, "/");
  while (s.length % 4) s += "=";
  const bin = atob(s);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}
function b64urlToString(s) { return new TextDecoder().decode(b64urlToBytes(s)); }

// ================================================================
// Text helpers
// ================================================================

function str(v) { return v === null || v === undefined ? "" : String(v); }
function toNumberOrNull(v) {
  if (v === null || v === undefined || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}
function round1(n) { return Math.round(n * 10) / 10; }
function round2(n) { return Math.round(n * 100) / 100; }

const EMOJI_RE = /[\u{1F000}-\u{1FAFF}\u{1F1E6}-\u{1F1FF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{FE0F}\u{200D}\u{20E3}\u{1F3FB}-\u{1F3FF}]/gu;

function escapeHtml(s) {
  return str(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
}
function cleanText(input, maxLen, label) {
  if (typeof input !== "string") throw new HttpError(400, `${label} must be text`);
  const t = input.trim().replace(EMOJI_RE, "");
  if (!t) throw new HttpError(400, `${label} is required`);
  if (t.length > maxLen) throw new HttpError(400, `${label} is too long (max ${maxLen})`);
  return t;
}
function cleanOptionalText(input, maxLen, label) {
  if (input === null || input === undefined || input === "") return "";
  return cleanText(input, maxLen, label);
}
function normalizeHeaderCell(v) {
  return String(v ?? "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// ================================================================
// Names
// ================================================================

function nameKey(firstName, lastName) {
  const norm = (v) =>
    String(v ?? "").toLowerCase().replace(/[^\p{L}\p{N}\s'-]+/gu, " ").replace(/\s+/g, " ").trim();
  const words = `${norm(firstName)} ${norm(lastName)}`.split(" ").filter(Boolean);
  words.sort();
  return words.join(" ");
}

function splitFullName(full) {
  const parts = String(full || "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return { firstName: "", lastName: "" };
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };
  return {
    firstName: parts.slice(0, -1).join(" "),
    lastName: parts[parts.length - 1],
  };
}

function displayName(firstName, lastName, fullName) {
  if (firstName || lastName) return `${str(firstName)} ${str(lastName)}`.trim();
  if (fullName) {
    const { firstName: f, lastName: l } = splitFullName(fullName);
    return `${f} ${l}`.trim();
  }
  return "";
}

// ================================================================
// Deterministic IDs
// ================================================================

async function sha256Hex(s) {
  const buf = await crypto.subtle.digest("SHA-256", utf8ToBytes(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
async function computeMarkId(termId, rosterId, subject, typeId) {
  return (await sha256Hex(`${termId}|${rosterId}|${subject}|${typeId}`)).slice(0, 24);
}
async function computeSheetId(termId, className, stream, subject) {
  return (await sha256Hex(`${termId}|${className}|${stream}|${subject}`)).slice(0, 24);
}

// ================================================================
// Firebase ID token verification (JWKS; project edutrace-ug)
// ================================================================

let jwksCache = { keys: null, expiry: 0 };
async function getJwks() {
  const now = Date.now();
  if (jwksCache.keys && now < jwksCache.expiry) return jwksCache.keys;
  const res = await fetch("https://www.googleapis.com/service_accounts/v1/jwk/securetoken@system.gserviceaccount.com");
  if (!res.ok) throw new HttpError(502, "JWKS fetch failed");
  const data = await res.json();
  jwksCache = { keys: data.keys, expiry: now + 3600_000 };
  return data.keys;
}

async function verifyIdToken(token, projectId) {
  const parts = String(token || "").split(".");
  if (parts.length !== 3) throw new HttpError(401, "Malformed token");
  const [headerB64, payloadB64, sigB64] = parts;

  let header, payload;
  try {
    header = JSON.parse(b64urlToString(headerB64));
    payload = JSON.parse(b64urlToString(payloadB64));
  } catch {
    throw new HttpError(401, "Malformed token");
  }

  const keys = await getJwks();
  const jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) throw new HttpError(401, "Unknown signing key");

  const { alg, ...cleanJwk } = jwk;
  const key = await crypto.subtle.importKey(
    "jwk", cleanJwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]
  );
  const ok = await crypto.subtle.verify(
    "RSASSA-PKCS1-v1_5", key, b64urlToBytes(sigB64), utf8ToBytes(`${headerB64}.${payloadB64}`)
  );
  if (!ok) throw new HttpError(401, "Invalid token signature");

  const now = Math.floor(Date.now() / 1000);
  if (payload.exp < now) throw new HttpError(401, "Token expired");
  if (payload.iat > now + 60) throw new HttpError(401, "Token issued in future");
  if (payload.aud !== projectId) throw new HttpError(401, "Invalid audience");
  if (payload.iss !== `https://securetoken.google.com/${projectId}`) throw new HttpError(401, "Invalid issuer");
  if (!payload.sub) throw new HttpError(401, "Missing subject");
  return payload;
}

async function requireUser(request, env, schoolId, allowedRoles) {
  const auth = request.headers.get("Authorization") || "";
  const m = auth.match(/^Bearer\s+(.+)$/i);
  if (!m) throw new HttpError(401, "Missing Bearer token");
  const payload = await verifyIdToken(m[1], FIREBASE_PROJECT_ID);
  if (!allowedRoles.includes(payload.role)) throw new HttpError(403, "Not allowed for your role");
  if (String(payload.schoolId) !== String(schoolId)) throw new HttpError(403, "Not a member of this school");
  return payload;
}

// ================================================================
// In-worker cache
// ================================================================

const _cache = new Map();
function cacheGet(key) {
  const entry = _cache.get(key);
  if (!entry) return null;
  if (entry.expiry < Date.now()) { _cache.delete(key); return null; }
  return entry.value;
}
function cacheSet(key, value, ttlMs) {
  _cache.set(key, { value, expiry: Date.now() + ttlMs });
}
const CACHE_TTL_MS = 3 * 60_000;

// ================================================================
// D1 helpers
// ================================================================

function db(env) {
  if (!env.DB) throw new HttpError(500, "Server misconfigured: DB binding missing");
  return env.DB;
}

async function d1Get(env, sql, ...params) {
  const stmt = db(env).prepare(sql);
  const row = await (params.length ? stmt.bind(...params) : stmt).first();
  return row || null;
}
async function d1All(env, sql, ...params) {
  const stmt = db(env).prepare(sql);
  const res = await (params.length ? stmt.bind(...params) : stmt).all();
  return res.results || [];
}
async function d1Run(env, sql, ...params) {
  const stmt = db(env).prepare(sql);
  return (params.length ? stmt.bind(...params) : stmt).run();
}

async function runBatch(env, statements) {
  if (statements.length === 0) return;
  const conn = db(env);
  for (let i = 0; i < statements.length; i += BATCH_SIZE) {
    const slice = statements.slice(i, i + BATCH_SIZE).map((s) =>
      s.params && s.params.length ? conn.prepare(s.sql).bind(...s.params) : conn.prepare(s.sql)
    );
    await conn.batch(slice);
  }
}

// ================================================================
// Audit log — never insert id; everything else goes into details_json
// ================================================================

function auditStmt(schoolId, entry) {
  const at = new Date().toISOString();
  const details = { ...entry };
  delete details.action;
  delete details.by;
  delete details.target;
  const byUid = entry && (entry.by || entry.byUid) ? String(entry.by || entry.byUid) : null;
  let target = null;
  if (entry && entry.target) target = String(entry.target);
  else if (entry) {
    const bits = [entry.termId, entry.class, entry.stream, entry.subject, entry.rosterId]
      .filter((v) => v !== null && v !== undefined && v !== "");
    if (bits.length) target = bits.join("|");
  }
  return {
    sql: `INSERT INTO audit_log (school_id, action, by_uid, target, details_json, at)
          VALUES (?, ?, ?, ?, ?, ?)`,
    params: [schoolId, entry.action, byUid, target, JSON.stringify(details), at],
  };
}

async function logChange(env, schoolId, entry) {
  try {
    const stmt = auditStmt(schoolId, entry);
    await d1Run(env, stmt.sql, ...stmt.params);
  } catch (err) {
    console.error("Audit log write failed:", err && err.message);
  }
}

// ================================================================
// School settings (single read; used for scoring config, theme, access mode)
// ================================================================

async function getSchoolSettings(env, schoolId) {
  const key = `school_settings:${schoolId}`;
  const cached = cacheGet(key);
  if (cached) return cached;
  let row = null;
  try {
    row = await d1Get(env, `SELECT * FROM school_settings WHERE school_id = ?`, schoolId);
  } catch { row = null; }
  const scoringMode = row && SCORING_MODES.includes(row.scoring_mode) ? row.scoring_mode : "percentage";
  let cbcLevels = null;
  if (row && row.cbc_levels_json) {
    try {
      const parsed = JSON.parse(row.cbc_levels_json);
      if (Array.isArray(parsed) && parsed.length > 0) cbcLevels = parsed;
    } catch { /* ignore */ }
  }
  if (!cbcLevels) {
    cbcLevels = [
      { level: 1, label: "Beginning", descriptor: "Needs support" },
      { level: 2, label: "Developing", descriptor: "Grasping with guidance" },
      { level: 3, label: "Achieved", descriptor: "Consistently meets standard" },
    ];
  }
  let gradingScale = DEFAULT_GRADING_SCALE;
  if (row && row.grading_scale_json) {
    try {
      const parsed = JSON.parse(row.grading_scale_json);
      if (parsed && Array.isArray(parsed.bands) && parsed.bands.length > 0) gradingScale = { bands: parsed.bands };
    } catch { /* ignore */ }
  }
  let pointsBands = DEFAULT_POINTS_BANDS;
  if (row && row.points_bands_json) {
    try {
      const parsed = JSON.parse(row.points_bands_json);
      if (Array.isArray(parsed) && parsed.length > 0) pointsBands = parsed;
    } catch { /* ignore */ }
  }
  const subsidiaryPassMark =
    row && Number.isFinite(row.subsidiary_pass_mark) ? row.subsidiary_pass_mark : DEFAULT_SUBSIDIARY_PASS_MARK;
  const teacherAccessMode = row && row.teacher_access_mode === "assigned" ? "assigned" : "open";
  const publishPosition = !!(row && (row.publish_position === 1 || row.publish_position === true));
  let subjectCodes = {};
  if (row && row.subject_codes_json) {
    try { subjectCodes = JSON.parse(row.subject_codes_json) || {}; } catch { subjectCodes = {}; }
  }
  const out = {
    scoringMode,
    cbcLevels,
    gradingScale,
    pointsBands,
    subsidiaryPassMark,
    teacherAccessMode,
    publishPosition,
    cardMotto: (row && row.card_motto) || "",
    cardWebsite: (row && row.card_website) || "",
    stampUrl: (row && row.stamp_url) || "",
    subjectCodes,
  };
  cacheSet(key, out, CACHE_TTL_MS);
  return out;
}

// ================================================================
// Assessment types
// ================================================================

async function listAssessmentTypes(env, schoolId, { includeArchived = false } = {}) {
  const key = `types:${schoolId}:${includeArchived ? "all" : "active"}`;
  const cached = cacheGet(key);
  if (cached) return cached;
  const rows = await d1All(
    env,
    `SELECT id, name, order_ AS ord, archived FROM assessment_types
     WHERE school_id = ? ORDER BY order_ ASC, name ASC LIMIT ?`,
    schoolId, MAX_TYPES
  );
  const types = rows.map((r) => ({
    id: r.id,
    name: r.name || "",
    order: typeof r.ord === "number" ? r.ord : 100,
    archived: r.archived === 1 || r.archived === true,
  }));
  const filtered = includeArchived ? types : types.filter((t) => !t.archived);
  filtered.sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
  cacheSet(key, filtered, CACHE_TTL_MS);
  return filtered;
}

// ================================================================
// Term assessment settings (max scores + weights)
// ================================================================

async function getTermAssessmentSettings(env, schoolId, termId) {
  const key = `tas:${schoolId}:${termId}`;
  const cached = cacheGet(key);
  if (cached) return cached;
  const rows = await d1All(
    env,
    `SELECT type_id, weight, max_score FROM term_assessment_settings
     WHERE school_id = ? AND term_id = ?`,
    schoolId, termId
  );
  const byType = new Map();
  for (const r of rows) {
    byType.set(r.type_id, {
      weight: Number.isFinite(r.weight) ? r.weight : 0,
      maxScore: Number.isFinite(r.max_score) ? r.max_score : null,
    });
  }
  cacheSet(key, byType, CACHE_TTL_MS);
  return byType;
}

function resolveMaxScores(termSettings, types) {
  const out = {};
  for (const t of types) {
    const row = termSettings.get(t.id);
    if (row && row.maxScore && row.maxScore > 0) out[t.id] = row.maxScore;
    else if (DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()]) out[t.id] = DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()];
    else out[t.id] = FALLBACK_MAX_SCORE;
  }
  return out;
}

function resolveWeights(termSettings, types) {
  const out = {};
  for (const t of types) {
    const row = termSettings.get(t.id);
    out[t.id] = row && Number.isFinite(row.weight) ? row.weight : 0;
  }
  return out;
}

function termPublishRank(term) {
  return term && (term.publish_rank === 1 || term.publish_rank === true);
}

// ================================================================
// A-level helpers (S5/S6 only)
// ================================================================

function isALevelClass(className) {
  const c = String(className || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  return /^(senior|s)[56](?!\d)/.test(c);
}

function isSubsidiarySubject(subject) {
  const s = String(subject || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  if (s === "generalpaper" || s === "gp") return true;
  if (s === "ict") return true;
  if (s === "subsidiaryict") return true;
  if (s === "informationcommunicationtechnology") return true;
  if (s === "computer") return true;
  if (s === "subsidiarymathematics") return true;
  if (s === "subsidiarymaths") return true;
  if (s === "subsidiarymath") return true;
  return false;
}

function gradeFromPct(pct, pointsBands) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return null;
  const bands = (pointsBands || DEFAULT_POINTS_BANDS).slice().sort((a, b) => b.min - a.min);
  for (const band of bands) if (pct >= band.min) return band.grade;
  return null;
}


function principalPointsForSubject(pct, subject) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return 0;
  if (isSubsidiarySubject(subject)) return 0;
  if (pct >= 80) return 5;
  if (pct >= 70) return 4;
  if (pct >= 60) return 3;
  if (pct >= 50) return 2;
  return 1;
}
function subsidiaryPointsForSubject(pct, subject, passMark) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return 0;
  if (!isSubsidiarySubject(subject)) return 0;
  return pct >= (passMark ?? DEFAULT_SUBSIDIARY_PASS_MARK) ? 1 : 0;
}

function applyGradingScale(pct, scale) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return null;
  const bands = (scale && Array.isArray(scale.bands) ? scale.bands : DEFAULT_GRADING_SCALE.bands)
    .slice().sort((a, b) => b.min - a.min);
  for (const band of bands) if (pct >= band.min) return band;
  return null;
}

// ================================================================
// Terms (teachers-worker-owned)
// ================================================================

async function getTermOrThrow(env, schoolId, termId) {
  const row = await d1Get(env, `SELECT * FROM terms WHERE school_id = ? AND id = ?`, schoolId, termId);
  if (!row) throw new HttpError(404, "Term not found");
  return row;
}

async function resolveTermId(env, schoolId, requested) {
  if (requested) {
    const ck = `termOk:${schoolId}:${requested}`;
    if (cacheGet(ck)) return requested;
    const t = await d1Get(env, `SELECT id FROM terms WHERE school_id = ? AND id = ?`, schoolId, requested);
    if (!t) throw new HttpError(404, "Term not found");
    cacheSet(ck, true, CACHE_TTL_MS);
    return requested;
  }
  const key = `currentTerm:${schoolId}`;
  const cached = cacheGet(key);
  if (cached) return cached;
  const row = await d1Get(env, `SELECT id FROM terms WHERE school_id = ? AND is_current = 1 LIMIT 1`, schoolId);
  if (!row) throw new HttpError(404, "No current term set");
  cacheSet(key, row.id, CACHE_TTL_MS);
  return row.id;
}

// ================================================================
// Teacher settings + assignments (teachers-worker-owned)
// ================================================================

async function loadTeacherSettings(env, schoolId, uid) {
  const row = await d1Get(env, `SELECT * FROM teacher_settings WHERE school_id = ? AND uid = ?`, schoolId, uid);
  return {
    canEnterMarks: !!(row && (row.can_enter_marks === 1 || row.can_enter_marks === true)),
    canUploadHoliday: !!(row && (row.can_upload_holiday === 1 || row.can_upload_holiday === true)),
    canReadParentReports: !!(row && (row.can_read_parent_reports === 1 || row.can_read_parent_reports === true)),
    classTeacherClass: (row && row.class_teacher_class) || "",
    classTeacherStream: (row && row.class_teacher_stream) || "",
  };
}

async function loadTeacherAssignments(env, schoolId, uid) {
  const rows = await d1All(
    env,
    `SELECT class_name, stream, subject FROM teacher_assignments WHERE school_id = ? AND uid = ?`,
    schoolId, uid
  );
  return rows.map((r) => ({ class: r.class_name || "", stream: r.stream || "", subject: r.subject || "" }));
}

function assignmentCovers(assignments, className, stream, subject) {
  const norm = (v) => String(v || "").trim().toLowerCase();
  return assignments.some((a) =>
    norm(a.class) === norm(className) &&
    (!a.stream || norm(a.stream) === norm(stream)) &&
    norm(a.subject) === norm(subject)
  );
}

function classTeacherCovers(settings, className, stream) {
  if (!settings.classTeacherClass) return false;
  if (String(settings.classTeacherClass) !== String(className)) return false;
  if (settings.classTeacherStream && stream && String(settings.classTeacherStream) !== String(stream)) return false;
  return true;
}

async function requireCanEnterMarks(env, schoolId, uid) {
  const s = await loadTeacherSettings(env, schoolId, uid);
  if (!s.canEnterMarks) throw new HttpError(403, "You do not have the can_enter_marks permission");
  return s;
}

async function assertTeacherAssignment(env, schoolId, uid, className, stream, subject, schoolSettings) {
  if (schoolSettings.teacherAccessMode !== "assigned") return;
  const assignments = await loadTeacherAssignments(env, schoolId, uid);
  if (!assignmentCovers(assignments, className, stream, subject)) {
    throw new HttpError(403, "You are not assigned to this class/stream/subject");
  }
}

async function assertClassTeacher(env, schoolId, uid, className, stream) {
  const s = await loadTeacherSettings(env, schoolId, uid);
  if (!classTeacherCovers(s, className, stream)) {
    throw new HttpError(403, "You are not the class teacher for this class/stream");
  }
}

// ================================================================
// Roster (roster-worker-owned)
// ================================================================

async function fetchRosterStudents(env, schoolId, className, stream) {
  let sql = `SELECT * FROM roster WHERE school_id = ? AND record_type = 'student' AND class_name = ?`;
  const params = [schoolId, className];
  if (stream) { sql += ` AND stream = ?`; params.push(stream); }
  sql += ` ORDER BY last_name ASC, first_name ASC LIMIT ?`;
  params.push(MAX_ROSTER_ROWS);
  const rows = await d1All(env, sql, ...params);
  return rows.map((r) => {
    const firstName = r.first_name || "";
    const lastName = r.last_name || "";
    const nk = r.name_key || nameKey(firstName, lastName);
    let subjects = [];
    try { subjects = JSON.parse(r.subjects_json || "[]"); } catch { subjects = []; }
    return {
      id: r.id,
      firstName, lastName,
      nameKey: nk,
      class: r.class_name || "",
      stream: r.stream || "",
      gender: r.gender || "",
      subjects: Array.isArray(subjects) ? subjects : [],
    };
  });
}

// ================================================================
// Cloudmersive
// ================================================================

async function scanFileForViruses(env, bytes, filename) {
  if (!env.CLOUDMERSIVE_API_KEY) throw new HttpError(500, "Server misconfigured: CLOUDMERSIVE_API_KEY");
  const form = new FormData();
  form.append("inputFile", new Blob([bytes]), filename || "upload");

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), SCAN_TIMEOUT_MS);
  let res;
  try {
    res = await fetch(SCAN_URL, {
      method: "POST",
      headers: { Apikey: env.CLOUDMERSIVE_API_KEY, Accept: "application/json" },
      body: form, signal: controller.signal,
    });
  } catch (err) {
    throw new HttpError(502, `Virus scan service unreachable: ${err.message}`);
  } finally { clearTimeout(timer); }
  if (!res.ok) {
    const detail = await res.text().catch(() => "");
    throw new HttpError(502, `Virus scan failed (${res.status}): ${detail.slice(0, 200)}`);
  }
  const data = await res.json().catch(() => ({}));
  const found = Array.isArray(data.FoundViruses) ? data.FoundViruses : [];
  const clean = data.CleanResult !== false && found.length === 0;
  return { clean, foundViruses: found };
}

// ================================================================
// Cloudinary (raw upload)
// ================================================================

function parseCloudinaryUrl(env) {
  const url = env.CLOUDINARY_URL;
  if (!url) throw new HttpError(500, "Server misconfigured: CLOUDINARY_URL");
  const m = String(url).match(/^cloudinary:\/\/([^:]+):([^@]+)@(.+)$/);
  if (!m) throw new HttpError(500, "CLOUDINARY_URL is malformed");
  return { apiKey: m[1], apiSecret: m[2], cloudName: m[3] };
}
async function sha1Hex(s) {
  const buf = await crypto.subtle.digest("SHA-1", utf8ToBytes(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}
function safeBaseName(name) {
  return String(name || "upload").replace(/\.[^.]+$/, "").replace(/[^a-zA-Z0-9_-]+/g, "_").slice(0, 60) || "upload";
}
async function uploadRawToCloudinary(env, bytes, filename, folder) {
  const { apiKey, apiSecret, cloudName } = parseCloudinaryUrl(env);
  const dot = String(filename || "").lastIndexOf(".");
  const ext = dot > 0 ? filename.slice(dot).toLowerCase() : ".xlsx";
  const publicId = `${safeBaseName(filename)}_${Date.now()}${ext}`;
  const timestamp = Math.floor(Date.now() / 1000);
  const signedParams = { folder, public_id: publicId, timestamp: String(timestamp) };
  const toSign = Object.keys(signedParams).sort().map((k) => `${k}=${signedParams[k]}`).join("&") + apiSecret;
  const signature = await sha1Hex(toSign);

  const form = new FormData();
  form.append("file", new Blob([bytes]), filename || "upload");
  form.append("api_key", apiKey);
  form.append("timestamp", signedParams.timestamp);
  form.append("public_id", publicId);
  form.append("folder", folder);
  form.append("signature", signature);

  const res = await fetch(`https://api.cloudinary.com/v1_1/${cloudName}/raw/upload`, { method: "POST", body: form });
  if (!res.ok) throw new HttpError(502, "Cloudinary upload failed: " + (await res.text()).slice(0, 200));
  const data = await res.json();
  if (!data.secure_url) throw new HttpError(502, "Cloudinary response missing secure_url");
  return { url: data.secure_url, publicId };
}

// ================================================================
// Excel parsing
// ================================================================

function cellStr(values, idx) {
  if (idx === undefined || idx === null || idx < 0) return "";
  const raw = values[idx];
  if (raw === undefined || raw === null) return "";
  return String(raw).replace(/\s+/g, " ").trim();
}

function detectHeaderRow(matrix, activeTypes) {
  const scanLimit = Math.min(matrix.length, 15);
  const typeNamesNormalized = new Map();
  for (const t of activeTypes) typeNamesNormalized.set(normalizeHeaderCell(t.name), t.id);

  let best = null;
  for (let r = 0; r < scanLimit; r += 1) {
    const row = matrix[r] || [];
    const map = {};
    let score = 0;
    for (let c = 0; c < row.length; c += 1) {
      const key = normalizeHeaderCell(row[c]);
      if (!key) continue;
      for (const field of ["rosterId", "firstName", "lastName", "name", "class", "stream"]) {
        if (map[field] !== undefined) continue;
        if (ALIASES[field].includes(key)) { map[field] = c; score += 1; break; }
      }
      if (typeNamesNormalized.has(key)) {
        const typeId = typeNamesNormalized.get(key);
        if (!map.types) map.types = {};
        if (map.types[typeId] === undefined) { map.types[typeId] = c; score += 1; }
      }
    }
    const hasIdentity = map.rosterId !== undefined || (map.firstName !== undefined && map.lastName !== undefined) || map.name !== undefined;
    const hasType = map.types && Object.keys(map.types).length > 0;
    if (hasIdentity && hasType && (!best || score > best.score)) best = { headerRowIndex: r, map, score, positional: false };
  }
  if (best) return best;

  const map = { types: {} };
  map.rosterId = 0;
  map.lastName = 1;
  let col = 2;
  for (const t of activeTypes) { map.types[t.id] = col; col += 1; }
  return { headerRowIndex: -1, map, score: 0, positional: true };
}

function readRowsFromWorkbook(buffer, activeTypes) {
  let wb;
  try { wb = XLSX.read(new Uint8Array(buffer), { type: "array", cellDates: false }); }
  catch (err) { throw new HttpError(400, `Could not read the spreadsheet: ${err.message}. Is the file corrupt?`); }
  if (!wb.SheetNames || wb.SheetNames.length === 0) throw new HttpError(400, "Workbook has no sheets");
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false, raw: false });
  if (matrix.length === 0) throw new HttpError(400, "Sheet is empty");

  const columns = detectHeaderRow(matrix, activeTypes);
  const startIndex = columns.headerRowIndex + 1;
  const rows = [];
  for (let i = startIndex; i < matrix.length; i += 1) {
    const raw = matrix[i] || [];
    const blank = !raw.some((v) => String(v ?? "").trim() !== "");
    if (blank) continue;
    if (/^enter (scores|levels)/i.test(String(raw[0] ?? "").trim())) continue;
    const row = {
      excelRow: i + 1,
      rosterId: cellStr(raw, columns.map.rosterId),
      firstName: cellStr(raw, columns.map.firstName),
      lastName: cellStr(raw, columns.map.lastName),
      fullName: cellStr(raw, columns.map.name),
      class: cellStr(raw, columns.map.class),
      stream: cellStr(raw, columns.map.stream),
      cells: {},
    };
    if (columns.map.types) {
      for (const [typeId, colIdx] of Object.entries(columns.map.types)) row.cells[typeId] = cellStr(raw, colIdx);
    }
    rows.push(row);
  }
  if (rows.length === 0) throw new HttpError(400, "No data rows found below the header");
  return { columns, rows };
}

function interpretCell(raw, scoringMode, maxLevel) {
  const v = String(raw ?? "").trim();
  if (v === "") return null;
  const lower = v.toLowerCase();
  if (ABSENT_TOKENS.has(lower)) return { status: "absent", score: null, level: null };
  if (EXCUSED_TOKENS.has(lower)) return { status: "excused", score: null, level: null };
  if (NA_TOKENS.has(v) || NA_TOKENS.has(lower)) return { status: "na", score: null, level: null };
  const num = Number(v);
  if (!Number.isFinite(num)) return { error: `"${v}" is not a number or known marker` };
  if (scoringMode === "cbc") {
    if (!Number.isInteger(num) || num < 1 || num > maxLevel) return { error: `Level must be an integer 1-${maxLevel} (got "${v}")` };
    return { status: "entered", score: null, level: num };
  }
  if (num < 0) return { error: `Score cannot be negative ("${v}")` };
  return { status: "entered", score: num, level: null };
}

// ================================================================
// Mark sheet state
// ================================================================

async function getSheet(env, schoolId, sheetId) {
  return d1Get(env, `SELECT * FROM mark_sheets WHERE school_id = ? AND id = ?`, schoolId, sheetId);
}

async function assertSheetWritable(env, schoolId, termId, className, stream, subject, user) {
  const sheetId = await computeSheetId(termId, className, stream, subject);
  const sheet = await getSheet(env, schoolId, sheetId);
  if (!sheet) return { sheetId, sheet: null };
  if (sheet.status === "locked") throw new HttpError(409, "This mark sheet is locked. Ask an admin to reopen it.");
  if (sheet.status === "submitted" && user.role !== "schoolAdmin") {
    throw new HttpError(409, "This mark sheet has been submitted. Ask an admin to reopen it before editing.");
  }
  return { sheetId, sheet };
}

// ================================================================
// State signing (HMAC key from env.STATE_SIGNING_KEY)
// ================================================================

async function getStateKey(env) {
  const raw = env.STATE_SIGNING_KEY;
  if (!raw) throw new HttpError(500, "Server misconfigured: STATE_SIGNING_KEY");
  return crypto.subtle.importKey("raw", utf8ToBytes(String(raw)), { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
}
async function signState(env, payload) {
  const key = await getStateKey(env);
  const body = bytesToBase64Url(utf8ToBytes(JSON.stringify(payload)));
  const sig = await crypto.subtle.sign("HMAC", key, utf8ToBytes(body));
  return `${body}.${bytesToBase64Url(new Uint8Array(sig))}`;
}
async function verifyState(env, token) {
  if (typeof token !== "string" || !token.includes(".")) throw new HttpError(400, "Malformed state token");
  const [body, sigB64] = token.split(".");
  if (!body || !sigB64) throw new HttpError(400, "Malformed state token");
  const key = await getStateKey(env);
  const valid = await crypto.subtle.verify("HMAC", key, b64urlToBytes(sigB64), utf8ToBytes(body));
  if (!valid) throw new HttpError(400, "Invalid state token signature");
  const payload = JSON.parse(b64urlToString(body));
  if (payload.exp && payload.exp < Math.floor(Date.now() / 1000)) throw new HttpError(400, "State token expired");
  return payload;
}

// ================================================================
// Grid: manual entry
// ================================================================

async function fetchGridMarks(env, schoolId, termId, className, stream, subject, typeId) {
  let sql = `SELECT * FROM marks WHERE school_id = ? AND term_id = ? AND class_name = ? AND subject = ? AND type_id = ?`;
  const params = [schoolId, termId, className, subject, typeId];
  if (stream) { sql += ` AND stream = ?`; params.push(stream); }
  sql += ` LIMIT ?`;
  params.push(MAX_GRID_ROWS);
  const rows = await d1All(env, sql, ...params);
  const byRoster = new Map();
  for (const r of rows) byRoster.set(r.roster_id, r);
  return byRoster;
}

async function handleGetGrid(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";
  const subject = cleanText(url.searchParams.get("subject") || "", 120, "subject");
  const typeId = cleanText(url.searchParams.get("typeId") || "", 40, "typeId");

  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }

  const [term, types, termSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getTermAssessmentSettings(env, schoolId, termId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, `Assessment type "${typeId}" not found or archived`);

  const maxScores = resolveMaxScores(termSettings, types);
  const weights = resolveWeights(termSettings, types);
  const [students, marksByRoster] = await Promise.all([
    fetchRosterStudents(env, schoolId, className, stream),
    fetchGridMarks(env, schoolId, termId, className, stream, subject, typeId),
  ]);
  const sheetId = await computeSheetId(termId, className, stream, subject);
  const sheet = await getSheet(env, schoolId, sheetId);

  let enteredCount = 0;
  const rows = students.map((s) => {
    const m = marksByRoster.get(s.id);
    const status = m ? m.status : "missing";
    if (status !== "missing") enteredCount += 1;
    return {
      rosterId: s.id,
      firstName: s.firstName || "",
      lastName: s.lastName || "",
      gender: s.gender || "",
      markId: m ? m.id : null,
      status,
      score: m && typeof m.score === "number" ? m.score : null,
      level: m && typeof m.level === "number" ? m.level : null,
      maxScore: m && typeof m.max_score === "number" ? m.max_score : maxScores[typeId],
      updatedAt: m ? m.updated_at : null,
      updatedBy: m ? m.updated_by : null,
    };
  });

  return {
    scoringMode: schoolSettings.scoringMode,
    cbcLevels: schoolSettings.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null },
    class: className, stream, subject, typeId,
    typeName: types.find((t) => t.id === typeId)?.name || typeId,
    maxScore: maxScores[typeId],
    weight: weights[typeId],
    sheet: sheet
      ? { id: sheetId, status: sheet.status, submittedAt: sheet.submitted_at || null, lockedAt: sheet.locked_at || null }
      : { id: sheetId, status: "open", submittedAt: null, lockedAt: null },
    progress: { entered: enteredCount, total: rows.length },
    rows,
  };
}

async function handlePutGrid(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const className = cleanText(body.class, 60, "class");
  const stream = body.stream ? cleanText(body.stream, 60, "stream") : "";
  const subject = cleanText(body.subject, 120, "subject");
  const typeId = cleanText(body.typeId, 40, "typeId");

  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);
  if (!Array.isArray(body.entries) || body.entries.length === 0) throw new HttpError(400, "entries required");
  if (body.entries.length > MAX_GRID_ROWS) throw new HttpError(400, `Max ${MAX_GRID_ROWS} entries`);

  const [term, types, termSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getTermAssessmentSettings(env, schoolId, termId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, "Assessment type not found");

  const maxScores = resolveMaxScores(termSettings, types);
  const defaultMax = toNumberOrNull(body.maxScore) || maxScores[typeId];
  const maxLevel = schoolSettings.cbcLevels.length;

  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const validRoster = new Set(students.map((s) => s.id));

  const nowIso = new Date().toISOString();
  const uid = String(user.sub);
  const stmts = [];
  const auditRows = [];
  const affectedRoster = new Set();

  for (const e of body.entries) {
    if (!e || typeof e !== "object") throw new HttpError(400, "Each entry must be an object");
    const rosterId = str(e.rosterId).trim();
    if (!rosterId || !validRoster.has(rosterId)) throw new HttpError(400, `rosterId not in class: ${rosterId}`);

    const status = MARK_STATUSES.includes(e.status) ? e.status : null;
    if (!status) throw new HttpError(400, `Invalid status for ${rosterId}`);

    let score = null, level = null, cellMax = null;
    if (schoolSettings.scoringMode === "percentage") {
      cellMax = toNumberOrNull(e.maxScore) || defaultMax;
      if (!(cellMax > 0)) throw new HttpError(400, `Invalid maxScore for ${rosterId}`);
      if (status === "entered") {
        const s = toNumberOrNull(e.score);
        if (s === null || s < 0 || s > cellMax) throw new HttpError(400, `score for ${rosterId} must be between 0 and ${cellMax}`);
        score = s;
      }
    } else if (status === "entered") {
      const l = toNumberOrNull(e.level);
      if (!Number.isInteger(l) || l < 1 || l > maxLevel) throw new HttpError(400, `level for ${rosterId} must be 1-${maxLevel}`);
      level = l;
    }

    const markId = await computeMarkId(termId, rosterId, subject, typeId);
    stmts.push({
      sql: `INSERT INTO marks (school_id, id, term_id, roster_id, class_name, stream, subject, type_id, scoring_mode, status, score, level, max_score, source, cloudinary_url, entered_by, entered_at, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?)
            ON CONFLICT(school_id, id) DO UPDATE SET
              status = excluded.status, score = excluded.score, level = excluded.level, max_score = excluded.max_score,
              updated_by = excluded.updated_by, updated_at = excluded.updated_at`,
      params: [
        schoolId, markId, termId, rosterId, className, stream || "", subject, typeId, schoolSettings.scoringMode,
        status, score, level, cellMax, uid, nowIso, uid, nowIso,
      ],
    });
    auditRows.push({ rosterId, markId, status, score, level, maxScore: cellMax });
    affectedRoster.add(rosterId);
  }

  await runBatch(env, stmts);

  const summaryStmts = await buildSummariesBatch(env, schoolId, {
    termId, className, stream, subject, typeId,
    rosterIds: Array.from(affectedRoster), schoolSettings, nowIso, uid, types,
  });
  summaryStmts.push(auditStmt(schoolId, {
    action: "marks.grid.write",
    by: uid, termId, class: className, stream, subject, typeId,
    count: body.entries.length, entries: auditRows,
  }));
  await runBatch(env, summaryStmts);
  return { status: "ok", updated: body.entries.length, termId, class: className, stream, subject, typeId };
}

async function handleDeleteMark(request, env, schoolId, markId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const doc = await d1Get(env, `SELECT * FROM marks WHERE school_id = ? AND id = ?`, schoolId, markId);
  if (!doc) throw new HttpError(404, "Mark not found");
  await getTermOrThrow(env, schoolId, doc.term_id);
  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), doc.class_name, doc.stream || "", doc.subject, schoolSettings);
  }
  await assertSheetWritable(env, schoolId, doc.term_id, doc.class_name, doc.stream || "", doc.subject, user);
  const nowIso = new Date().toISOString();
  const uid = String(user.sub);

  await runBatch(env, [
    { sql: `DELETE FROM marks WHERE school_id = ? AND id = ?`, params: [schoolId, markId] },
  ]);
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const summaryStmts = await buildSummariesBatch(env, schoolId, {
    termId: doc.term_id, className: doc.class_name, stream: doc.stream || "", subject: doc.subject, typeId: doc.type_id,
    rosterIds: [doc.roster_id], schoolSettings, nowIso, uid, types,
  });
  summaryStmts.push(auditStmt(schoolId, { action: "marks.cell.delete", by: uid, rosterId: doc.roster_id, termId: doc.term_id }));
  await runBatch(env, summaryStmts);
  return { status: "ok" };
}

// ================================================================
// Summary builders (subject_summaries + term_summaries)
// ================================================================

async function computeSubjectPct(env, schoolId, termId, rosterId, subject, schoolSettings) {
  // Pull all marks for this student+subject; compute weighted pct using
  // term_assessment_settings; fall back to the mark's own max_score if missing.
  const rows = await d1All(
    env,
    `SELECT type_id, status, score, level, max_score FROM marks
     WHERE school_id = ? AND term_id = ? AND roster_id = ? AND subject = ?`,
    schoolId, termId, rosterId, subject
  );
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const termSettings = await getTermAssessmentSettings(env, schoolId, termId);
  const weights = resolveWeights(termSettings, types);

  let weightedRaw = 0, weightsPresent = 0;
  let levelMax = null;
  if (schoolSettings.scoringMode === "cbc") {
    for (const r of rows) if (r.status === "entered" && Number.isFinite(r.level)) {
      if (levelMax === null || r.level > levelMax) levelMax = r.level;
    }
    return { pct: null, level: levelMax, weightsPresent: null };
  }
  for (const r of rows) {
    if (r.status !== "entered" || !Number.isFinite(r.score)) continue;
    const w = weights[r.type_id] ?? 0;
    let mx = r.max_score;
    if (!mx || !(mx > 0)) {
      const ts = termSettings.get(r.type_id);
      if (ts && ts.maxScore && ts.maxScore > 0) mx = ts.maxScore;
      else mx = DEFAULT_MAX_SCORES_BY_NAME[String(r.type_id).toUpperCase()] || FALLBACK_MAX_SCORE;
    }
    const contribution = (r.score / mx) * w;
    weightedRaw += contribution;
    weightsPresent += w;
  }
  const pct = weightsPresent > 0 ? round1((weightedRaw / weightsPresent) * 100) : null;
  return { pct, level: null, weightsPresent: round2(weightsPresent) };
}

async function upsertSubjectSummary(env, schoolId, { termId, rosterId, subject, className, stream, pct, grade, points, level, weightsPresent, isSubsidiary, nowIso }) {
  return {
    sql: `INSERT INTO subject_summaries (school_id, term_id, roster_id, subject, class_name, stream, pct, grade, points, level, weights_present, is_subsidiary, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(school_id, term_id, roster_id, subject) DO UPDATE SET
            class_name = excluded.class_name, stream = excluded.stream,
            pct = excluded.pct, grade = excluded.grade, points = excluded.points, level = excluded.level,
            weights_present = excluded.weights_present, is_subsidiary = excluded.is_subsidiary, updated_at = excluded.updated_at`,
    params: [schoolId, termId, rosterId, subject, className, stream || "", pct, grade, points, level, weightsPresent, isSubsidiary ? 1 : 0, nowIso],
  };
}

async function buildSummariesBatch(env, schoolId, { termId, className, stream, subject, rosterIds, schoolSettings, nowIso, uid, types }) {
  const stmts = [];
  if (!rosterIds || rosterIds.length === 0) return stmts;
  const aLevel = isALevelClass(className);

  // Fetch per-roster+subject rows in one shot.
  const subjectRows = await d1All(
    env,
    `SELECT roster_id, type_id, status, score, level, max_score FROM marks
     WHERE school_id = ? AND term_id = ? AND class_name = ? AND subject = ?`,
    schoolId, termId, className, subject
  );
  const byRoster = new Map();
  for (const r of subjectRows) {
    if (!byRoster.has(r.roster_id)) byRoster.set(r.roster_id, []);
    byRoster.get(r.roster_id).push(r);
  }

  const termSettings = await getTermAssessmentSettings(env, schoolId, termId);
  const weights = resolveWeights(termSettings, types);

  for (const rosterId of rosterIds) {
    const rows = byRoster.get(rosterId) || [];
    let pct = null, level = null, weightsPresent = null;
    if (schoolSettings.scoringMode === "cbc") {
      let highest = null;
      for (const r of rows) if (r.status === "entered" && Number.isFinite(r.level)) {
        if (highest === null || r.level > highest) highest = r.level;
      }
      level = highest;
    } else {
      let weightedRaw = 0, present = 0;
      for (const r of rows) {
        if (r.status !== "entered" || !Number.isFinite(r.score)) continue;
        const w = weights[r.type_id] ?? 0;
        let mx = r.max_score;
        if (!mx || !(mx > 0)) {
          const ts = termSettings.get(r.type_id);
          if (ts && ts.maxScore && ts.maxScore > 0) mx = ts.maxScore;
          else mx = FALLBACK_MAX_SCORE;
        }
        weightedRaw += (r.score / mx) * w;
        present += w;
      }
      weightsPresent = round2(present);
      pct = present > 0 ? round1((weightedRaw / present) * 100) : null;
    }

    const sub = isSubsidiarySubject(subject);
    let grade = null, points = null;
    if (schoolSettings.scoringMode === "percentage") {
      const band = applyGradingScale(pct, schoolSettings.gradingScale);
      grade = band ? band.grade : null;
      if (aLevel) {
        points = sub ? subsidiaryPointsForSubject(pct, subject, schoolSettings.subsidiaryPassMark) : principalPointsForSubject(pct, subject, schoolSettings.pointsBands);
      }
    }

    stmts.push(await upsertSubjectSummary(env, schoolId, {
      termId, rosterId, subject, className, stream,
      pct, grade, points, level, weightsPresent, isSubsidiary: sub, nowIso,
    }));
  }

  return stmts;
}

// ================================================================
// Recompute term_summaries for a class (all streams together)
// Called after writes/submits to keep positions fresh.
// ================================================================

async function recomputeTermSummariesForClass(env, schoolId, { termId, className, schoolSettings, nowIso }) {
  const students = await fetchRosterStudents(env, schoolId, className, "");
  const aLevel = isALevelClass(className);
  const rows = await d1All(
    env,
    `SELECT roster_id, subject, pct, grade, points, level, is_subsidiary
     FROM subject_summaries
     WHERE school_id = ? AND term_id = ? AND class_name = ?`,
    schoolId, termId, className
  );
  const byRoster = new Map();
  for (const r of rows) {
    if (!byRoster.has(r.roster_id)) byRoster.set(r.roster_id, []);
    byRoster.get(r.roster_id).push(r);
  }

  const scored = [];
  for (const s of students) {
    const list = byRoster.get(s.id) || [];
    if (list.length === 0) continue;
    let overallPct = null;
    let principalPoints = 0, subsidiaryPoints = 0;
    if (schoolSettings.scoringMode === "cbc") {
      // Overall pct not meaningful in CBC; still compute from levels if any.
      const levels = list.map((r) => r.level).filter((v) => Number.isFinite(v));
      overallPct = levels.length ? round1(levels.reduce((a, b) => a + b, 0) / levels.length) : null;
    } else {
      const pcts = list.map((r) => r.pct).filter((v) => v !== null && v !== undefined);
      overallPct = pcts.length ? round1(pcts.reduce((a, b) => a + b, 0) / pcts.length) : null;
      if (aLevel) {
        for (const r of list) {
          if (r.is_subsidiary) subsidiaryPoints += (r.points || 0);
          else principalPoints += (r.points || 0);
        }
      }
    }
    const totalPoints = aLevel ? (principalPoints + subsidiaryPoints) : null;
    if (overallPct === null && totalPoints === null) continue;
    scored.push({ rosterId: s.id, overallPct, principalPoints, subsidiaryPoints, totalPoints });
  }

  // Rank: A-level by total_points desc, O-level by overall_pct desc; ties share position.
  const keyOf = (r) => aLevel ? r.totalPoints : r.overallPct;
  const sortable = scored.filter((r) => keyOf(r) !== null);
  sortable.sort((a, b) => (keyOf(b) ?? -Infinity) - (keyOf(a) ?? -Infinity));
  const positions = new Map();
  let lastKey = null, lastPos = 0;
  sortable.forEach((r, i) => {
    const k = keyOf(r);
    if (k !== lastKey) { lastPos = i + 1; lastKey = k; }
    positions.set(r.rosterId, lastPos);
  });
  const outOf = sortable.length;

  const stmts = [];
  for (const r of scored) {
    const overallGrade = schoolSettings.scoringMode === "percentage"
      ? (applyGradingScale(r.overallPct, schoolSettings.gradingScale)?.grade || null)
      : null;
    const position = positions.get(r.rosterId) || null;
    stmts.push({
      sql: `INSERT INTO term_summaries (school_id, term_id, roster_id, class_name, stream, overall_pct, overall_grade, principal_points, subsidiary_points, total_points, position, out_of, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(school_id, term_id, roster_id) DO UPDATE SET
              class_name = excluded.class_name, stream = excluded.stream,
              overall_pct = excluded.overall_pct, overall_grade = excluded.overall_grade,
              principal_points = excluded.principal_points, subsidiary_points = excluded.subsidiary_points,
              total_points = excluded.total_points, position = excluded.position, out_of = excluded.out_of,
              updated_at = excluded.updated_at`,
      params: [
        schoolId, termId, r.rosterId, className, "", r.overallPct, overallGrade,
        aLevel ? r.principalPoints : null, aLevel ? r.subsidiaryPoints : null, aLevel ? r.totalPoints : null,
        position, outOf, nowIso,
      ],
    });
  }
  return stmts;
}

// ================================================================
// Excel: template
// ================================================================

async function handleDownloadTemplate(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";
  const subject = cleanText(url.searchParams.get("subject") || "", 120, "subject");
  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }

  const [types] = await Promise.all([listAssessmentTypes(env, schoolId)]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");
  const students = await fetchRosterStudents(env, schoolId, className, stream);

  const headers = ["rosterId", "Name", ...types.map((t) => t.name)];
  const rows = students.map((s) => [s.id, displayName(s.firstName, s.lastName), ...types.map(() => "")]);
  const hint = schoolSettings.scoringMode === "cbc"
    ? ["Enter levels (1, 2, 3) or ABS / EXC / N/A"]
    : ["Enter scores or ABS / EXC / N/A"];

  const ws = XLSX.utils.aoa_to_sheet([headers, hint, ...rows]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Marks");
  const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
  const filename = `marks_${className}${stream ? "_" + stream : ""}_${subject}.xlsx`.replace(/[^\w.-]+/g, "_");
  return binary(buf, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", {
    "Content-Disposition": `attachment; filename="${filename}"`,
  });
}

// ================================================================
// Excel: upload + confirm
// ================================================================

async function handleExcelUpload(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);

  let form;
  try { form = await request.formData(); }
  catch (err) { throw new HttpError(400, `Could not parse multipart body: ${err.message}`); }

  const file = form.get("file");
  if (!file || typeof file === "string" || typeof file.arrayBuffer !== "function") {
    throw new HttpError(400, '"file" is required (multipart, field name "file")');
  }
  const className = cleanText(str(form.get("class")), 60, "class");
  const stream = form.get("stream") ? cleanText(str(form.get("stream")), 60, "stream") : "";
  const subject = cleanText(str(form.get("subject")), 120, "subject");
  const requestedTermId = form.get("termId") ? str(form.get("termId")) : null;

  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }

  const termId = await resolveTermId(env, schoolId, requestedTermId);
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  const ext = (() => {
    const name = file.name || "";
    const dot = name.lastIndexOf(".");
    return dot >= 0 ? name.slice(dot).toLowerCase() : "";
  })();
  if (!ACCEPTED_EXTENSIONS.includes(ext)) {
    throw new HttpError(400, `Unsupported file type "${ext}". Accepted: ${ACCEPTED_EXTENSIONS.join(", ")}`);
  }
  const mime = String(file.type || "").toLowerCase();
  if (!ACCEPTED_MIME.has(mime)) throw new HttpError(400, `Unsupported content type "${file.type}"`);
  if (file.size > MAX_FILE_SIZE_MB * 1024 * 1024) throw new HttpError(400, `File too large (max ${MAX_FILE_SIZE_MB} MB)`);

  const bytes = await file.arrayBuffer();

  const scan = await scanFileForViruses(env, bytes, file.name || "upload");
  if (!scan.clean) throw new HttpError(400, "File failed the virus scan and was not processed.", { foundViruses: scan.foundViruses });

  let cloudinaryUrl = null;
  try {
    const uploaded = await uploadRawToCloudinary(
      env, bytes, file.name || "upload.xlsx",
      `edutraceug/marks-uploads/${schoolId}/${className}${stream ? "_" + stream : ""}/${subject}`
    );
    cloudinaryUrl = uploaded.url;
  } catch (err) {
    console.error("Cloudinary archive failed:", err && err.message);
  }

  const [types, term, termSettings] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getTermOrThrow(env, schoolId, termId),
    getTermAssessmentSettings(env, schoolId, termId),
  ]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");

  const { rows: parsedRows } = readRowsFromWorkbook(bytes, types);
  const maxScores = resolveMaxScores(termSettings, types);
  const maxLevel = schoolSettings.cbcLevels.length;

  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const byId = new Map(students.map((s) => [s.id, s]));
  const byKey = new Map();
  for (const s of students) {
    const key = s.nameKey || nameKey(s.firstName, s.lastName);
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push(s);
  }

  const matched = [];
  const ambiguous = [];
  const notFound = [];
  const cellErrors = [];

  for (const row of parsedRows) {
    let student = null;
    if (row.rosterId && byId.has(row.rosterId)) {
      student = byId.get(row.rosterId);
    } else {
      let fn = row.firstName, ln = row.lastName;
      if (!fn && !ln && row.fullName) {
        const split = splitFullName(row.fullName);
        fn = split.firstName;
        ln = split.lastName;
      }
      if (fn || ln) {
        const key = nameKey(fn, ln);
        const candidates = byKey.get(key) || [];
        if (candidates.length === 1) student = candidates[0];
        else if (candidates.length > 1) {
          ambiguous.push({
            excelRow: row.excelRow,
            name: displayName(fn, ln, row.fullName),
            candidates: candidates.map((c) => ({
              rosterId: c.id,
              firstName: c.firstName,
              lastName: c.lastName,
              class: c.class,
              stream: c.stream || "",
            })),
          });
          continue;
        }
      }
    }

    if (!student) {
      notFound.push({
        excelRow: row.excelRow,
        rosterId: row.rosterId,
        name: displayName(row.firstName, row.lastName, row.fullName),
      });
      continue;
    }

    const cleaned = {};
    let rowHasError = false;
    for (const type of types) {
      const raw = row.cells[type.id];
      if (raw === undefined || raw === "") continue;
      const parsed = interpretCell(raw, schoolSettings.scoringMode, maxLevel);
      if (!parsed) continue;
      if (parsed.error) { cellErrors.push({ excelRow: row.excelRow, typeName: type.name, error: parsed.error }); rowHasError = true; continue; }
      cleaned[type.id] = parsed;
    }
    if (rowHasError) continue;
    if (Object.keys(cleaned).length === 0) continue;
    matched.push({ student, cells: cleaned });
  }

  if (ambiguous.length > 0) {
    const state = await signState(env, {
      termId, className, stream, subject, schoolId, cloudinaryUrl,
      matched: matched.map((m) => ({ rosterId: m.student.id, cells: m.cells })),
      exp: Math.floor(Date.now() / 1000) + 900,
    });
    return {
      status: "needs_confirmation",
      summary: { matched: matched.length, ambiguous: ambiguous.length, notFound: notFound.length, cellErrors: cellErrors.length },
      ambiguous, notFound, cellErrors, state,
    };
  }

  const writeResult = await writeParsedMarks(env, schoolId, {
    termId, className, stream, subject, schoolSettings, types, termSettings, matched, cloudinaryUrl, uid: String(user.sub),
  });

  return {
    status: "ok", ...writeResult,
    summary: { matched: matched.length, ambiguous: 0, notFound: notFound.length, cellErrors: cellErrors.length },
    notFound, cellErrors, cloudinaryUrl,
  };
}

async function writeParsedMarks(env, schoolId, ctx) {
  const { termId, className, stream, subject, schoolSettings, termSettings, matched, cloudinaryUrl, uid } = ctx;
  const nowIso = new Date().toISOString();
  const stmts = [];
  let written = 0;
  const affectedRoster = new Set();
  const maxScores = resolveMaxScores(termSettings, ctx.types || []);

  for (const entry of matched) {
    const rosterId = entry.student.id;
    affectedRoster.add(rosterId);
    for (const [typeId, parsed] of Object.entries(entry.cells)) {
      const markId = await computeMarkId(termId, rosterId, subject, typeId);
      const cellMax = schoolSettings.scoringMode === "percentage" ? (maxScores[typeId] || parsed.score || null) : null;
      if (schoolSettings.scoringMode === "percentage" && parsed.status === "entered") {
        if (cellMax && parsed.score > cellMax) throw new HttpError(400, `Score ${parsed.score} exceeds maxScore ${cellMax} (row ${rosterId}, type ${typeId})`);
      }
      stmts.push({
        sql: `INSERT INTO marks (school_id, id, term_id, roster_id, class_name, stream, subject, type_id, scoring_mode, status, score, level, max_score, source, cloudinary_url, entered_by, entered_at, updated_by, updated_at)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'excel', ?, ?, ?, ?, ?)
              ON CONFLICT(school_id, id) DO UPDATE SET
                status = excluded.status, score = excluded.score, level = excluded.level, max_score = excluded.max_score,
                source = excluded.source, cloudinary_url = excluded.cloudinary_url,
                updated_by = excluded.updated_by, updated_at = excluded.updated_at`,
        params: [
          schoolId, markId, termId, rosterId, className, stream || "", subject, typeId, schoolSettings.scoringMode,
          parsed.status, parsed.score, parsed.level, cellMax, cloudinaryUrl || null, uid, nowIso, uid, nowIso,
        ],
      });
      written += 1;
    }
  }

  await runBatch(env, stmts);

  const types = ctx.types || await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const summaryStmts = await buildSummariesBatch(env, schoolId, {
    termId, className, stream, subject, rosterIds: Array.from(affectedRoster), schoolSettings, nowIso, uid, types,
  });
  summaryStmts.push(auditStmt(schoolId, {
    action: "marks.excel.write", by: uid, termId, class: className, stream, subject, count: written, cloudinaryUrl,
  }));
  await runBatch(env, summaryStmts);
  return { written, studentsAffected: matched.length };
}

async function handleExcelConfirm(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);
  const { state, resolutions } = body || {};
  if (!state) throw new HttpError(400, "state is required");
  if (!resolutions || typeof resolutions !== "object") throw new HttpError(400, "resolutions is required");

  const payload = await verifyState(env, state);
  if (String(payload.schoolId) !== String(schoolId)) throw new HttpError(400, "State token school mismatch");
  const { termId, className, stream, subject, matched, cloudinaryUrl } = payload;

  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  const [types, termSettings] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getTermAssessmentSettings(env, schoolId, termId),
  ]);
  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const byId = new Map(students.map((s) => [s.id, s]));

  const extraMatched = [];
  if (Array.isArray(body.ambiguousRows)) {
    for (const row of body.ambiguousRows) {
      const resolvedRosterId = resolutions[String(row.excelRow)];
      if (!resolvedRosterId) continue;
      const student = byId.get(resolvedRosterId);
      if (!student) throw new HttpError(400, `Resolved rosterId ${resolvedRosterId} not in class`);
      const cleaned = {};
      const cells = row.cells || {};
      for (const type of types) {
        const raw = cells[type.id];
        if (raw === undefined || raw === "") continue;
        const parsed = interpretCell(raw, schoolSettings.scoringMode, schoolSettings.cbcLevels.length);
        if (!parsed || parsed.error) continue;
        cleaned[type.id] = parsed;
      }
      if (Object.keys(cleaned).length > 0) extraMatched.push({ student, cells: cleaned });
    }
  }

  const allMatched = [
    ...matched.map((m) => ({ student: byId.get(m.rosterId), cells: m.cells })),
    ...extraMatched,
  ].filter((m) => m.student);
  if (allMatched.length === 0) throw new HttpError(400, "No rows to write after resolution");

  const result = await writeParsedMarks(env, schoolId, {
    termId, className, stream, subject, schoolSettings, types, termSettings, matched: allMatched, cloudinaryUrl, uid: String(user.sub),
  });
  return { status: "ok", ...result };
}

// ================================================================
// Mark sheets
// ================================================================

async function handleListSheets(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rows = await d1All(env, `SELECT * FROM mark_sheets WHERE school_id = ? AND term_id = ? LIMIT ?`, schoolId, termId, MAX_SHEETS);
  return {
    termId,
    sheets: rows.map((r) => ({
      id: r.id, class: r.class_name || "", stream: r.stream || "", subject: r.subject || "",
      status: r.status || "open",
      submittedAt: r.submitted_at || null,
      lockedAt: r.locked_at || null,
      updatedAt: r.updated_at || null,
    })),
  };
}

async function setSheetStatus(request, env, schoolId, targetStatus) {
  const allowedRoles = targetStatus === "submitted" ? ["schoolAdmin", "teacher"] : ["schoolAdmin"];
  const user = await requireUser(request, env, schoolId, allowedRoles);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const className = cleanText(body.class, 60, "class");
  const stream = body.stream ? cleanText(body.stream, 60, "stream") : "";
  const subject = cleanText(body.subject, 120, "subject");

  const schoolSettings = await getSchoolSettings(env, schoolId);
  if (user.role === "teacher") {
    await requireCanEnterMarks(env, schoolId, String(user.sub));
    await assertTeacherAssignment(env, schoolId, String(user.sub), className, stream, subject, schoolSettings);
  }

  const sheetId = await computeSheetId(termId, className, stream, subject);
  const existing = (await getSheet(env, schoolId, sheetId)) || {};
  const nowIso = new Date().toISOString();
  const uid = String(user.sub);

  const fields = {
    school_id: schoolId, id: sheetId,
    term_id: termId, class_name: className, stream: stream || "", subject,
    status: targetStatus, updated_at: nowIso, updated_by: uid,
  };
  if (targetStatus === "submitted") { fields.submitted_at = nowIso; fields.submitted_by = uid; }
  if (targetStatus === "locked")    { fields.locked_at = nowIso; fields.locked_by = uid; }
  if (targetStatus === "open")      { fields.reopened_at = nowIso; fields.reopened_by = uid; }

  const cols = Object.keys(fields);
  const stmts = [{
    sql: `INSERT INTO mark_sheets (${cols.join(", ")}) VALUES (${cols.map(() => "?").join(", ")})
          ON CONFLICT(school_id, id) DO UPDATE SET ${cols.filter((c) => c !== "id" && c !== "school_id").map((c) => `${c} = excluded.${c}`).join(", ")}`,
    params: cols.map((c) => fields[c]),
  }];

  // On submission, refresh term_summaries for the class so reports read fresh positions.
  if (targetStatus === "submitted") {
    const summaryStmts = await recomputeTermSummariesForClass(env, schoolId, { termId, className, schoolSettings, nowIso });
    for (const s of summaryStmts) stmts.push(s);
  }

  stmts.push(auditStmt(schoolId, {
    action: `sheet.${targetStatus}`, by: uid, termId, class: className, stream, subject,
    before: existing.status || null, after: targetStatus,
  }));
  await runBatch(env, stmts);
  return { status: "ok", sheetId, sheetStatus: targetStatus };
}

// ================================================================
// Config
// ================================================================

async function handleGetConfig(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const [term, types, schoolSettings, termSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getSchoolSettings(env, schoolId),
    getTermAssessmentSettings(env, schoolId, termId),
  ]);
  return {
    scoringMode: schoolSettings.scoringMode,
    cbcLevels: schoolSettings.cbcLevels,
    term: {
      id: termId, name: term.name || "", year: term.year || null,
      isCurrent: term.is_current === 1 || term.is_current === true,
      nextTermStart: term.next_term_start || null,
    },
    assessmentTypes: types,
    maxScores: resolveMaxScores(termSettings, types),
    weights: resolveWeights(termSettings, types),
    publishRank: termPublishRank(term),
    publishPosition: schoolSettings.publishPosition,
    gradingScale: schoolSettings.gradingScale,
    pointsBands: schoolSettings.pointsBands,
    subsidiaryPassMark: schoolSettings.subsidiaryPassMark,
  };
}

async function handleSetMaxScores(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  await getTermOrThrow(env, schoolId, termId);

  if (!body.maxScores || typeof body.maxScores !== "object" || Array.isArray(body.maxScores)) {
    throw new HttpError(400, "maxScores must be an object keyed by typeId");
  }
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const typeIds = new Set(types.map((t) => t.id));

  const nowIso = new Date().toISOString();
  const stmts = [];
  const merged = {};
  for (const [typeId, v] of Object.entries(body.maxScores)) {
    if (!typeIds.has(typeId)) throw new HttpError(400, `Unknown assessment type: ${typeId}`);
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0 || n > 1000) throw new HttpError(400, `maxScores.${typeId} must be 1-1000`);
    merged[typeId] = n;
    stmts.push({
      sql: `INSERT INTO term_assessment_settings (school_id, term_id, type_id, weight, max_score)
            VALUES (?, ?, ?, COALESCE((SELECT weight FROM term_assessment_settings WHERE school_id = ? AND term_id = ? AND type_id = ?), 0), ?)
            ON CONFLICT(school_id, term_id, type_id) DO UPDATE SET max_score = excluded.max_score`,
      params: [schoolId, termId, typeId, schoolId, termId, typeId, n],
    });
  }
  // Also persist weights if provided alongside.
  const weightsIn = body.weights && typeof body.weights === "object" ? body.weights : null;
  if (weightsIn) {
    for (const [typeId, v] of Object.entries(weightsIn)) {
      if (!typeIds.has(typeId)) throw new HttpError(400, `Unknown assessment type: ${typeId}`);
      const w = Number(v);
      if (!Number.isFinite(w) || w < 0) throw new HttpError(400, `weights.${typeId} must be >= 0`);
      stmts.push({
        sql: `INSERT INTO term_assessment_settings (school_id, term_id, type_id, weight, max_score)
              VALUES (?, ?, ?, ?, COALESCE((SELECT max_score FROM term_assessment_settings WHERE school_id = ? AND term_id = ? AND type_id = ?), NULL))
              ON CONFLICT(school_id, term_id, type_id) DO UPDATE SET weight = excluded.weight`,
        params: [schoolId, termId, typeId, w, schoolId, termId, typeId],
      });
    }
  }
  if (stmts.length === 0) throw new HttpError(400, "No maxScores or weights to update");
  stmts.push(auditStmt(schoolId, { action: "term.maxScores.set", by: String(admin.sub), termId, after: merged, weights: weightsIn || null }));
  await runBatch(env, stmts);
  return { termId, maxScores: merged, weights: weightsIn || undefined };
}

async function handleSetPublishRank(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  if (typeof body.publishRank !== "boolean") throw new HttpError(400, "publishRank must be boolean");
  const stmts = [{
    sql: `UPDATE terms SET publish_rank = ?, updated_at = ?, updated_by = ? WHERE school_id = ? AND id = ?`,
    params: [body.publishRank ? 1 : 0, new Date().toISOString(), String(admin.sub), schoolId, termId],
  }];
  stmts.push(auditStmt(schoolId, { action: "term.publishRank.set", by: String(admin.sub), termId, after: body.publishRank }));
  await runBatch(env, stmts);
  return { termId, publishRank: body.publishRank };
}

async function handleSetGradingScale(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  if (!Array.isArray(body.bands) || body.bands.length === 0) throw new HttpError(400, "bands must be a non-empty array");

  const clean = body.bands.map((b) => {
    if (!b || typeof b !== "object") throw new HttpError(400, "Invalid band");
    const min = Number(b.min);
    const max = Number(b.max);
    if (!Number.isFinite(min) || !Number.isFinite(max) || min < 0 || max > 100 || min >= max) {
      throw new HttpError(400, "Each band needs 0 <= min < max <= 100");
    }
    return {
      min, max,
      grade: cleanText(b.grade, 10, "band.grade"),
      points: Number.isFinite(Number(b.points)) ? Number(b.points) : null,
      descriptor: cleanOptionalText(b.descriptor, 60, "band.descriptor"),
    };
  });

  const nowIso = new Date().toISOString();
  const stmts = [{
    sql: `INSERT INTO school_settings (school_id, grading_scale_json, updated_at, updated_by)
          VALUES (?, ?, ?, ?)
          ON CONFLICT(school_id) DO UPDATE SET grading_scale_json = excluded.grading_scale_json, updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [schoolId, JSON.stringify({ bands: clean }), nowIso, String(admin.sub)],
  }];
  stmts.push(auditStmt(schoolId, { action: "settings.gradingScale.set", by: String(admin.sub), after: clean }));
  await runBatch(env, stmts);
  return { bands: clean };
}

async function handleSetPointsSettings(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);

  let bands = DEFAULT_POINTS_BANDS;
  if (body.bands !== undefined) {
    if (!Array.isArray(body.bands) || body.bands.length === 0) throw new HttpError(400, "bands must be a non-empty array");
    bands = body.bands.map((b) => {
      if (!b || typeof b !== "object") throw new HttpError(400, "Invalid band");
      const min = Number(b.min);
      if (!Number.isFinite(min) || min < 0 || min > 100) throw new HttpError(400, "Each band needs min 0-100");
      return { min, grade: cleanText(b.grade, 2, "band.grade").toUpperCase() };
    });
  }
  let passMark = DEFAULT_SUBSIDIARY_PASS_MARK;
  if (body.subsidiaryPassMark !== undefined) {
    const n = Number(body.subsidiaryPassMark);
    if (!Number.isFinite(n) || n < 0 || n > 100) throw new HttpError(400, "subsidiaryPassMark must be 0-100");
    passMark = n;
  }

  const nowIso = new Date().toISOString();
  const stmts = [{
    sql: `INSERT INTO school_settings (school_id, points_bands_json, subsidiary_pass_mark, updated_at, updated_by)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(school_id) DO UPDATE SET
            points_bands_json = excluded.points_bands_json,
            subsidiary_pass_mark = excluded.subsidiary_pass_mark,
            updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [schoolId, JSON.stringify(bands), passMark, nowIso, String(admin.sub)],
  }];
  stmts.push(auditStmt(schoolId, { action: "settings.points.set", by: String(admin.sub), after: { bands, subsidiaryPassMark: passMark } }));
  await runBatch(env, stmts);
  return { bands, subsidiaryPassMark: passMark };
}

// ================================================================
// Report assembly (reads summaries)
// ================================================================

async function loadSchool(env, schoolId) {
  const row = await d1Get(env, `SELECT * FROM schools WHERE id = ?`, schoolId);
  if (!row) return { name: "School" };
  return {
    name: row.name || "School",
    slug: row.slug || "",
    logoUrl: row.logo_url || "",
    location: row.location || "",
    phone: row.phone || "",
    email: row.email || "",
    themeColor: row.theme_color || "",
    headTeacher: row.head_teacher || "",
  };
}

async function buildStudentReport(env, schoolId, rosterId, termId, { includeRank, includeHold }) {
  const [term, types, schoolSettings, studentRow] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId, { includeArchived: true }),
    getSchoolSettings(env, schoolId),
    d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId),
  ]);
  if (!studentRow) throw new HttpError(404, "Student not found in roster");
  if (studentRow.record_type && studentRow.record_type !== "student") throw new HttpError(400, "rosterId is not a student");

  const student = {
    id: studentRow.id,
    firstName: studentRow.first_name || "",
    lastName: studentRow.last_name || "",
    class: studentRow.class_name || "",
    stream: studentRow.stream || "",
    gender: studentRow.gender || "",
    subjects: (() => { try { return JSON.parse(studentRow.subjects_json || "[]"); } catch { return []; } })(),
  };

  const subjectRows = await d1All(
    env,
    `SELECT * FROM subject_summaries WHERE school_id = ? AND term_id = ? AND roster_id = ?`,
    schoolId, termId, rosterId
  );
  const termRow = await d1Get(
    env, `SELECT * FROM term_summaries WHERE school_id = ? AND term_id = ? AND roster_id = ?`,
    schoolId, termId, rosterId
  );

  const aLevel = isALevelClass(student.class);
  const subjectSummaries = subjectRows
    .map((r) => ({
      name: r.subject,
      pct: r.pct,
      grade: r.grade,
      points: aLevel ? r.points : null,
      level: r.level,
      isSubsidiary: !!r.is_subsidiary,
    }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const overall = termRow
    ? {
        pct: termRow.overall_pct,
        grade: termRow.overall_grade,
        principalPoints: aLevel ? termRow.principal_points : null,
        subsidiaryPoints: aLevel ? termRow.subsidiary_points : null,
        totalPoints: aLevel ? termRow.total_points : null,
        position: termRow.position,
        outOf: termRow.out_of,
      }
    : { pct: null, grade: null, principalPoints: null, subsidiaryPoints: null, totalPoints: null, position: null, outOf: null };

  const hold = includeHold ? await getReportHold(env, schoolId, rosterId, termId) : null;
  const note = await getReportNote(env, schoolId, termId, rosterId);
  const publishPosition = includeRank && schoolSettings.publishPosition && termPublishRank(term);

  return {
    scoringMode: schoolSettings.scoringMode,
    cbcLevels: schoolSettings.cbcLevels,
    term: {
      id: termId, name: term.name || "", year: term.year || null,
      startDate: term.start_date || "", endDate: term.end_date || "",
      nextTermStart: term.next_term_start || null,
    },
    student: {
      rosterId,
      firstName: student.firstName,
      lastName: student.lastName,
      class: student.class,
      stream: student.stream,
      gender: student.gender,
    },
    subjects: subjectSummaries,
    overall,
    rollup: null,
    ...(aLevel && schoolSettings.scoringMode === "percentage" ? {
      points: {
        totalPoints: overall.totalPoints,
        principalPoints: overall.principalPoints,
        subsidiaryPoints: overall.subsidiaryPoints,
        principalCount: subjectSummaries.filter((s) => !s.isSubsidiary).length,
        subsidiaryCount: subjectSummaries.filter((s) => s.isSubsidiary).length,
      },
    } : {}),
    rank: publishPosition && overall.position ? { position: overall.position, total: overall.outOf } : null,
    hold,
    note,
    gradingScale: schoolSettings.scoringMode === "percentage" ? schoolSettings.gradingScale : null,
  };
}

async function getReportHold(env, schoolId, rosterId, termId) {
  const row = await d1Get(env, `SELECT * FROM report_holds WHERE school_id = ? AND roster_id = ?`, schoolId, rosterId);
  if (!row) return null;
  if (row.term_id && row.term_id !== termId) return null;
  return { rosterId: row.roster_id, termId: row.term_id || null, reason: row.reason || "", by: row.by_uid || null, at: row.at || null };
}

async function getReportNote(env, schoolId, termId, rosterId) {
  const row = await d1Get(env, `SELECT * FROM report_notes WHERE school_id = ? AND term_id = ? AND roster_id = ?`, schoolId, termId, rosterId);
  if (!row) return null;
  return {
    comment: row.comment_text || "", conduct: row.conduct || "",
    headComment: row.head_comment || "", additionalComment: row.additional_comment || "",
  };
}

// ================================================================
// Report handlers
// ================================================================

async function assertNotHeldForReader(env, schoolId, rosterId, termId, user) {
  if (user.role === "schoolAdmin" || user.role === "teacher") return;
  const hold = await getReportHold(env, schoolId, rosterId, termId);
  if (hold) throw new HttpError(403, `Report withheld: ${hold.reason || "contact the school"}`);
}

async function handleGetMyReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["student"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.rosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false });
}

async function handleGetChildReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["parent"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.childRosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a child's roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false });
}

async function handleGetStudentReport(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    await assertClassTeacher(env, schoolId, String(user.sub), student.class_name, student.stream || "");
  }
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true });
}

async function handleGetClassReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";

  if (user.role === "teacher") {
    await assertClassTeacher(env, schoolId, String(user.sub), className, stream);
  }

  const [term, students, schoolSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    fetchRosterStudents(env, schoolId, className, stream),
    getSchoolSettings(env, schoolId),
  ]);
  const aLevel = isALevelClass(className);
  const publishPosition = schoolSettings.publishPosition && termPublishRank(term);

  const rosterIds = students.map((s) => s.id);
  const chunk = (arr, n) => { const out = []; for (let i = 0; i < arr.length; i += n) out.push(arr.slice(i, i + n)); return out; };

  const summaries = [];
  const subjectRows = [];
  for (const ids of chunk(rosterIds, 90)) {
    const ph = ids.map(() => "?").join(",");
    const t = await d1All(
      env,
      `SELECT * FROM term_summaries WHERE school_id = ? AND term_id = ? AND roster_id IN (${ph})`,
      schoolId, termId, ...ids
    );
    for (const r of t) summaries.push(r);
    const s = await d1All(
      env,
      `SELECT roster_id, subject, pct, grade, points, level, is_subsidiary FROM subject_summaries
       WHERE school_id = ? AND term_id = ? AND roster_id IN (${ph})`,
      schoolId, termId, ...ids
    );
    for (const r of s) subjectRows.push(r);
  }
  const byRoster = new Map(summaries.map((r) => [r.roster_id, r]));
  const subjectsByRoster = new Map();
  for (const r of subjectRows) {
    if (!subjectsByRoster.has(r.roster_id)) subjectsByRoster.set(r.roster_id, []);
    subjectsByRoster.get(r.roster_id).push({
      name: r.subject, pct: r.pct, grade: r.grade, points: aLevel ? r.points : null,
      level: r.level, isSubsidiary: !!r.is_subsidiary,
    });
  }

  const rows = students.map((s) => {
    const t = byRoster.get(s.id);
    const subjects = (subjectsByRoster.get(s.id) || []).sort((a, b) => a.name.localeCompare(b.name));
    const out = {
      rosterId: s.id,
      firstName: s.firstName || "",
      lastName: s.lastName || "",
      stream: s.stream || "",
      subjects,
      overall: {
        pct: t ? t.overall_pct : null,
        grade: t ? t.overall_grade : null,
      },
      rollup: null,
    };
    if (aLevel && schoolSettings.scoringMode === "percentage") {
      out.points = {
        principalPoints: t ? t.principal_points : null,
        subsidiaryPoints: t ? t.subsidiary_points : null,
        totalPoints: t ? t.total_points : null,
      };
    }
    if (publishPosition && t && t.position) out.position = t.position;
    return out;
  });

  return {
    scoringMode: schoolSettings.scoringMode,
    cbcLevels: schoolSettings.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null },
    class: className, stream,
    publishRank: termPublishRank(term),
    publishPosition: schoolSettings.publishPosition,
    students: rows,
  };
}

// ================================================================
// Report card HTML
// ================================================================

const REPORT_CARD_CSS = `
  :root{--primary:{{PRIMARY}};--text-dark:#1A1A1A;--text-muted:#555;--border:#E5E5E5;--bg-light:#F8F9FA;--points:#1a7f37;--points-bg:#e8f6ec}
  *{margin:0;padding:0;box-sizing:border-box;border-radius:0!important}
  body{font-family:'Plus Jakarta Sans',Arial,sans-serif;background:#fff;color:var(--text-dark);padding:24px;line-height:1.5}
  .card{max-width:820px;margin:0 auto;border:2px solid var(--primary)}
  .header{display:flex;align-items:center;gap:20px;padding:22px 26px;border-bottom:3px solid var(--primary)}
  .logo{width:78px;height:78px;object-fit:contain;flex-shrink:0}
  .logo-fallback{width:78px;height:78px;background:var(--primary);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:1.6rem;flex-shrink:0}
  .school-info h1{font-size:1.5rem;font-weight:800;color:var(--primary);line-height:1.2}
  .school-info p{font-size:.8rem;color:var(--text-muted);margin-top:2px}
  .motto{font-size:.8rem;color:var(--text-muted);font-style:italic;margin-top:4px}
  .title-bar{background:var(--primary);color:#fff;padding:12px 26px;text-align:center}
  .title-bar h2{font-size:1rem;font-weight:700;text-transform:uppercase;letter-spacing:1.5px}
  .student-info{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;padding:18px 26px;background:var(--bg-light);border-bottom:1px solid var(--border)}
  .student-info div{font-size:.85rem}
  .student-info label{display:block;font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);font-weight:700;margin-bottom:2px}
  table{width:100%;border-collapse:collapse;table-layout:fixed}
  th,td{padding:9px 8px;border-bottom:1px solid var(--border);font-size:.85rem;text-align:center;word-wrap:break-word;overflow-wrap:break-word}
  th{background:var(--primary);color:#fff;font-weight:700;text-transform:uppercase;font-size:.7rem;letter-spacing:.5px;padding:10px 6px}
  th.subject,td.subject{text-align:left;font-weight:600;width:40%}
  th.total,td.total{width:20%}
  th.grade,td.grade{width:20%}
  th.points,td.points{width:20%}
  td.total{font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}
  td.grade,td.points{font-weight:800}
  .sub-tag{font-size:.65rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;margin-left:4px}
  .points-strip{display:grid;grid-template-columns:repeat(4,1fr);background:var(--points-bg);border-top:2px solid var(--points);border-bottom:2px solid var(--points)}
  .points-strip>div{padding:14px 16px;text-align:center;border-right:1px solid #c7e5d1}
  .points-strip>div:last-child{border-right:none}
  .points-strip label{display:block;font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;color:var(--points);font-weight:700;margin-bottom:4px}
  .points-strip strong{display:block;font-size:1.35rem;font-weight:800;color:var(--points)}
  .points-strip small{display:block;font-size:.75rem;color:#2a6b3d;margin-top:2px}
  .summary{display:grid;grid-template-columns:repeat(3,1fr);border-top:2px solid var(--primary)}
  .summary>div{padding:16px 20px;text-align:center;border-right:1px solid var(--border)}
  .summary>div:last-child{border-right:none}
  .summary label{display:block;font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700;margin-bottom:4px}
  .summary strong{display:block;font-size:1.35rem;font-weight:800;color:var(--primary)}
  .summary small{display:block;font-size:.75rem;color:var(--text-muted);margin-top:2px}
  .legend{padding:10px 26px;background:var(--bg-light);font-size:.75rem;color:var(--text-muted);border-top:1px solid var(--border)}
  .legend b{color:var(--text-dark)}
  .comment-box{padding:20px 26px;border-top:1px solid var(--border)}
  .comment-box h3{font-size:.75rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);margin-bottom:10px;font-weight:700}
  .comment-box p{font-size:.9rem;min-height:44px}
  .signature{display:flex;justify-content:space-between;gap:30px;padding:26px;border-top:1px solid var(--border)}
  .signature>div{flex:1}
  .signature .line{border-bottom:1px solid #999;height:38px;margin-bottom:6px}
  .signature label{font-size:.7rem;color:var(--text-muted);text-transform:uppercase;letter-spacing:.5px;font-weight:700}
  .footer{background:#111;color:#ccc;padding:12px 26px;text-align:center;font-size:.75rem}
  .footer a{color:#9cc;text-decoration:none}
  .no-print{padding:16px;text-align:center}
  .no-print button{padding:10px 22px;background:var(--primary);color:#fff;border:none;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;font-family:inherit}
  .stamp{display:flex;justify-content:flex-end;padding:12px 26px}
  .stamp img{max-height:90px;max-width:120px;object-fit:contain;opacity:.85}
  @page{size:A4;margin:12mm}
  @media print{
    body{padding:0}
    .card{border:none}
    .no-print{display:none}
    thead{display:table-header-group}
    tr{break-inside:avoid;page-break-inside:avoid}
  }
  @media(max-width:640px){.student-info{grid-template-columns:repeat(2,1fr)}.summary{grid-template-columns:1fr}.points-strip{grid-template-columns:repeat(2,1fr)}th,td{padding:6px 4px;font-size:.75rem}}
`;

function reportCardHead(title, themeColor) {
  return `<!DOCTYPE html><html lang="en-UG"><head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${escapeHtml(title)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>${REPORT_CARD_CSS.replace("{{PRIMARY}}", themeColor)}</style></head><body>`;
}

function cardHeader(school, schoolSettings) {
  const schoolName = escapeHtml(school.name || "School");
  const logoUrl = school.logoUrl ? escapeHtml(school.logoUrl) : "";
  const motto = schoolSettings && schoolSettings.cardMotto ? `<p class="motto">${escapeHtml(schoolSettings.cardMotto)}</p>` : "";
  return `<div class="header">
    ${logoUrl ? `<img class="logo" src="${logoUrl}" alt="">` : `<div class="logo-fallback">${escapeHtml((schoolName[0] || "S").toUpperCase())}</div>`}
    <div class="school-info"><h1>${schoolName}</h1>
      ${school.location ? `<p>${escapeHtml(school.location)}</p>` : ""}
      ${school.phone ? `<p>Tel: ${escapeHtml(school.phone)}</p>` : ""}
      ${school.email ? `<p>Email: ${escapeHtml(school.email)}</p>` : ""}
      ${motto}
    </div>
  </div>`;
}

function cardFooter(school, schoolSettings) {
  const schoolName = escapeHtml(school.name || "School");
  const website = schoolSettings && schoolSettings.cardWebsite ? `<a href="${escapeHtml(schoolSettings.cardWebsite)}">${escapeHtml(schoolSettings.cardWebsite)}</a>` : "";
  return `<div class="footer">© ${new Date().getFullYear()} ${schoolName}.${website ? " " + website : ""} Powered by Edutrace.</div>`;
}

function cardComments(note) {
  return `
  <div class="comment-box"><h3>Class Teacher's Comment</h3>${note.conduct ? `<p><b>Conduct:</b> ${escapeHtml(note.conduct)}</p>` : ""}<p>${escapeHtml(note.comment || "") || "&nbsp;"}</p></div>
  <div class="comment-box"><h3>Head Teacher's Comment</h3><p>${escapeHtml(note.headComment || "") || "&nbsp;"}</p></div>
  <div class="comment-box"><h3>Additional Comment</h3><p>${escapeHtml(note.additionalComment || "") || "&nbsp;"}</p></div>
  <div class="signature">
    <div><div class="line"></div><label>Class Teacher</label></div>
    <div><div class="line"></div><label>Head Teacher</label></div>
    <div><div class="line"></div><label>Date</label></div>
  </div>`;
}

function renderPercentageReportHtml({ report, school, schoolSettings }) {
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";
  const t = report.term, s = report.student;
  const isALevel = !!(report.points);
  const publishedRank = report.rank && report.rank.position;

  const rowsHtml = report.subjects.map((sub) => {
    const pct = sub.pct !== null && sub.pct !== undefined ? Number(sub.pct).toFixed(1) : "—";
    const grade = sub.grade ? escapeHtml(sub.grade) : "—";
    const points = isALevel ? (sub.points === null || sub.points === undefined ? "—" : String(sub.points)) : "—";
    const subTag = sub.isSubsidiary ? `<span class="sub-tag">sub</span>` : "";
    return `<tr>
      <td class="subject">${escapeHtml(sub.name)}${subTag}</td>
      <td class="total">${pct}</td>
      <td class="grade">${grade}</td>
      ${isALevel ? `<td class="points">${points}</td>` : ""}
    </tr>`;
  }).join("");

  const overallPct = report.overall.pct !== null && report.overall.pct !== undefined ? Number(report.overall.pct).toFixed(1) : "—";
  const overallGrade = report.overall.grade ? escapeHtml(report.overall.grade) : "—";

  let pointsHtml = "";
  if (isALevel) {
    pointsHtml = `<div class="points-strip">
      <div><label>Total Points</label><strong>${report.points.totalPoints ?? "—"}</strong><small>principal + subsidiary</small></div>
      <div><label>Principal</label><strong>${report.points.principalPoints ?? "—"}</strong><small>${report.points.principalCount} subject(s)</small></div>
      <div><label>Subsidiary</label><strong>${report.points.subsidiaryPoints ?? "—"}</strong><small>${report.points.subsidiaryCount} subject(s)</small></div>
      <div><label>Position</label><strong>${publishedRank || "—"}</strong><small>of ${report.rank?.total ?? "—"}</small></div>
    </div>`;
  }

  let summaryHtml;
  if (isALevel) {
    summaryHtml = `<div class="summary">
      <div><label>Overall</label><strong>${overallPct}%</strong></div>
      <div><label>Grade</label><strong>${overallGrade}</strong></div>
      <div><label>Subjects</label><strong>${report.subjects.length}</strong></div>
    </div>`;
  } else if (publishedRank) {
    summaryHtml = `<div class="summary">
      <div><label>Overall</label><strong>${overallPct}%</strong></div>
      <div><label>Grade</label><strong>${overallGrade}</strong></div>
      <div><label>Position</label><strong>${publishedRank} of ${report.rank.total}</strong></div>
    </div>`;
  } else {
    summaryHtml = `<div class="summary">
      <div><label>Overall</label><strong>${overallPct}%</strong></div>
      <div><label>Grade</label><strong>${overallGrade}</strong></div>
      <div><label>Subjects</label><strong>${report.subjects.length}</strong></div>
    </div>`;
  }

  const scale = report.gradingScale || { bands: DEFAULT_GRADING_SCALE.bands };
  const legendHtml = scale.bands
    .slice().sort((a, b) => b.min - a.min)
    .map((b) => {
      const range = b.max === 100 ? `${b.min}+` : `${b.min}–${b.max - 1}`;
      return `<b>${escapeHtml(b.grade)}</b> ${range}%${b.descriptor ? ` (${escapeHtml(b.descriptor)})` : ""}`;
    })
    .join(" · ");

  const note = report.note || {};
  const nextTerm = t.nextTermStart ? `<p style="padding:10px 26px;font-size:.8rem;color:#555"><b>Next term begins:</b> ${escapeHtml(String(t.nextTermStart))}</p>` : "";
  const stampHtml = schoolSettings && schoolSettings.stampUrl ? `<div class="stamp"><img src="${escapeHtml(schoolSettings.stampUrl)}" alt=""></div>` : "";

  return `${reportCardHead(`Report Card — ${s.firstName} ${s.lastName} — ${t.name}`, themeColor)}
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  ${cardHeader(school, schoolSettings)}
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <colgroup>
      <col style="width:40%"><col style="width:20%"><col style="width:20%">
      ${isALevel ? '<col style="width:20%">' : ""}
    </colgroup>
    <thead><tr>
      <th class="subject">Subject</th>
      <th class="total">Total %</th>
      <th class="grade">Grade</th>
      ${isALevel ? '<th class="points">Points</th>' : ""}
    </tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="${isALevel ? 4 : 3}" class="grade">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  ${pointsHtml}
  <div class="legend"><b>Grade legend:</b> ${legendHtml}</div>
  ${summaryHtml}
  ${nextTerm}
  ${cardComments(note)}
  ${stampHtml}
  ${cardFooter(school, schoolSettings)}
</div></body></html>`;
}

function renderCbcReportHtml({ report, school, schoolSettings }) {
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";
  const t = report.term, s = report.student;
  const levels = (report.cbcLevels || []).slice().sort((a, b) => a.level - b.level);
  const labelFor = (lv) => {
    const m = levels.find((l) => l.level === lv);
    return m ? m.label : (lv === null || lv === undefined ? "—" : String(lv));
  };

  const rowsHtml = report.subjects.map((sub) => {
    return `<tr>
      <td class="subject">${escapeHtml(sub.name)}</td>
      <td class="grade">${escapeHtml(labelFor(sub.level))}</td>
    </tr>`;
  }).join("");

  const legendHtml = levels.map((l) => `<b>${l.level}</b> = ${escapeHtml(l.label)}`).join(" · ");
  const note = report.note || {};
  const nextTerm = t.nextTermStart ? `<p style="padding:10px 26px;font-size:.8rem;color:#555"><b>Next term begins:</b> ${escapeHtml(String(t.nextTermStart))}</p>` : "";
  const stampHtml = schoolSettings && schoolSettings.stampUrl ? `<div class="stamp"><img src="${escapeHtml(schoolSettings.stampUrl)}" alt=""></div>` : "";

  return `${reportCardHead(`Report Card — ${s.firstName} ${s.lastName} — ${t.name}`, themeColor)}
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  ${cardHeader(school, schoolSettings)}
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <colgroup><col style="width:60%"><col style="width:40%"></colgroup>
    <thead><tr><th class="subject">Subject</th><th class="grade">Level</th></tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="2" class="grade">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  <div class="legend"><b>Level legend:</b> ${legendHtml}</div>
  ${nextTerm}
  ${cardComments(note)}
  ${stampHtml}
  ${cardFooter(school, schoolSettings)}
</div></body></html>`;
}

function renderReportCardHtml({ report, school, schoolSettings }) {
  return report.scoringMode === "cbc"
    ? renderCbcReportHtml({ report, school, schoolSettings })
    : renderPercentageReportHtml({ report, school, schoolSettings });
}

async function handleGetReportCard(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    await assertClassTeacher(env, schoolId, String(user.sub), student.class_name, student.stream || "");
  }
  const [report, school, schoolSettings] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true }),
    loadSchool(env, schoolId),
    getSchoolSettings(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school, schoolSettings }));
}

async function handleGetMyReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["student"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.rosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school, schoolSettings] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    loadSchool(env, schoolId),
    getSchoolSettings(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school, schoolSettings }));
}

async function handleGetChildReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["parent"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.childRosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a child's roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school, schoolSettings] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    loadSchool(env, schoolId),
    getSchoolSettings(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school, schoolSettings }));
}

// ================================================================
// Notes
// ================================================================

async function handleSetNote(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const rosterId = cleanText(body.rosterId, 40, "rosterId");

  const student = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    await assertClassTeacher(env, schoolId, String(user.sub), student.class_name, student.stream || "");
  }

  const hasComment = Object.prototype.hasOwnProperty.call(body, "comment");
  const hasConduct = Object.prototype.hasOwnProperty.call(body, "conduct");
  const hasHead = Object.prototype.hasOwnProperty.call(body, "headComment");
  if (hasHead && user.role !== "schoolAdmin") throw new HttpError(403, "Only the school admin can write the headteacher comment");

  const fields = {
    school_id: schoolId, term_id: termId, roster_id: rosterId,
    class_name: student.class_name || "", stream: student.stream || "",
    updated_at: new Date().toISOString(), updated_by: String(user.sub),
  };
  if (hasComment) fields.comment_text = cleanOptionalText(body.comment, 2000, "comment");
  if (hasConduct) fields.conduct = cleanOptionalText(body.conduct, 200, "conduct");
  if (hasHead) fields.head_comment = cleanOptionalText(body.headComment, 2000, "headComment");
  if (Object.prototype.hasOwnProperty.call(body, "additionalComment")) {
    fields.additional_comment = cleanOptionalText(body.additionalComment, 2000, "additionalComment");
  }

  const existing = await d1Get(env, `SELECT * FROM report_notes WHERE school_id = ? AND term_id = ? AND roster_id = ?`, schoolId, termId, rosterId);
  const merged = {
    comment: hasComment ? fields.comment_text : (existing?.comment_text || ""),
    conduct: hasConduct ? fields.conduct : (existing?.conduct || ""),
    head_comment: hasHead ? fields.head_comment : (existing?.head_comment || ""),
    additional_comment: Object.prototype.hasOwnProperty.call(body, "additionalComment")
      ? fields.additional_comment
      : (existing?.additional_comment || ""),
  };

  const stmts = [{
    sql: `INSERT INTO report_notes (school_id, term_id, roster_id, class_name, stream, conduct, comment_text, head_comment, additional_comment, updated_at, updated_by)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(school_id, term_id, roster_id) DO UPDATE SET
            class_name = excluded.class_name, stream = excluded.stream,
            conduct = excluded.conduct, comment_text = excluded.comment_text,
            head_comment = excluded.head_comment, additional_comment = excluded.additional_comment,
            updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [
      schoolId, termId, rosterId, fields.class_name, fields.stream,
      merged.conduct, merged.comment, merged.head_comment, merged.additional_comment,
      fields.updated_at, fields.updated_by,
    ],
  }];
  stmts.push(auditStmt(schoolId, { action: "report.note.set", by: String(user.sub), termId, rosterId }));
  await runBatch(env, stmts);

  return {
    status: "ok",
    note: {
      termId, rosterId,
      class: fields.class_name, stream: fields.stream,
      comment: merged.comment, conduct: merged.conduct,
      headComment: merged.head_comment, additionalComment: merged.additional_comment,
      updatedAt: fields.updated_at, updatedBy: fields.updated_by,
    },
  };
}

// ================================================================
// Holds
// ================================================================

async function handleListHolds(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const termId = url.searchParams.get("termId");
  let rows;
  if (termId) rows = await d1All(env, `SELECT * FROM report_holds WHERE school_id = ? AND term_id = ? LIMIT ?`, schoolId, termId, MAX_HOLDS);
  else rows = await d1All(env, `SELECT * FROM report_holds WHERE school_id = ? LIMIT ?`, schoolId, MAX_HOLDS);
  return {
    holds: rows.map((r) => ({
      rosterId: r.roster_id, termId: r.term_id || null,
      reason: r.reason || "", by: r.by_uid || null, at: r.at || null,
    })),
  };
}

async function handleCreateHold(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const rosterId = cleanText(body.rosterId, 40, "rosterId");
  const reason = cleanOptionalText(body.reason, 300, "reason");
  const owns = await d1Get(env, `SELECT id FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!owns) throw new HttpError(404, "Student not found");
  const at = new Date().toISOString();

  const stmts = [{
    sql: `INSERT INTO report_holds (school_id, roster_id, term_id, reason, by_uid, at)
          VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT(school_id, roster_id) DO UPDATE SET
            term_id = excluded.term_id, reason = excluded.reason, by_uid = excluded.by_uid, at = excluded.at`,
    params: [schoolId, rosterId, termId, reason, String(admin.sub), at],
  }];
  stmts.push(auditStmt(schoolId, { action: "report.hold.create", by: String(admin.sub), termId, rosterId, reason }));
  await runBatch(env, stmts);
  return { status: "ok", hold: { termId, rosterId, reason, by: String(admin.sub), at } };
}

async function handleDeleteHold(request, env, schoolId, rosterId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const owns = await d1Get(env, `SELECT id FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!owns) throw new HttpError(404, "Student not found");
  const stmts = [
    { sql: `DELETE FROM report_holds WHERE school_id = ? AND roster_id = ?`, params: [schoolId, rosterId] },
    auditStmt(schoolId, { action: "report.hold.delete", by: String(admin.sub), rosterId }),
  ];
  await runBatch(env, stmts);
  return { status: "ok" };
}

// ================================================================
// Audit read
// ================================================================

async function handleGetAuditLog(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const requested = parseInt(url.searchParams.get("limit") || "50", 10);
  const limit = Math.min(Math.max(Number.isInteger(requested) ? requested : 50, 1), 100);
  const rows = await d1All(env, `SELECT * FROM audit_log WHERE school_id = ? ORDER BY at DESC LIMIT ?`, schoolId, limit);
  return {
    entries: rows.map((r) => {
      let details = {};
      try { details = JSON.parse(r.details_json || "{}"); } catch { details = {}; }
      return {
        id: r.rowid || null,
        action: r.action,
        by: r.by_uid,
        at: r.at,
        target: r.target,
        ...details,
      };
    }),
  };
}

// ================================================================
// One-time summary rebuild
// ================================================================

async function handleRebuildSummaries(request, env, schoolId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const uid = String(user.sub);
  const nowIso = new Date().toISOString();
  const schoolSettings = await getSchoolSettings(env, schoolId);
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });

  const groups = await d1All(
    env,
    `SELECT DISTINCT term_id, class_name, stream, subject FROM marks WHERE school_id = ?`,
    schoolId
  );
  let sheets = 0;
  const classKeys = new Map();
  for (const g of groups) {
    const rs = await d1All(
      env,
      `SELECT DISTINCT roster_id FROM marks WHERE school_id = ? AND term_id = ? AND class_name = ? AND subject = ?`,
      schoolId, g.term_id, g.class_name, g.subject
    );
    const stmts = await buildSummariesBatch(env, schoolId, {
      termId: g.term_id, className: g.class_name, stream: g.stream || "", subject: g.subject,
      rosterIds: rs.map((r) => r.roster_id), schoolSettings, nowIso, uid, types,
    });
    await runBatch(env, stmts);
    classKeys.set(`${g.term_id}|${g.class_name}`, { termId: g.term_id, className: g.class_name });
    sheets += 1;
  }
  for (const k of classKeys.values()) {
    const stmts = await recomputeTermSummariesForClass(env, schoolId, { termId: k.termId, className: k.className, schoolSettings, nowIso });
    await runBatch(env, stmts);
  }
  return { status: "ok", subjectGroups: sheets, classes: classKeys.size };
}

// ================================================================
// Router
// ================================================================

async function handle(request, env) {
  const url = new URL(request.url);
  const path = url.pathname;
  const method = request.method;

  if (method === "OPTIONS") return new Response(null, { status: 204, headers: CORS_HEADERS });
  if (method === "GET" && path === "/health") return json({ status: "ok", service: "edutraceug-marks-worker" });

  let m;
  const d = decodeURIComponent;
  const S = "([^/]+)";

  try {
    // Grid
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/grid$`)))) return json(await handleGetGrid(request, env, d(m[1]), url));
    if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/grid$`)))) return json(await handlePutGrid(request, env, d(m[1])));
    if (method === "DELETE" && (m = path.match(new RegExp(`^/schools/${S}/marks/${S}$`)))) return json(await handleDeleteMark(request, env, d(m[1]), d(m[2])));

    // Excel
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/marks/template$`)))) return handleDownloadTemplate(request, env, d(m[1]), url);
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/marks/upload$`)))) return json(await handleExcelUpload(request, env, d(m[1])));
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/marks/upload/confirm$`)))) return json(await handleExcelConfirm(request, env, d(m[1])));

    // Sheets
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/sheets$`)))) return json(await handleListSheets(request, env, d(m[1]), url));
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/submit$`)))) return json(await setSheetStatus(request, env, d(m[1]), "submitted"));
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/reopen$`)))) return json(await setSheetStatus(request, env, d(m[1]), "open"));
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/sheets/lock$`)))) return json(await setSheetStatus(request, env, d(m[1]), "locked"));

    // Config
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/config$`)))) return json(await handleGetConfig(request, env, d(m[1]), url));
    if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/max-scores$`)))) return json(await handleSetMaxScores(request, env, d(m[1])));
    if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/publish-rank$`)))) return json(await handleSetPublishRank(request, env, d(m[1])));
    if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/grading-scale$`)))) return json(await handleSetGradingScale(request, env, d(m[1])));
    if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/config/points-settings$`)))) return json(await handleSetPointsSettings(request, env, d(m[1])));

    // Reports
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/me$`)))) return json(await handleGetMyReport(request, env, d(m[1]), url));
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/me/card$`)))) return handleGetMyReportCard(request, env, d(m[1]), url);
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/child$`)))) return json(await handleGetChildReport(request, env, d(m[1]), url));
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/child/card$`)))) return handleGetChildReportCard(request, env, d(m[1]), url);
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/class$`)))) return json(await handleGetClassReport(request, env, d(m[1]), url));
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/student/${S}$`)))) return json(await handleGetStudentReport(request, env, d(m[1]), d(m[2]), url));
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/student/${S}/card$`)))) return handleGetReportCard(request, env, d(m[1]), d(m[2]), url);

    // Notes
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/notes$`)))) return json(await handleSetNote(request, env, d(m[1])));

    // Holds
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) return json(await handleListHolds(request, env, d(m[1]), url));
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) return json(await handleCreateHold(request, env, d(m[1])));
    if (method === "DELETE" && (m = path.match(new RegExp(`^/schools/${S}/holds/${S}$`)))) return json(await handleDeleteHold(request, env, d(m[1]), d(m[2])));

    // Rebuild
    if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/rebuild-summaries$`)))) return json(await handleRebuildSummaries(request, env, d(m[1])));

    // Audit
    if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/audit$`)))) return json(await handleGetAuditLog(request, env, d(m[1]), url));

    return json({ error: "Not found" }, 404);
  } catch (err) {
    if (err instanceof HttpError) return json({ error: err.message, ...err.extra }, err.status);
    console.error("Unhandled error:", err && err.stack ? err.stack : err);
    return json({ error: "Internal server error" }, 500);
  }
}

export default {
  async fetch(request, env) {
    try {
      return await handle(request, env);
    } catch (err) {
      console.error("Fatal:", err);
      return json({ error: "Internal server error" }, 500);
    }
  },
};
