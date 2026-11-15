// edutraceug-marks-worker (D1)
//
// Records assessment marks, computes term reports, generates printable report cards.
// Uses Cloudflare D1 (binding: DB) for all storage. Firebase ID token verification unchanged.
//
// D1 tables owned by this worker:
//   marks, mark_sheets, report_holds, report_notes, grading_scale,
//   audit_log, points_settings
//
// D1 tables owned by other workers (read-only here):
//   roster (roster-worker)             — record_type, class_name, stream, names, subjects, gender, name_key
//   terms, assessment_types,
//   scoring_settings, teachers (teachers-worker)
//
// Secrets:
//   ACCOUNT_SERVICE_FIREBASE  (firebase project id + JWKS verification)
//   CLOUDINARY_URL
//   CLOUDMERSIVE_API_KEY

import * as XLSX from "xlsx";

// ================================================================
// Constants
// ================================================================

const DEFAULT_PROJECT_ID = "edutrace-ug";

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

// A-level defaults (Uganda)
const DEFAULT_POINTS_SETTINGS = {
  bands: [
    { min: 80, grade: "A" },
    { min: 70, grade: "B" },
    { min: 60, grade: "C" },
    { min: 50, grade: "D" },
    { min: 40, grade: "E" },
    { min: 30, grade: "O" },
    { min: 0,  grade: "F" },
  ],
  subsidiaryPassMark: 50,
};

const PRINCIPAL_POINTS = { A: 6, B: 5, C: 4, D: 3, E: 2, O: 1, F: 0 };

const MAX_FILE_SIZE_MB = 25;
const BATCH_SIZE = 100; // D1 statement batch cap
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
// Names (identical to roster-worker)
// ================================================================

// Order-independent name key. Accepts either (firstName, lastName) or a
// single combined string; passing one argument still works.
function nameKey(firstName, lastName) {
  const norm = (v) =>
    String(v ?? "").toLowerCase().replace(/[^\p{L}\p{N}\s'-]+/gu, " ").replace(/\s+/g, " ").trim();
  const words = `${norm(firstName)} ${norm(lastName)}`.split(" ").filter(Boolean);
  words.sort();
  return words.join(" ");
}

// Split a single "Name" column into first/last. Everything except the final
// token becomes the first name so middle names are preserved.
function splitFullName(full) {
  const parts = String(full || "").trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return { firstName: "", lastName: "" };
  if (parts.length === 1) return { firstName: parts[0], lastName: "" };
  return {
    firstName: parts.slice(0, -1).join(" "),
    lastName: parts[parts.length - 1],
  };
}

// Order the display name as "John Doe", never "DOE JOHN".
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
// Firebase ID token verification (UNCHANGED)
// ================================================================

let _saCache = { raw: null, parsed: null };
function getServiceAccount(env) {
  const raw = env["ACCOUNT_SERVICE_FIREBASE"] || env["ACCOUNT-SERVICE-FIREBASE"];
  if (!raw) throw new HttpError(500, "Server misconfigured: ACCOUNT_SERVICE_FIREBASE");
  if (_saCache.raw === raw && _saCache.parsed) return _saCache.parsed;
  const parsed = typeof raw === "object" ? raw : JSON.parse(raw);
  _saCache = { raw, parsed };
  return parsed;
}
function getProjectId(env) { return getServiceAccount(env).project_id || DEFAULT_PROJECT_ID; }

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
  const payload = await verifyIdToken(m[1], getProjectId(env));
  if (!allowedRoles.includes(payload.role)) throw new HttpError(403, "Not allowed for your role");
  if (String(payload.schoolId) !== String(schoolId)) throw new HttpError(403, "Not a member of this school");
  return payload;
}

// ================================================================
// In-worker cache (a few minutes for read-heavy lookups)
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
const CACHE_TTL_MS = 3 * 60_000; // few minutes

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
// Schema (created on first use; also provided as migration SQL)
// ================================================================

let _schemaReady = false;
async function ensureSchema(env) {
  if (_schemaReady) return;
  const conn = db(env);
  const stmts = [
    `CREATE TABLE IF NOT EXISTS marks (
      id TEXT PRIMARY KEY,
      term_id TEXT NOT NULL,
      roster_id TEXT NOT NULL,
      class TEXT NOT NULL,
      stream TEXT,
      subject TEXT NOT NULL,
      type_id TEXT NOT NULL,
      scoring_mode TEXT NOT NULL,
      status TEXT NOT NULL,
      score REAL,
      level INTEGER,
      max_score REAL,
      source TEXT,
      cloudinary_url TEXT,
      entered_by TEXT NOT NULL,
      entered_at TEXT NOT NULL,
      updated_by TEXT NOT NULL,
      updated_at TEXT NOT NULL
    )`,
    `CREATE INDEX IF NOT EXISTS idx_marks_term_class ON marks(term_id, class, stream)`,
    `CREATE INDEX IF NOT EXISTS idx_marks_term_roster ON marks(term_id, roster_id)`,
    `CREATE INDEX IF NOT EXISTS idx_marks_grid ON marks(term_id, class, stream, subject, type_id)`,

    `CREATE TABLE IF NOT EXISTS mark_sheets (
      id TEXT PRIMARY KEY,
      term_id TEXT NOT NULL,
      class TEXT NOT NULL,
      stream TEXT,
      subject TEXT NOT NULL,
      status TEXT NOT NULL,
      submitted_at TEXT,
      submitted_by TEXT,
      locked_at TEXT,
      locked_by TEXT,
      reopened_at TEXT,
      reopened_by TEXT,
      updated_at TEXT NOT NULL,
      updated_by TEXT NOT NULL
    )`,
    `CREATE INDEX IF NOT EXISTS idx_sheets_term ON mark_sheets(term_id)`,

    `CREATE TABLE IF NOT EXISTS report_holds (
      roster_id TEXT PRIMARY KEY,
      term_id TEXT,
      reason TEXT,
      by_ TEXT,
      at_ TEXT
    )`,

    `CREATE TABLE IF NOT EXISTS report_notes (
      id TEXT PRIMARY KEY,
      term_id TEXT NOT NULL,
      roster_id TEXT NOT NULL,
      class TEXT,
      stream TEXT,
      comment TEXT,
      conduct TEXT,
      head_comment TEXT,
      additional_comment TEXT,
      updated_at TEXT NOT NULL,
      updated_by TEXT NOT NULL
    )`,
    `CREATE INDEX IF NOT EXISTS idx_notes_term ON report_notes(term_id)`,

    `CREATE TABLE IF NOT EXISTS grading_scale (
      school_id TEXT PRIMARY KEY,
      bands_json TEXT NOT NULL,
      updated_at TEXT,
      updated_by TEXT
    )`,

    `CREATE TABLE IF NOT EXISTS points_settings (
      school_id TEXT PRIMARY KEY,
      bands_json TEXT NOT NULL,
      subsidiary_pass_mark REAL NOT NULL,
      updated_at TEXT,
      updated_by TEXT
    )`,

    `CREATE TABLE IF NOT EXISTS audit_log (
      id TEXT PRIMARY KEY,
      school_id TEXT NOT NULL,
      at_ TEXT NOT NULL,
      action TEXT NOT NULL,
      by_ TEXT,
      term_id TEXT,
      class TEXT,
      stream TEXT,
      subject TEXT,
      roster_id TEXT,
      payload_json TEXT
    )`,
    `CREATE INDEX IF NOT EXISTS idx_audit_school_at ON audit_log(school_id, at_ DESC)`,
  ];
  for (let i = 0; i < stmts.length; i += 50) {
    await conn.batch(stmts.slice(i, i + 50).map((s) => conn.prepare(s)));
  }
  _schemaReady = true;
}

async function bootstrap(env) {
  try { await ensureSchema(env); } catch (err) {
    console.error("Schema bootstrap failed:", err && err.message);
  }
}

// ================================================================
// Audit log
// ================================================================

async function auditStmt(env, schoolId, entry) {
  const id = (await sha256Hex(`${Date.now()}|${Math.random()}|${entry.action}`)).slice(0, 24);
  return {
    sql: `INSERT INTO audit_log
      (id, school_id, at_, action, by_, term_id, class, stream, subject, roster_id, payload_json)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    params: [
      id, schoolId, new Date().toISOString(), entry.action,
      entry.by || null, entry.termId || null, entry.class || null, entry.stream || null,
      entry.subject || null, entry.rosterId || null, JSON.stringify(entry),
    ],
  };
}

async function logChange(env, schoolId, entry) {
  try {
    const stmt = await auditStmt(env, schoolId, entry);
    await d1Run(env, stmt.sql, ...stmt.params);
  } catch (err) {
    console.error("Audit log write failed:", err && err.message);
  }
}

// ================================================================
// Assessment types / scoring settings (D1, teachers-worker-owned)
// ================================================================

async function listAssessmentTypes(env, schoolId, { includeArchived = false } = {}) {
  const key = `types:${schoolId}:${includeArchived ? "all" : "active"}`;
  const cached = cacheGet(key);
  if (cached) return cached;

  let rows;
  try {
    rows = await d1All(
      env,
      `SELECT id, name, order_ AS ord, archived FROM assessment_types WHERE school_id = ? ORDER BY order_ ASC, name ASC LIMIT ?`,
      schoolId, MAX_TYPES
    );
  } catch {
    rows = await d1All(env, `SELECT * FROM assessment_types WHERE school_id = ? LIMIT ?`, schoolId, MAX_TYPES);
  }
  const types = rows.map((r) => ({
    id: r.id,
    name: r.name || "",
    order: typeof r.ord === "number" ? r.ord : (typeof r.order === "number" ? r.order : 100),
    archived: r.archived === 1 || r.archived === true,
  }));
  const filtered = includeArchived ? types : types.filter((t) => !t.archived);
  filtered.sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
  cacheSet(key, filtered, CACHE_TTL_MS);
  return filtered;
}

async function getScoringSettings(env, schoolId) {
  const key = `scoring:${schoolId}`;
  const cached = cacheGet(key);
  if (cached) return cached;

  let row;
  try {
    row = await d1Get(env, `SELECT mode, cbc_levels_json FROM scoring_settings WHERE school_id = ?`, schoolId);
  } catch { row = null; }
  const mode = row && SCORING_MODES.includes(row.mode) ? row.mode : "percentage";
  let cbcLevels;
  try {
    const parsed = row && row.cbc_levels_json ? JSON.parse(row.cbc_levels_json) : null;
    cbcLevels = Array.isArray(parsed) && parsed.length > 0 ? parsed : null;
  } catch { cbcLevels = null; }
  if (!cbcLevels) {
    cbcLevels = [
      { level: 1, label: "Beginning", descriptor: "Needs support" },
      { level: 2, label: "Developing", descriptor: "Grasping with guidance" },
      { level: 3, label: "Achieved", descriptor: "Consistently meets standard" },
    ];
  }
  const out = { mode, cbcLevels };
  cacheSet(key, out, CACHE_TTL_MS);
  return out;
}

// ================================================================
// Grading scale (school-owned, D1)
// ================================================================

async function getGradingScale(env, schoolId) {
  const row = await d1Get(env, `SELECT bands_json FROM grading_scale WHERE school_id = ?`, schoolId);
  if (!row || !row.bands_json) return DEFAULT_GRADING_SCALE;
  try {
    const bands = JSON.parse(row.bands_json);
    if (Array.isArray(bands) && bands.length > 0) return { bands };
  } catch { /* fall through */ }
  return DEFAULT_GRADING_SCALE;
}

function applyGradingScale(pct, scale) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return null;
  const bands = (scale && Array.isArray(scale.bands) ? scale.bands : DEFAULT_GRADING_SCALE.bands)
    .slice().sort((a, b) => b.min - a.min);
  for (const band of bands) if (pct >= band.min) return band;
  return null;
}

// ================================================================
// A-level points settings (school-owned, D1)
// ================================================================

async function getPointsSettings(env, schoolId) {
  const row = await d1Get(env, `SELECT bands_json, subsidiary_pass_mark FROM points_settings WHERE school_id = ?`, schoolId);
  if (!row) return DEFAULT_POINTS_SETTINGS;
  let bands = DEFAULT_POINTS_SETTINGS.bands;
  try {
    const parsed = JSON.parse(row.bands_json || "[]");
    if (Array.isArray(parsed) && parsed.length > 0) bands = parsed;
  } catch { /* keep defaults */ }
  const passMark = Number.isFinite(row.subsidiary_pass_mark) ? row.subsidiary_pass_mark : DEFAULT_POINTS_SETTINGS.subsidiaryPassMark;
  return { bands, subsidiaryPassMark: passMark };
}

// A-level detection: Senior 5 / Senior 6 (with or without a trailing stream
// label). Anything else — including Senior 1 through 4 — is not A-level.
function isALevelClass(className) {
  const c = String(className || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  return /^(senior|s)[56](?!\d)/.test(c);
}
function isSubsidiarySubject(subject) {
  const s = String(subject || "").toLowerCase().replace(/[^a-z0-9]/g, "");
  return s === "generalpaper" || s === "gp" || /^sub(sidiary)?(ict|math|maths|mathematics)$/.test(s);
}

function gradeFromPercent(pct, pointsSettings) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return null;
  const bands = (pointsSettings.bands || []).slice().sort((a, b) => b.min - a.min);
  for (const band of bands) if (pct >= band.min) return band.grade;
  return null;
}

function pointsForSubject(pct, subject, pointsSettings) {
  if (pct === null || pct === undefined || !Number.isFinite(pct)) return 0;
  if (isSubsidiarySubject(subject)) {
    return pct >= (pointsSettings.subsidiaryPassMark ?? 50) ? 1 : 0;
  }
  const grade = gradeFromPercent(pct, pointsSettings);
  return grade && grade in PRINCIPAL_POINTS ? PRINCIPAL_POINTS[grade] : 0;
}

// ================================================================
// Terms (teachers-worker-owned, D1)
// ================================================================

async function getTermOrThrow(env, schoolId, termId) {
  const row = await d1Get(env, `SELECT * FROM terms WHERE school_id = ? AND id = ?`, schoolId, termId);
  if (!row) throw new HttpError(404, "Term not found");
  return row;
}

async function resolveTermId(env, schoolId, requested) {
  // Never trust a caller-supplied termId without verifying it belongs to this school.
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

function parseJsonColumn(raw, fallback) {
  if (!raw) return fallback;
  try { return JSON.parse(raw); } catch { return fallback; }
}

function resolveMaxScores(term, types) {
  const custom = parseJsonColumn(term && term.max_scores_json, {});
  const out = {};
  for (const t of types) {
    const v = toNumberOrNull(custom[t.id]);
    if (v && v > 0) out[t.id] = v;
    else if (DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()]) out[t.id] = DEFAULT_MAX_SCORES_BY_NAME[t.name.toUpperCase()];
    else out[t.id] = FALLBACK_MAX_SCORE;
  }
  return out;
}

function resolveWeights(term, types) {
  const arr = parseJsonColumn(term && term.weights_json, []);
  const out = {};
  for (const t of types) out[t.id] = 0;
  for (const w of arr) {
    if (w && types.some((t) => t.id === w.typeId) && typeof w.weight === "number") out[w.typeId] = w.weight;
  }
  return out;
}

function termPublishRank(term) {
  return term && (term.publish_rank === 1 || term.publish_rank === true);
}

// ================================================================
// Teachers / permissions (D1)
// ================================================================

async function loadTeacher(env, schoolId, uid) {
  const row = await d1Get(env, `SELECT * FROM teachers WHERE school_id = ? AND uid = ?`, schoolId, uid);
  if (!row) return { assignments: [], classTeacherOf: null, permissions: {} };
  return {
    assignments: parseJsonColumn(row.assignments_json, []),
    classTeacherOf: parseJsonColumn(row.class_teacher_of_json, null),
    permissions: parseJsonColumn(row.permissions_json, {}),
  };
}

const DEFAULT_PERMISSIONS = { canEnterMarks: true, canUploadHoliday: true, canReadParentReports: true };

async function requirePermission(env, schoolId, uid, key) {
  const t = await loadTeacher(env, schoolId, uid);
  const perms = { ...DEFAULT_PERMISSIONS, ...(t.permissions || {}) };
  if (!perms[key]) throw new HttpError(403, `You do not have the ${key} permission`);
  return t;
}

function assignmentCovers(teacher, className, stream, subject) {
  const list = Array.isArray(teacher.assignments) ? teacher.assignments : [];
  const norm = (v) => String(v || "").trim().toLowerCase();
  return list.some((a) =>
    a && norm(a.class) === norm(className) &&
    (!a.stream || norm(a.stream) === norm(stream)) &&
    norm(a.subject) === norm(subject)
  );
}
function classTeacherCovers(teacher, className, stream) {
  const ct = teacher.classTeacherOf;
  if (!ct || typeof ct !== "object") return false;
  if (ct.class !== className) return false;
  if (ct.stream && stream && ct.stream !== stream) return false;
  return true;
}

// ================================================================
// Roster (D1, roster-worker-owned). Column is class_name.
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
    // Prefer the roster's canonical name_key; fall back to a locally computed one.
    const nk = r.name_key || nameKey(firstName, lastName);
    return {
      id: r.id,
      firstName, lastName,
      nameKey: nk,
      class: r.class_name || "",
      stream: r.stream || "",
      gender: r.gender || "",
      subjects: parseJsonColumn(r.subjects_json, []),
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
// Cloudinary (UNCHANGED)
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
    // Skip the template hint row that starts "Enter scores…" / "Enter levels…".
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
// Mark sheet state (D1)
// ================================================================

async function getSheet(env, schoolId, sheetId) {
  return d1Get(env, `SELECT * FROM mark_sheets WHERE id = ?`, sheetId);
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
// State signing (ambiguous rows)
// ================================================================

async function getStateKey(env) {
  const sa = getServiceAccount(env);
  const material = utf8ToBytes(String(sa.private_key) + "|" + sa.project_id + "|edutrace-marks-v1");
  const digest = await crypto.subtle.digest("SHA-256", material);
  return crypto.subtle.importKey("raw", digest, { name: "HMAC", hash: "SHA-256" }, false, ["sign", "verify"]);
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
  let sql = `SELECT * FROM marks WHERE term_id = ? AND class = ? AND subject = ? AND type_id = ?`;
  const params = [termId, className, subject, typeId];
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

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "You are not assigned to this class/stream/subject");
  }

  const [term, types, scoring] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, `Assessment type "${typeId}" not found or archived`);

  const maxScores = resolveMaxScores(term, types);
  const weights = resolveWeights(term, types);
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
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
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

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);
  if (!Array.isArray(body.entries) || body.entries.length === 0) throw new HttpError(400, "entries required");
  if (body.entries.length > MAX_GRID_ROWS) throw new HttpError(400, `Max ${MAX_GRID_ROWS} entries`);

  const [term, types, scoring] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (!types.some((t) => t.id === typeId)) throw new HttpError(404, "Assessment type not found");

  const maxScores = resolveMaxScores(term, types);
  const defaultMax = toNumberOrNull(body.maxScore) || maxScores[typeId];
  const maxLevel = scoring.cbcLevels.length;

  const students = await fetchRosterStudents(env, schoolId, className, stream);
  const validRoster = new Set(students.map((s) => s.id));

  const nowIso = new Date().toISOString();
  const uid = String(user.sub);
  const stmts = [];
  const auditRows = [];

  for (const e of body.entries) {
    if (!e || typeof e !== "object") throw new HttpError(400, "Each entry must be an object");
    const rosterId = str(e.rosterId).trim();
    if (!rosterId || !validRoster.has(rosterId)) throw new HttpError(400, `rosterId not in class: ${rosterId}`);

    const status = MARK_STATUSES.includes(e.status) ? e.status : null;
    if (!status) throw new HttpError(400, `Invalid status for ${rosterId}`);

    let score = null, level = null, cellMax = null;
    if (scoring.mode === "percentage") {
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
      sql: `INSERT INTO marks (id, term_id, roster_id, class, stream, subject, type_id, scoring_mode, status, score, level, max_score, source, cloudinary_url, entered_by, entered_at, updated_by, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, ?, ?, ?, ?)
            ON CONFLICT(id) DO UPDATE SET
              status = excluded.status, score = excluded.score, level = excluded.level, max_score = excluded.max_score,
              updated_by = excluded.updated_by, updated_at = excluded.updated_at`,
      params: [
        markId, termId, rosterId, className, stream || "", subject, typeId, scoring.mode,
        status, score, level, cellMax, uid, nowIso, uid, nowIso,
      ],
    });
    auditRows.push({ rosterId, markId, status, score, level, maxScore: cellMax });
  }

  stmts.push(await auditStmt(env, schoolId, {
    action: "marks.grid.write",
    by: uid, termId, class: className, stream, subject, typeId,
    count: stmts.length, entries: auditRows,
  }));

  await runBatch(env, stmts);
  return { status: "ok", updated: stmts.length - 1, termId, class: className, stream, subject, typeId };
}

async function handleDeleteMark(request, env, schoolId, markId) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const doc = await d1Get(env, `SELECT * FROM marks WHERE id = ?`, markId);
  if (!doc) throw new HttpError(404, "Mark not found");
  await getTermOrThrow(env, schoolId, doc.term_id);
  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, doc.class, doc.stream || "", doc.subject)) throw new HttpError(403, "Not your assignment");
  }
  await assertSheetWritable(env, schoolId, doc.term_id, doc.class, doc.stream || "", doc.subject, user);
  const stmts = [
    { sql: `DELETE FROM marks WHERE id = ?`, params: [markId] },
    await auditStmt(env, schoolId, { action: "marks.cell.delete", by: String(user.sub), rosterId: doc.roster_id }),
  ];
  await runBatch(env, stmts);
  return { status: "ok" };
}

// ================================================================
// Excel: template
// ================================================================

async function handleDownloadTemplate(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";
  const subject = cleanText(url.searchParams.get("subject") || "", 120, "subject");
  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "You are not assigned to this class/stream/subject");
  }

  const [types, scoring] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
  ]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");
  const students = await fetchRosterStudents(env, schoolId, className, stream);

  const headers = ["rosterId", "Name", ...types.map((t) => t.name)];
  const rows = students.map((s) => [s.id, displayName(s.firstName, s.lastName), ...types.map(() => "")]);
  const hint = scoring.mode === "cbc"
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
// Excel: upload
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

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
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

  const [types, scoring, term] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getTermOrThrow(env, schoolId, termId),
  ]);
  if (types.length === 0) throw new HttpError(400, "No active assessment types");

  const { rows: parsedRows } = readRowsFromWorkbook(bytes, types);
  const maxScores = resolveMaxScores(term, types);
  const maxLevel = scoring.cbcLevels.length;

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
      const parsed = interpretCell(raw, scoring.mode, maxLevel);
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
    termId, className, stream, subject, scoring, types, maxScores, matched, cloudinaryUrl, uid: String(user.sub),
  });

  return {
    status: "ok", ...writeResult,
    summary: { matched: matched.length, ambiguous: 0, notFound: notFound.length, cellErrors: cellErrors.length },
    notFound, cellErrors, cloudinaryUrl,
  };
}

async function writeParsedMarks(env, schoolId, ctx) {
  const { termId, className, stream, subject, scoring, maxScores, matched, cloudinaryUrl, uid } = ctx;
  const nowIso = new Date().toISOString();
  const stmts = [];
  let written = 0;

  for (const entry of matched) {
    const rosterId = entry.student.id;
    for (const [typeId, parsed] of Object.entries(entry.cells)) {
      const markId = await computeMarkId(termId, rosterId, subject, typeId);
      const cellMax = scoring.mode === "percentage" ? maxScores[typeId] : null;
      if (scoring.mode === "percentage" && parsed.status === "entered") {
        if (parsed.score > cellMax) throw new HttpError(400, `Score ${parsed.score} exceeds maxScore ${cellMax} (row ${rosterId}, type ${typeId})`);
      }
      stmts.push({
        sql: `INSERT INTO marks (id, term_id, roster_id, class, stream, subject, type_id, scoring_mode, status, score, level, max_score, source, cloudinary_url, entered_by, entered_at, updated_by, updated_at)
              VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'excel', ?, ?, ?, ?, ?)
              ON CONFLICT(id) DO UPDATE SET
                status = excluded.status, score = excluded.score, level = excluded.level, max_score = excluded.max_score,
                source = excluded.source, cloudinary_url = excluded.cloudinary_url,
                updated_by = excluded.updated_by, updated_at = excluded.updated_at`,
        params: [
          markId, termId, rosterId, className, stream || "", subject, typeId, scoring.mode,
          parsed.status, parsed.score, parsed.level, cellMax, cloudinaryUrl || null, uid, nowIso, uid, nowIso,
        ],
      });
      written += 1;
    }
  }

  stmts.push(await auditStmt(env, schoolId, {
    action: "marks.excel.write", by: uid, termId, class: className, stream, subject, count: written, cloudinaryUrl,
  }));

  await runBatch(env, stmts);
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

  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }
  await assertSheetWritable(env, schoolId, termId, className, stream, subject, user);

  const [types, scoring, term] = await Promise.all([
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getTermOrThrow(env, schoolId, termId),
  ]);
  const maxScores = resolveMaxScores(term, types);
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
        const parsed = interpretCell(raw, scoring.mode, scoring.cbcLevels.length);
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
    termId, className, stream, subject, scoring, types, maxScores, matched: allMatched, cloudinaryUrl, uid: String(user.sub),
  });
  return { status: "ok", ...result };
}

// ================================================================
// Mark sheets (submit / lock / reopen)
// ================================================================

async function handleListSheets(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rows = await d1All(env, `SELECT * FROM mark_sheets WHERE term_id = ? LIMIT ?`, termId, MAX_SHEETS);
  return {
    termId,
    sheets: rows.map((r) => ({
      id: r.id, class: r.class || "", stream: r.stream || "", subject: r.subject || "",
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
  if (user.role === "teacher") {
    const t = await requirePermission(env, schoolId, String(user.sub), "canEnterMarks");
    if (!assignmentCovers(t, className, stream, subject)) throw new HttpError(403, "Not your assignment");
  }

  const sheetId = await computeSheetId(termId, className, stream, subject);
  const existing = (await getSheet(env, schoolId, sheetId)) || {};
  const nowIso = new Date().toISOString();
  const uid = String(user.sub);

  const fields = { term_id: termId, class: className, stream: stream || "", subject, status: targetStatus, updated_at: nowIso, updated_by: uid };
  if (targetStatus === "submitted") { fields.submitted_at = nowIso; fields.submitted_by = uid; }
  if (targetStatus === "locked")    { fields.locked_at = nowIso; fields.locked_by = uid; }
  if (targetStatus === "open")      { fields.reopened_at = nowIso; fields.reopened_by = uid; }

  const cols = Object.keys(fields);
  const stmts = [{
    sql: `INSERT INTO mark_sheets (id, ${cols.join(", ")}) VALUES (?, ${cols.map(() => "?").join(", ")})
          ON CONFLICT(id) DO UPDATE SET ${cols.filter((c) => c !== "id").map((c) => `${c} = excluded.${c}`).join(", ")}`,
    params: [sheetId, ...cols.map((c) => fields[c])],
  }];
  stmts.push(await auditStmt(env, schoolId, {
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
  const [term, types, scoring, scale, points] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
    getPointsSettings(env, schoolId),
  ]);
  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null, isCurrent: term.is_current === 1 || term.is_current === true },
    assessmentTypes: types,
    maxScores: resolveMaxScores(term, types),
    weights: resolveWeights(term, types),
    publishRank: termPublishRank(term),
    gradingScale: scale,
    pointsSettings: points,
  };
}

async function handleSetMaxScores(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);
  const termId = await resolveTermId(env, schoolId, body.termId);
  const existing = await getTermOrThrow(env, schoolId, termId);

  if (!body.maxScores || typeof body.maxScores !== "object" || Array.isArray(body.maxScores)) {
    throw new HttpError(400, "maxScores must be an object keyed by typeId");
  }
  const types = await listAssessmentTypes(env, schoolId, { includeArchived: true });
  const typeIds = new Set(types.map((t) => t.id));
  const clean = {};
  for (const [k, v] of Object.entries(body.maxScores)) {
    if (!typeIds.has(k)) throw new HttpError(400, `Unknown assessment type: ${k}`);
    const n = Number(v);
    if (!Number.isFinite(n) || n <= 0 || n > 1000) throw new HttpError(400, `maxScores.${k} must be 1-1000`);
    clean[k] = n;
  }
  const merged = { ...parseJsonColumn(existing.max_scores_json, {}), ...clean };
  const stmts = [{
    sql: `UPDATE terms SET max_scores_json = ?, updated_at = ?, updated_by = ? WHERE school_id = ? AND id = ?`,
    params: [JSON.stringify(merged), new Date().toISOString(), String(admin.sub), schoolId, termId],
  }];
  stmts.push(await auditStmt(env, schoolId, {
    action: "term.maxScores.set", by: String(admin.sub), termId,
    before: parseJsonColumn(existing.max_scores_json, {}), after: merged,
  }));
  await runBatch(env, stmts);
  return { termId, maxScores: merged };
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
  stmts.push(await auditStmt(env, schoolId, { action: "term.publishRank.set", by: String(admin.sub), termId, after: body.publishRank }));
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
    sql: `INSERT INTO grading_scale (school_id, bands_json, updated_at, updated_by) VALUES (?, ?, ?, ?)
          ON CONFLICT(school_id) DO UPDATE SET bands_json = excluded.bands_json, updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [schoolId, JSON.stringify(clean), nowIso, String(admin.sub)],
  }];
  stmts.push(await auditStmt(env, schoolId, { action: "settings.gradingScale.set", by: String(admin.sub), after: clean }));
  await runBatch(env, stmts);
  return { bands: clean };
}

async function handleSetPointsSettings(request, env, schoolId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const body = await safeJson(request);

  let bands = DEFAULT_POINTS_SETTINGS.bands;
  if (body.bands !== undefined) {
    if (!Array.isArray(body.bands) || body.bands.length === 0) throw new HttpError(400, "bands must be a non-empty array");
    bands = body.bands.map((b) => {
      if (!b || typeof b !== "object") throw new HttpError(400, "Invalid band");
      const min = Number(b.min);
      if (!Number.isFinite(min) || min < 0 || min > 100) throw new HttpError(400, "Each band needs min 0-100");
      return { min, grade: cleanText(b.grade, 2, "band.grade").toUpperCase() };
    });
  }
  let passMark = DEFAULT_POINTS_SETTINGS.subsidiaryPassMark;
  if (body.subsidiaryPassMark !== undefined) {
    const n = Number(body.subsidiaryPassMark);
    if (!Number.isFinite(n) || n < 0 || n > 100) throw new HttpError(400, "subsidiaryPassMark must be 0-100");
    passMark = n;
  }

  const nowIso = new Date().toISOString();
  const stmts = [{
    sql: `INSERT INTO points_settings (school_id, bands_json, subsidiary_pass_mark, updated_at, updated_by)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(school_id) DO UPDATE SET
            bands_json = excluded.bands_json, subsidiary_pass_mark = excluded.subsidiary_pass_mark,
            updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [schoolId, JSON.stringify(bands), passMark, nowIso, String(admin.sub)],
  }];
  stmts.push(await auditStmt(env, schoolId, { action: "settings.points.set", by: String(admin.sub), after: { bands, subsidiaryPassMark: passMark } }));
  await runBatch(env, stmts);
  return { bands, subsidiaryPassMark: passMark };
}

// ================================================================
// Report computation
// ================================================================

async function fetchMarksForStudent(env, schoolId, termId, rosterId) {
  return d1All(env, `SELECT * FROM marks WHERE term_id = ? AND roster_id = ?`, termId, rosterId);
}

async function fetchMarksForClass(env, schoolId, termId, className, stream) {
  let sql = `SELECT * FROM marks WHERE term_id = ? AND class = ?`;
  const params = [termId, className];
  if (stream) { sql += ` AND stream = ?`; params.push(stream); }
  return d1All(env, sql, ...params);
}

function computePercentageReport(marks, types, term, extra = {}) {
  const maxScores = resolveMaxScores(term, types);
  const weights = resolveWeights(term, types);
  const pointsSettings = extra.pointsSettings || DEFAULT_POINTS_SETTINGS;
  const className = extra.className || "";
  const aLevel = isALevelClass(className);

  const bySubject = new Map();
  for (const m of marks) {
    if (!m.subject || !m.type_id) continue;
    if (!bySubject.has(m.subject)) bySubject.set(m.subject, new Map());
    bySubject.get(m.subject).set(m.type_id, m);
  }

  const subjects = [];
  let totalPct = 0, subjectsWithData = 0;
  let totalPoints = 0, principalCount = 0, subsidiaryCount = 0;

  for (const [subject, cellsByType] of bySubject.entries()) {
    const assessments = {};
    let weightedRaw = 0, weightsPresent = 0;

    for (const t of types) {
      const c = cellsByType.get(t.id);
      const w = weights[t.id] ?? 0;
      if (!c) {
        assessments[t.id] = { typeId: t.id, typeName: t.name, status: "missing", score: null, maxScore: maxScores[t.id], weight: w, contribution: 0 };
        continue;
      }
      if (c.status === "entered" && typeof c.score === "number" && c.max_score > 0) {
        const contribution = (c.score / c.max_score) * w;
        weightedRaw += contribution;
        weightsPresent += w;
        assessments[t.id] = { typeId: t.id, typeName: t.name, status: "entered", score: c.score, maxScore: c.max_score, weight: w, contribution: round2(contribution) };
      } else {
        assessments[t.id] = { typeId: t.id, typeName: t.name, status: c.status || "missing", score: null, maxScore: c.max_score || maxScores[t.id], weight: w, contribution: 0 };
      }
    }

    const pct = weightsPresent > 0 ? round1((weightedRaw / weightsPresent) * 100) : null;
    const sub = {
      name: subject,
      assessments: Object.values(assessments),
      weightedRaw: round2(weightedRaw),
      weightsPresent: round2(weightsPresent),
      pct,
      grade: null,
    };

    if (aLevel) {
      const points = pointsForSubject(pct, subject, pointsSettings);
      sub.points = points;
      sub.isPrincipal = !isSubsidiarySubject(subject);
      sub.isSubsidiary = isSubsidiarySubject(subject);
      if (sub.isPrincipal) principalCount += 1; else subsidiaryCount += 1;
      totalPoints += points;
    }

    subjects.push(sub);
    if (pct !== null) { totalPct += pct; subjectsWithData += 1; }
  }

  subjects.sort((a, b) => a.name.localeCompare(b.name));
  const overallPct = subjectsWithData > 0 ? round1(totalPct / subjectsWithData) : null;

  return {
    subjects,
    overall: { pct: overallPct },
    rollup: null,
    aLevel: aLevel ? { totalPoints, averagePct: overallPct, principalCount, subsidiaryCount } : null,
  };
}

function computeCbcReport(marks, types, cbcLevels) {
  const levels = cbcLevels.slice().sort((a, b) => a.level - b.level);
  const levelByValue = new Map(levels.map((l) => [l.level, l]));

  const bySubject = new Map();
  for (const m of marks) {
    if (!m.subject || !m.type_id) continue;
    if (!bySubject.has(m.subject)) bySubject.set(m.subject, new Map());
    bySubject.get(m.subject).set(m.type_id, m);
  }

  const subjects = [];
  const rollup = { byLevel: new Map() };
  for (const l of levels) rollup.byLevel.set(l.level, 0);

  for (const [subject, cellsByType] of bySubject.entries()) {
    const assessments = [];
    let highest = 0;
    for (const t of types) {
      const c = cellsByType.get(t.id);
      if (!c) { assessments.push({ typeId: t.id, typeName: t.name, status: "missing", level: null }); continue; }
      if (c.status === "entered" && typeof c.level === "number") {
        assessments.push({ typeId: t.id, typeName: t.name, status: "entered", level: c.level });
        if (c.level > highest) highest = c.level;
      } else {
        assessments.push({ typeId: t.id, typeName: t.name, status: c.status, level: null });
      }
    }
    const subjectLevel = highest > 0 ? levelByValue.get(highest) || null : null;
    subjects.push({ name: subject, assessments, level: subjectLevel });
    if (subjectLevel) rollup.byLevel.set(subjectLevel.level, (rollup.byLevel.get(subjectLevel.level) || 0) + 1);
  }

  subjects.sort((a, b) => a.name.localeCompare(b.name));
  const rollupRows = levels
    .map((l) => ({ level: l.level, label: l.label, count: rollup.byLevel.get(l.level) || 0 }))
    .filter((r) => r.count > 0);
  const summaryText = rollupRows.length
    ? rollupRows.map((r) => `${r.label}: ${r.count} subject${r.count === 1 ? "" : "s"}`).join(" · ")
    : "No marks recorded yet";

  return { subjects, overall: { pct: null }, rollup: { rows: rollupRows, summaryText }, aLevel: null };
}

// Whole-class, all streams together. A-level ranks by total points,
// O-level by overall percentage. Ties share a position.
async function rankClass(env, schoolId, termId, className, students, marksByStudent, types, term, scoring, pointsSettings) {
  const aLevel = isALevelClass(className);
  const entries = [];

  for (const s of students) {
    const marks = marksByStudent.get(s.id) || [];
    let key = null;
    if (scoring.mode === "percentage") {
      const rep = computePercentageReport(marks, types, term, { pointsSettings, className });
      if (aLevel) {
        if (rep.aLevel) key = rep.aLevel.totalPoints;
      } else if (rep.overall.pct !== null) key = rep.overall.pct;
    } else {
      const maxLevel = scoring.cbcLevels.length;
      key = marks.filter((m) => m.status === "entered" && m.level === maxLevel).length;
    }
    if (key === null) continue;
    entries.push({ rosterId: s.id, key });
  }

  entries.sort((a, b) => b.key - a.key);
  const positions = new Map();
  let lastKey = null, lastPos = 0;
  entries.forEach((e, i) => {
    if (e.key !== lastKey) { lastPos = i + 1; lastKey = e.key; }
    positions.set(e.rosterId, lastPos);
  });
  return { positions, totalRanked: entries.length };
}

async function getReportHold(env, schoolId, rosterId, termId) {
  const hold = await d1Get(env, `SELECT * FROM report_holds WHERE roster_id = ?`, rosterId);
  if (!hold) return null;
  if (hold.term_id && hold.term_id !== termId) return null;
  return { rosterId: hold.roster_id, termId: hold.term_id || null, reason: hold.reason || "", by: hold.by_ || null, at: hold.at_ || null };
}

async function getReportNote(env, schoolId, termId, rosterId) {
  const id = `${termId}_${rosterId}`;
  const row = await d1Get(env, `SELECT * FROM report_notes WHERE id = ?`, id);
  if (!row) return null;
  return {
    comment: row.comment || "", conduct: row.conduct || "",
    headComment: row.head_comment || "", additionalComment: row.additional_comment || "",
  };
}

async function buildStudentReport(env, schoolId, rosterId, termId, { includeRank, includeHold }) {
  const [term, types, scoring, scale, pointsSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    listAssessmentTypes(env, schoolId, { includeArchived: true }),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
    getPointsSettings(env, schoolId),
  ]);

  const studentRow = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!studentRow) throw new HttpError(404, "Student not found in roster");
  if (studentRow.record_type && studentRow.record_type !== "student") throw new HttpError(400, "rosterId is not a student");

  const student = {
    id: studentRow.id,
    firstName: studentRow.first_name || "",
    lastName: studentRow.last_name || "",
    class: studentRow.class_name || "",
    stream: studentRow.stream || "",
    gender: studentRow.gender || "",
    subjects: parseJsonColumn(studentRow.subjects_json, []),
  };

  const marks = await fetchMarksForStudent(env, schoolId, termId, rosterId);
  const activeTypes = types.filter((t) => !t.archived);
  const reportTypes = activeTypes;

  let report;
  if (scoring.mode === "cbc") {
    report = computeCbcReport(marks, reportTypes, scoring.cbcLevels);
  } else {
    report = computePercentageReport(marks, reportTypes, term, { pointsSettings, className: student.class });
    for (const s of report.subjects) s.grade = applyGradingScale(s.pct, scale);
    report.overall.grade = applyGradingScale(report.overall.pct, scale);
  }

  const taken = Array.isArray(student.subjects) ? student.subjects : [];
  const have = new Set(report.subjects.map((s) => String(s.name).toLowerCase()));
  for (const name of taken) {
    const key = String(name || "").toLowerCase();
    if (!key || have.has(key)) continue;
    have.add(key);
    const blank = reportTypes.map((t) => scoring.mode === "cbc"
      ? { typeId: t.id, typeName: t.name, status: "missing", level: null }
      : { typeId: t.id, typeName: t.name, status: "missing", score: null, maxScore: null, weight: 0, contribution: 0 });
    report.subjects.push(scoring.mode === "cbc"
      ? { name, assessments: blank, level: null }
      : { name, assessments: blank, weightedRaw: 0, weightsPresent: 0, pct: null, grade: null });
  }
  report.subjects.sort((a, b) => a.name.localeCompare(b.name));

  // Rank across the whole class, all streams together.
  let rank = null;
  const publishRank = termPublishRank(term);
  if (includeRank && publishRank) {
    const students = await fetchRosterStudents(env, schoolId, student.class, "");
    const marksForClass = await fetchMarksForClass(env, schoolId, termId, student.class, "");
    const marksByStudent = new Map();
    for (const m of marksForClass) {
      if (!m.roster_id) continue;
      if (!marksByStudent.has(m.roster_id)) marksByStudent.set(m.roster_id, []);
      marksByStudent.get(m.roster_id).push(m);
    }
    const { positions, totalRanked } = await rankClass(env, schoolId, termId, student.class, students, marksByStudent, reportTypes, term, scoring, pointsSettings);
    rank = { position: positions.get(rosterId) || null, total: totalRanked };
  }

  const hold = includeHold ? await getReportHold(env, schoolId, rosterId, termId) : null;
  const note = await getReportNote(env, schoolId, termId, rosterId);
  const aLevel = report.aLevel;

  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: {
      id: termId, name: term.name || "", year: term.year || null,
      startDate: term.start_date || "", endDate: term.end_date || "",
    },
    student: {
      rosterId,
      firstName: student.firstName,
      lastName: student.lastName,
      class: student.class,
      stream: student.stream,
      gender: student.gender,
    },
    subjects: report.subjects,
    overall: report.overall,
    rollup: report.rollup,
    ...(aLevel ? {
      points: {
        totalPoints: aLevel.totalPoints,
        averagePct: aLevel.averagePct,
        principalCount: aLevel.principalCount,
        subsidiaryCount: aLevel.subsidiaryCount,
      },
    } : {}),
    rank,
    hold,
    note,
    gradingScale: scoring.mode === "percentage" ? scale : null,
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
  let consent = null;
  try { consent = await d1Get(env, `SELECT consented FROM consent WHERE user_uid = ? AND school_id = ?`, String(user.sub), schoolId); }
  catch { consent = null; }
  if (!consent || !(consent.consented === 1 || consent.consented === true)) throw new HttpError(403, "Parent consent is missing");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false });
}

async function handleGetStudentReport(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class_name, student.stream || "")) throw new HttpError(403, "You are not the class teacher for this student");
  }
  return buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true });
}

async function handleGetClassReport(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const className = cleanText(url.searchParams.get("class") || "", 60, "class");
  const stream = url.searchParams.get("stream") ? cleanText(url.searchParams.get("stream"), 60, "stream") : "";

  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, className, stream)) throw new HttpError(403, "Not your class");
  }

  const [term, students, types, scoring, scale, pointsSettings] = await Promise.all([
    getTermOrThrow(env, schoolId, termId),
    fetchRosterStudents(env, schoolId, className, stream),
    listAssessmentTypes(env, schoolId),
    getScoringSettings(env, schoolId),
    getGradingScale(env, schoolId),
    getPointsSettings(env, schoolId),
  ]);

  const allMarks = await fetchMarksForClass(env, schoolId, termId, className, stream);
  const byStudent = new Map();
  for (const m of allMarks) {
    if (!m.roster_id) continue;
    if (!byStudent.has(m.roster_id)) byStudent.set(m.roster_id, []);
    byStudent.get(m.roster_id).push(m);
  }

  const rows = students.map((s) => {
    const marks = byStudent.get(s.id) || [];
    let report;
    if (scoring.mode === "cbc") report = computeCbcReport(marks, types, scoring.cbcLevels);
    else {
      report = computePercentageReport(marks, types, term, { pointsSettings, className });
      for (const sub of report.subjects) sub.grade = applyGradingScale(sub.pct, scale);
      report.overall.grade = applyGradingScale(report.overall.pct, scale);
    }
    const out = {
      rosterId: s.id,
      firstName: s.firstName || "",
      lastName: s.lastName || "",
      stream: s.stream || "",
      subjects: report.subjects,
      overall: report.overall,
      rollup: report.rollup,
    };
    if (report.aLevel) out.points = report.aLevel;
    return out;
  });

  if (termPublishRank(term) || user.role === "schoolAdmin") {
    // Position is across the whole class, all streams together. If the caller
    // asked for one stream, we widen the ranking pool to the full class.
    const rankStudents = stream ? await fetchRosterStudents(env, schoolId, className, "") : students;
    let rankMarks = byStudent;
    if (stream) {
      rankMarks = new Map();
      for (const m of await fetchMarksForClass(env, schoolId, termId, className, "")) {
        if (!m.roster_id) continue;
        if (!rankMarks.has(m.roster_id)) rankMarks.set(m.roster_id, []);
        rankMarks.get(m.roster_id).push(m);
      }
    }
    const { positions } = await rankClass(env, schoolId, termId, className, rankStudents, rankMarks, types, term, scoring, pointsSettings);
    for (const r of rows) {
      const p = positions.get(r.rosterId);
      if (p) r.position = p;
    }
  }

  return {
    scoringMode: scoring.mode,
    cbcLevels: scoring.cbcLevels,
    term: { id: termId, name: term.name || "", year: term.year || null },
    class: className, stream,
    publishRank: termPublishRank(term),
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
  .title-bar{background:var(--primary);color:#fff;padding:12px 26px;text-align:center}
  .title-bar h2{font-size:1rem;font-weight:700;text-transform:uppercase;letter-spacing:1.5px}
  .student-info{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;padding:18px 26px;background:var(--bg-light);border-bottom:1px solid var(--border)}
  .student-info div{font-size:.85rem}
  .student-info label{display:block;font-size:.7rem;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);font-weight:700;margin-bottom:2px}
  table{width:100%;border-collapse:collapse;table-layout:fixed}
  th,td{padding:9px 8px;border-bottom:1px solid var(--border);font-size:.85rem;text-align:center;word-wrap:break-word;overflow-wrap:break-word}
  th{background:var(--primary);color:#fff;font-weight:700;text-transform:uppercase;font-size:.7rem;letter-spacing:.5px;padding:10px 6px}
  th.subject,td.subject{text-align:left;font-weight:600;width:26%}
  th.assessments,td.assessments{text-align:left;width:44%}
  th.total,td.total{width:15%}
  th.grade,td.grade{width:15%}
  td.total{font-weight:800;color:var(--primary);font-variant-numeric:tabular-nums}
  td.grade{font-weight:800}
  .chip{display:inline-block;margin:2px 3px 2px 0;padding:2px 6px;border:1px solid var(--border);background:var(--bg-light);font-size:.7rem;line-height:1.2;white-space:nowrap}
  .chip .t{font-weight:700;color:var(--text-muted)}
  .chip.abs{background:#fef3f2;border-color:#f2c7c3;color:#8a1c1c}
  .chip.exc{background:#fff8e6;border-color:#f0dfa6;color:#7a5a00}
  .chip.na{background:#f3f4f6;border-color:#e1e4e8;color:#6b7280}
  .chip.missing{opacity:.55}
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
  .no-print{padding:16px;text-align:center}
  .no-print button{padding:10px 22px;background:var(--primary);color:#fff;border:none;font-weight:700;text-transform:uppercase;letter-spacing:.5px;cursor:pointer;font-family:inherit}
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

function cardHeader(school) {
  const schoolName = escapeHtml(school.name || "School");
  const logoUrl = school.logoUrl ? escapeHtml(school.logoUrl) : "";
  return `<div class="header">
    ${logoUrl ? `<img class="logo" src="${logoUrl}" alt="">` : `<div class="logo-fallback">${escapeHtml((schoolName[0] || "S").toUpperCase())}</div>`}
    <div class="school-info"><h1>${schoolName}</h1>
      ${school.location ? `<p>${escapeHtml(school.location)}</p>` : ""}
      ${school.phone ? `<p>Tel: ${escapeHtml(school.phone)}</p>` : ""}
      ${school.email ? `<p>Email: ${escapeHtml(school.email)}</p>` : ""}
    </div>
  </div>`;
}

function cardFooter(school) {
  const schoolName = escapeHtml(school.name || "School");
  return `<div class="footer">© ${new Date().getFullYear()} ${schoolName}. Powered by Edutrace.</div>`;
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

function assessmentChip(a) {
  const typeShort = escapeHtml(String(a.typeName || "").toUpperCase());
  if (a.status === "missing") return `<span class="chip missing"><span class="t">${typeShort}</span> —</span>`;
  if (a.status === "absent") return `<span class="chip abs"><span class="t">${typeShort}</span> ABS</span>`;
  if (a.status === "excused") return `<span class="chip exc"><span class="t">${typeShort}</span> EXC</span>`;
  if (a.status === "na") return `<span class="chip na"><span class="t">${typeShort}</span> N/A</span>`;
  if (typeof a.level === "number") return `<span class="chip"><span class="t">${typeShort}</span> ${a.level}</span>`;
  return `<span class="chip"><span class="t">${typeShort}</span> ${escapeHtml(String(a.score))}/${escapeHtml(String(a.maxScore))}</span>`;
}

function renderPercentageReportHtml({ report, school }) {
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";
  const t = report.term, s = report.student;
  const isALevel = !!(report.points);
  const publishedRank = report.rank && report.rank.position;

  const rowsHtml = report.subjects.map((sub) => {
    const chips = sub.assessments.map(assessmentChip).join("");
    const pct = sub.pct !== null ? sub.pct.toFixed(1) : "—";
    const grade = sub.grade ? escapeHtml(sub.grade.grade) : "—";
    return `<tr><td class="subject">${escapeHtml(sub.name)}</td><td class="assessments">${chips}</td><td class="total">${pct}</td><td class="grade">${grade}</td></tr>`;
  }).join("");

  const overallPct = report.overall.pct !== null ? report.overall.pct.toFixed(1) : "—";
  const overallGrade = report.overall.grade ? escapeHtml(report.overall.grade.grade) : "—";

  let pointsHtml = "";
  if (isALevel && report.points) {
    pointsHtml = `<div class="points-strip">
      <div><label>Total Points</label><strong>${report.points.totalPoints}</strong><small>principal + subsidiary</small></div>
      <div><label>Average %</label><strong>${report.points.averagePct !== null ? report.points.averagePct.toFixed(1) : "—"}%</strong><small>all subjects</small></div>
      <div><label>Position</label><strong>${publishedRank || "—"}</strong><small>of ${report.rank?.total ?? "—"}</small></div>
      <div><label>Principal / Subsidiary</label><strong>${report.points.principalCount} / ${report.points.subsidiaryCount}</strong><small>subjects</small></div>
    </div>`;
  }

  // On A-level classes, Position moves to the green strip, so it's dropped
  // from the top summary.
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

  return `${reportCardHead(`Report Card — ${s.firstName} ${s.lastName} — ${t.name}`, themeColor)}
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  ${cardHeader(school)}
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <colgroup><col style="width:26%"><col style="width:44%"><col style="width:15%"><col style="width:15%"></colgroup>
    <thead><tr><th class="subject">Subject</th><th class="assessments">Assessments</th><th class="total">Total %</th><th class="grade">Grade</th></tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="4" class="grade">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  ${pointsHtml}
  <div class="legend"><b>Grade legend:</b> ${legendHtml}</div>
  ${summaryHtml}
  ${cardComments(note)}
  ${cardFooter(school)}
</div></body></html>`;
}

function renderCbcReportHtml({ report, school }) {
  const themeColor = /^#[0-9a-fA-F]{6}$/.test(school.themeColor || "") ? school.themeColor : "#3570BC";
  const t = report.term, s = report.student;

  const rowsHtml = report.subjects.map((sub) => {
    const chips = sub.assessments.map(assessmentChip).join("");
    const finalLabel = sub.level ? escapeHtml(sub.level.label) : "—";
    return `<tr><td class="subject">${escapeHtml(sub.name)}</td><td class="assessments">${chips}</td><td class="grade">${finalLabel}</td></tr>`;
  }).join("");

  const rollup = report.rollup || { rows: [] };
  const rollupHtml = rollup.rows.map((r) => `<div><label>${escapeHtml(r.label)}</label><strong>${r.count}</strong><small>subject${r.count === 1 ? "" : "s"}</small></div>`).join("");
  const legendHtml = report.cbcLevels.map((l) => `<b>${l.level}</b> = ${escapeHtml(l.label)}`).join(" · ");
  const note = report.note || {};

  return `${reportCardHead(`Report Card — ${s.firstName} ${s.lastName} — ${t.name}`, themeColor)}
<div class="no-print"><button onclick="window.print()">Print Report Card</button></div>
<div class="card">
  ${cardHeader(school)}
  <div class="title-bar"><h2>Term Report Card — ${escapeHtml(t.name)}${t.year ? ", " + escapeHtml(String(t.year)) : ""}</h2></div>
  <div class="student-info">
    <div><label>Name</label><strong>${escapeHtml(s.firstName)} ${escapeHtml(s.lastName)}</strong></div>
    <div><label>Class</label><strong>${escapeHtml(s.class)}</strong></div>
    <div><label>Stream</label><strong>${escapeHtml(s.stream || "—")}</strong></div>
    <div><label>Gender</label><strong>${escapeHtml(s.gender || "—")}</strong></div>
  </div>
  <table>
    <colgroup><col style="width:36%"><col style="width:44%"><col style="width:20%"></colgroup>
    <thead><tr><th class="subject">Subject</th><th class="assessments">Assessments</th><th class="grade">Level</th></tr></thead>
    <tbody>${rowsHtml || `<tr><td colspan="3" class="grade">No marks recorded for this term.</td></tr>`}</tbody>
  </table>
  <div class="legend"><b>Level legend:</b> ${legendHtml}</div>
  <div class="summary">${rollupHtml || `<div><label>Subjects</label><strong>0</strong></div>`}</div>
  ${cardComments(note)}
  ${cardFooter(school)}
</div></body></html>`;
}

function renderReportCardHtml({ report, school }) {
  return report.scoringMode === "cbc"
    ? renderCbcReportHtml({ report, school })
    : renderPercentageReportHtml({ report, school });
}

async function loadSchool(env, schoolId) {
  try {
    const row = await d1Get(env, `SELECT * FROM schools WHERE id = ?`, schoolId);
    if (row) return {
      name: row.name || "School",
      logoUrl: row.logo_url || "",
      location: row.location || "",
      phone: row.phone || "",
      email: row.email || "",
      themeColor: row.theme_color || "",
    };
  } catch { /* fall through */ }
  return { name: "School" };
}

async function handleGetReportCard(request, env, schoolId, rosterId, url) {
  const user = await requireUser(request, env, schoolId, ["schoolAdmin", "teacher"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const student = await d1Get(env, `SELECT * FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!student) throw new HttpError(404, "Student not found");
  if (user.role === "teacher") {
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class_name, student.stream || "")) throw new HttpError(403, "Not your class");
  }
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: true }),
    loadSchool(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school }));
}

async function handleGetMyReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["student"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.rosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a roster record");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    loadSchool(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school }));
}

async function handleGetChildReportCard(request, env, schoolId, url) {
  const user = await requireUser(request, env, schoolId, ["parent"]);
  const termId = await resolveTermId(env, schoolId, url.searchParams.get("termId"));
  const rosterId = String(user.childRosterId || "");
  if (!rosterId) throw new HttpError(400, "Your account is not linked to a child's roster record");
  let consent = null;
  try { consent = await d1Get(env, `SELECT consented FROM consent WHERE user_uid = ? AND school_id = ?`, String(user.sub), schoolId); }
  catch { consent = null; }
  if (!consent || !(consent.consented === 1 || consent.consented === true)) throw new HttpError(403, "Parent consent is missing");
  await assertNotHeldForReader(env, schoolId, rosterId, termId, user);
  const [report, school] = await Promise.all([
    buildStudentReport(env, schoolId, rosterId, termId, { includeRank: true, includeHold: false }),
    loadSchool(env, schoolId),
  ]);
  return html(renderReportCardHtml({ report, school }));
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
    const t = await loadTeacher(env, schoolId, String(user.sub));
    if (!classTeacherCovers(t, student.class_name, student.stream || "")) throw new HttpError(403, "Only the class teacher can write this note");
  }

  const hasComment = Object.prototype.hasOwnProperty.call(body, "comment");
  const hasConduct = Object.prototype.hasOwnProperty.call(body, "conduct");
  const hasHead = Object.prototype.hasOwnProperty.call(body, "headComment");
  if (hasHead && user.role !== "schoolAdmin") throw new HttpError(403, "Only the school admin can write the headteacher comment");

  const id = `${termId}_${rosterId}`;
  const fields = {
    term_id: termId, roster_id: rosterId,
    class: student.class_name || "", stream: student.stream || "",
    updated_at: new Date().toISOString(), updated_by: String(user.sub),
  };
  if (hasComment) fields.comment = cleanOptionalText(body.comment, 2000, "comment");
  if (hasConduct) fields.conduct = cleanOptionalText(body.conduct, 200, "conduct");
  if (hasHead) fields.head_comment = cleanOptionalText(body.headComment, 2000, "headComment");
  if (Object.prototype.hasOwnProperty.call(body, "additionalComment")) {
    fields.additional_comment = cleanOptionalText(body.additionalComment, 2000, "additionalComment");
  }

  const existing = await d1Get(env, `SELECT * FROM report_notes WHERE id = ?`, id);
  const merged = {
    comment: hasComment ? fields.comment : (existing?.comment || ""),
    conduct: hasConduct ? fields.conduct : (existing?.conduct || ""),
    head_comment: hasHead ? fields.head_comment : (existing?.head_comment || ""),
    additional_comment: Object.prototype.hasOwnProperty.call(body, "additionalComment")
      ? fields.additional_comment
      : (existing?.additional_comment || ""),
  };

  const stmts = [{
    sql: `INSERT INTO report_notes (id, term_id, roster_id, class, stream, comment, conduct, head_comment, additional_comment, updated_at, updated_by)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT(id) DO UPDATE SET
            class = excluded.class, stream = excluded.stream,
            comment = excluded.comment, conduct = excluded.conduct,
            head_comment = excluded.head_comment, additional_comment = excluded.additional_comment,
            updated_at = excluded.updated_at, updated_by = excluded.updated_by`,
    params: [
      id, termId, rosterId, fields.class, fields.stream,
      merged.comment, merged.conduct, merged.head_comment, merged.additional_comment,
      fields.updated_at, fields.updated_by,
    ],
  }];
  stmts.push(await auditStmt(env, schoolId, { action: "report.note.set", by: String(user.sub), termId, rosterId }));
  await runBatch(env, stmts);

  const note = {
    termId, rosterId,
    class: fields.class, stream: fields.stream,
    comment: merged.comment, conduct: merged.conduct,
    headComment: merged.head_comment, additionalComment: merged.additional_comment,
    updatedAt: fields.updated_at, updatedBy: fields.updated_by,
  };
  return { status: "ok", note };
}

// ================================================================
// Holds (scoped to this school via roster join)
// ================================================================

async function handleListHolds(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const termId = url.searchParams.get("termId");
  let rows;
  if (termId) rows = await d1All(env, `SELECT h.* FROM report_holds h JOIN roster r ON r.id = h.roster_id WHERE r.school_id = ? AND h.term_id = ? LIMIT ?`, schoolId, termId, MAX_HOLDS);
  else rows = await d1All(env, `SELECT h.* FROM report_holds h JOIN roster r ON r.id = h.roster_id WHERE r.school_id = ? LIMIT ?`, schoolId, MAX_HOLDS);
  return {
    holds: rows.map((r) => ({
      rosterId: r.roster_id, termId: r.term_id || null,
      reason: r.reason || "", by: r.by_ || null, at: r.at_ || null,
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
    sql: `INSERT INTO report_holds (roster_id, term_id, reason, by_, at_) VALUES (?, ?, ?, ?, ?)
          ON CONFLICT(roster_id) DO UPDATE SET term_id = excluded.term_id, reason = excluded.reason, by_ = excluded.by_, at_ = excluded.at_`,
    params: [rosterId, termId, reason, String(admin.sub), at],
  }];
  stmts.push(await auditStmt(env, schoolId, { action: "report.hold.create", by: String(admin.sub), termId, rosterId, reason }));
  await runBatch(env, stmts);
  return { status: "ok", hold: { termId, rosterId, reason, by: String(admin.sub), at } };
}

async function handleDeleteHold(request, env, schoolId, rosterId) {
  const admin = await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const owns = await d1Get(env, `SELECT id FROM roster WHERE school_id = ? AND id = ?`, schoolId, rosterId);
  if (!owns) throw new HttpError(404, "Student not found");
  const stmts = [
    { sql: `DELETE FROM report_holds WHERE roster_id = ?`, params: [rosterId] },
    await auditStmt(env, schoolId, { action: "report.hold.delete", by: String(admin.sub), rosterId }),
  ];
  await runBatch(env, stmts);
  return { status: "ok" };
}

// ================================================================
// Audit
// ================================================================

async function handleGetAuditLog(request, env, schoolId, url) {
  await requireUser(request, env, schoolId, ["schoolAdmin"]);
  const requested = parseInt(url.searchParams.get("limit") || "50", 10);
  const limit = Math.min(Math.max(Number.isInteger(requested) ? requested : 50, 1), 100);
  const rows = await d1All(env, `SELECT * FROM audit_log WHERE school_id = ? ORDER BY at_ DESC LIMIT ?`, schoolId, limit);
  return {
    entries: rows.map((r) => ({
      id: r.id,
      action: r.action,
      by: r.by_,
      termId: r.term_id,
      class: r.class,
      stream: r.stream,
      subject: r.subject,
      rosterId: r.roster_id,
      at: r.at_,
      ...(() => { try { return JSON.parse(r.payload_json || "{}"); } catch { return {}; } })(),
    })),
  };
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

  await bootstrap(env);

  let m;
  const d = decodeURIComponent;
  const S = "([^/]+)";

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
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/card/${S}$`)))) return handleGetReportCard(request, env, d(m[1]), d(m[2]), url);
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/reports/student/${S}$`)))) return json(await handleGetStudentReport(request, env, d(m[1]), d(m[2]), url));

  // Notes
  if (method === "PUT" && (m = path.match(new RegExp(`^/schools/${S}/notes$`)))) return json(await handleSetNote(request, env, d(m[1])));

  // Holds
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) return json(await handleListHolds(request, env, d(m[1]), url));
  if (method === "POST" && (m = path.match(new RegExp(`^/schools/${S}/holds$`)))) return json(await handleCreateHold(request, env, d(m[1])));
  if (method === "DELETE" && (m = path.match(new RegExp(`^/schools/${S}/holds/${S}$`)))) return json(await handleDeleteHold(request, env, d(m[1]), d(m[2])));

  // Audit
  if (method === "GET" && (m = path.match(new RegExp(`^/schools/${S}/audit-log$`)))) return json(await handleGetAuditLog(request, env, d(m[1]), url));

  return json({ error: "Not found" }, 404);
}

export default {
  async fetch(request, env) {
    try {
      return await handle(request, env);
    } catch (err) {
      if (err instanceof HttpError) return json({ error: err.message, ...err.extra }, err.status);
      console.error("Unhandled error:", err && err.stack ? err.stack : err);
      return json({ error: "Internal server error" }, 500);
    }
  },
};
