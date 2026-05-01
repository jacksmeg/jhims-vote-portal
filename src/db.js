const fs = require("node:fs");
const path = require("node:path");
const { DatabaseSync } = require("node:sqlite");

const configuredDatabasePath = process.env.DATABASE_PATH
  ? path.isAbsolute(process.env.DATABASE_PATH)
    ? process.env.DATABASE_PATH
    : path.join(process.cwd(), process.env.DATABASE_PATH)
  : path.join(process.cwd(), "data", "vote-portal.sqlite");
const dataDirectory = path.dirname(configuredDatabasePath);
const databasePath = configuredDatabasePath;

fs.mkdirSync(dataDirectory, { recursive: true });

const db = new DatabaseSync(databasePath);

db.exec(`
  PRAGMA journal_mode = WAL;
  PRAGMA foreign_keys = ON;
  PRAGMA synchronous = NORMAL;
`);

function nowIso() {
  return new Date().toISOString();
}

function initDatabase(defaultElectionName) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS settings (
      key TEXT PRIMARY KEY,
      value TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS voters (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      staff_id TEXT NOT NULL UNIQUE,
      phone_number TEXT NOT NULL,
      full_name TEXT NOT NULL DEFAULT '',
      department TEXT NOT NULL DEFAULT '',
      has_voted INTEGER NOT NULL DEFAULT 0 CHECK (has_voted IN (0, 1)),
      voted_at TEXT,
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS positions (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT NOT NULL UNIQUE,
      sort_order INTEGER NOT NULL DEFAULT 0,
      is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS candidates (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      position_id INTEGER NOT NULL,
      name TEXT NOT NULL,
      photo_path TEXT NOT NULL DEFAULT '',
      bio TEXT NOT NULL DEFAULT '',
      sort_order INTEGER NOT NULL DEFAULT 0,
      is_active INTEGER NOT NULL DEFAULT 1 CHECK (is_active IN (0, 1)),
      created_at TEXT NOT NULL,
      updated_at TEXT NOT NULL,
      UNIQUE(position_id, name),
      FOREIGN KEY (position_id) REFERENCES positions (id) ON DELETE RESTRICT
    );

    CREATE TABLE IF NOT EXISTS ballots (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      voter_id INTEGER NOT NULL UNIQUE,
      submitted_at TEXT NOT NULL,
      ip_address TEXT NOT NULL DEFAULT '',
      user_agent TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL,
      FOREIGN KEY (voter_id) REFERENCES voters (id) ON DELETE RESTRICT
    );

    CREATE TABLE IF NOT EXISTS ballot_entries (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      ballot_id INTEGER NOT NULL,
      position_id INTEGER NOT NULL,
      candidate_id INTEGER NOT NULL,
      created_at TEXT NOT NULL,
      UNIQUE(ballot_id, position_id),
      FOREIGN KEY (ballot_id) REFERENCES ballots (id) ON DELETE CASCADE,
      FOREIGN KEY (position_id) REFERENCES positions (id) ON DELETE RESTRICT,
      FOREIGN KEY (candidate_id) REFERENCES candidates (id) ON DELETE RESTRICT
    );

    CREATE TABLE IF NOT EXISTS audit_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      actor_type TEXT NOT NULL,
      actor_identifier TEXT NOT NULL,
      action TEXT NOT NULL,
      details_json TEXT NOT NULL DEFAULT '{}',
      ip_address TEXT NOT NULL DEFAULT '',
      user_agent TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS election_archives (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      election_name TEXT NOT NULL,
      phase TEXT NOT NULL DEFAULT 'closed',
      opens_at TEXT NOT NULL DEFAULT '',
      closes_at TEXT NOT NULL DEFAULT '',
      archived_at TEXT NOT NULL,
      total_voters INTEGER NOT NULL DEFAULT 0,
      votes_cast INTEGER NOT NULL DEFAULT 0,
      turnout_rate REAL NOT NULL DEFAULT 0,
      settings_json TEXT NOT NULL DEFAULT '{}',
      metrics_json TEXT NOT NULL DEFAULT '{}',
      results_json TEXT NOT NULL DEFAULT '[]'
    );

    CREATE INDEX IF NOT EXISTS idx_voters_has_voted ON voters (has_voted);
    CREATE INDEX IF NOT EXISTS idx_candidates_position_id ON candidates (position_id);
    CREATE INDEX IF NOT EXISTS idx_ballot_entries_candidate_id ON ballot_entries (candidate_id);
    CREATE INDEX IF NOT EXISTS idx_ballot_entries_position_id ON ballot_entries (position_id);
    CREATE INDEX IF NOT EXISTS idx_audit_logs_created_at ON audit_logs (created_at DESC);
    CREATE INDEX IF NOT EXISTS idx_election_archives_archived_at
      ON election_archives (archived_at DESC);
  `);

  const defaults = [
    ["election_name", defaultElectionName || "Organization Election Portal"],
    ["election_phase", "setup"],
    ["opens_at", ""],
    ["closes_at", ""],
    ["results_visibility", "after_close"],
    ["theme_name", "heritage"],
  ];

  for (const [key, value] of defaults) {
    const existing = db.prepare("SELECT key FROM settings WHERE key = ?").get(key);
    if (!existing) {
      const timestamp = nowIso();
      db.prepare(
        "INSERT INTO settings (key, value, updated_at) VALUES (?, ?, ?)",
      ).run(key, String(value), timestamp);
    }
  }
}

function getAllSettings() {
  const rows = db.prepare("SELECT key, value FROM settings").all();
  return rows.reduce((accumulator, row) => {
    accumulator[row.key] = row.value;
    return accumulator;
  }, {});
}

function setSetting(key, value) {
  const timestamp = nowIso();
  db.prepare(`
    INSERT INTO settings (key, value, updated_at)
    VALUES (?, ?, ?)
    ON CONFLICT(key) DO UPDATE SET
      value = excluded.value,
      updated_at = excluded.updated_at
  `).run(key, String(value ?? ""), timestamp);
}

function runTransaction(callback) {
  db.exec("BEGIN IMMEDIATE");

  try {
    const result = callback(db);
    db.exec("COMMIT");
    return result;
  } catch (error) {
    db.exec("ROLLBACK");
    throw error;
  }
}

module.exports = {
  databasePath,
  db,
  getAllSettings,
  initDatabase,
  nowIso,
  runTransaction,
  setSetting,
};
