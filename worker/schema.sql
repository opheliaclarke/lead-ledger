-- Bob & Olivia lead ledger — server store
-- One shared ledger. Every save keeps the previous version, so nothing is ever
-- lost by a mistyped number, a wipe, or someone deleting a day.

CREATE TABLE IF NOT EXISTS users (
  code_hash  TEXT PRIMARY KEY,          -- sha-256 of the access code; the code itself is never stored
  name       TEXT NOT NULL,
  created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS state (
  id         INTEGER PRIMARY KEY CHECK (id = 1),
  json       TEXT NOT NULL,
  rev        INTEGER NOT NULL,          -- bumped on every save; the browser sends the rev it
  updated_at TEXT NOT NULL,             -- started from, so two people can't silently clobber
  updated_by TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS snapshots (
  id         INTEGER PRIMARY KEY AUTOINCREMENT,
  rev        INTEGER NOT NULL,
  ts         TEXT NOT NULL,
  who        TEXT NOT NULL,
  days       INTEGER NOT NULL,
  payments   INTEGER NOT NULL,
  net        TEXT NOT NULL,             -- outstanding at the time, as text, for the history list
  json       TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS snapshots_rev ON snapshots (rev DESC);
