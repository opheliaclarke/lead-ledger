/**
 * Bob & Olivia lead ledger — server store.
 *
 * The browser owns the money model; this Worker only stores bytes and keeps
 * every previous version. That is deliberate: the settlement arithmetic exists
 * in exactly one place (the page), so the two can never drift apart.
 *
 * Auth: Authorization: Bearer <access code>. The code is sha-256'd and matched
 * against `users`; the code itself is never stored anywhere.
 */

const ALLOWED_ORIGINS = [
  "https://opheliaclarke.github.io",
  "http://localhost:8788",
  "http://127.0.0.1:8788",
];
const MAX_BODY = 2 * 1024 * 1024;   // 2 MB of ledger is ~15 years of daily rows
const KEEP_SNAPSHOTS = 300;

function cors(origin) {
  const h = {
    "Access-Control-Allow-Methods": "POST, GET, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin",
  };
  if (origin && ALLOWED_ORIGINS.includes(origin)) h["Access-Control-Allow-Origin"] = origin;
  return h;
}

function json(data, status, origin) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: { "Content-Type": "application/json; charset=utf-8", ...cors(origin) },
  });
}

async function sha256(s) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

async function auth(req, env) {
  const h = req.headers.get("Authorization") || "";
  const m = h.match(/^Bearer\s+(.+)$/i);
  if (!m) return null;
  const code = m[1].trim();
  if (!code || code.length > 200) return null;
  const row = await env.DB.prepare("SELECT name FROM users WHERE code_hash = ?")
    .bind(await sha256(code)).first();
  return row ? row.name : null;
}

async function currentState(env) {
  const row = await env.DB.prepare(
    "SELECT json, rev, updated_at, updated_by FROM state WHERE id = 1").first();
  if (row) return row;
  return { json: null, rev: 0, updated_at: null, updated_by: null };
}

async function writeState(env, stateJson, who, meta) {
  const cur = await currentState(env);
  const rev = cur.rev + 1;
  const now = new Date().toISOString();
  await env.DB.prepare(
    `INSERT INTO state (id, json, rev, updated_at, updated_by) VALUES (1, ?, ?, ?, ?)
     ON CONFLICT(id) DO UPDATE SET json=excluded.json, rev=excluded.rev,
       updated_at=excluded.updated_at, updated_by=excluded.updated_by`
  ).bind(stateJson, rev, now, who).run();
  await env.DB.prepare(
    `INSERT INTO snapshots (rev, ts, who, days, payments, net, json) VALUES (?,?,?,?,?,?,?)`
  ).bind(rev, now, who, meta.days | 0, meta.payments | 0, String(meta.net || "0.00"), stateJson).run();
  // keep the history bounded, oldest first
  await env.DB.prepare(
    `DELETE FROM snapshots WHERE id NOT IN
       (SELECT id FROM snapshots ORDER BY id DESC LIMIT ?)`).bind(KEEP_SNAPSHOTS).run();
  return { rev, updated_at: now, updated_by: who };
}

export default {
  async fetch(req, env) {
    const origin = req.headers.get("Origin");
    const url = new URL(req.url);
    const path = url.pathname.replace(/\/+$/, "") || "/";

    if (req.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(origin) });

    if (path === "/" || path === "/api/ping") {
      return json({ ok: true, service: "lead-ledger", time: new Date().toISOString() }, 200, origin);
    }

    if (!path.startsWith("/api/")) return json({ error: "not found" }, 404, origin);
    if (req.method !== "POST") return json({ error: "use POST" }, 405, origin);

    const who = await auth(req, env);
    if (!who) return json({ error: "bad access code" }, 401, origin);

    let body = {};
    if (path !== "/api/hello" && path !== "/api/load" && path !== "/api/history") {
      const raw = await req.text();
      if (raw.length > MAX_BODY) return json({ error: "ledger too large" }, 413, origin);
      try { body = raw ? JSON.parse(raw) : {}; }
      catch (e) { return json({ error: "bad json" }, 400, origin); }
    }

    try {
      if (path === "/api/hello") return json({ ok: true, name: who }, 200, origin);

      if (path === "/api/load") {
        const cur = await currentState(env);
        return json({
          ok: true, name: who, rev: cur.rev,
          state: cur.json ? JSON.parse(cur.json) : null,
          updated_at: cur.updated_at, updated_by: cur.updated_by,
        }, 200, origin);
      }

      if (path === "/api/save") {
        if (!body.state || typeof body.state !== "object" || !Array.isArray(body.state.days)) {
          return json({ error: "that is not a ledger" }, 400, origin);
        }
        const cur = await currentState(env);
        const base = Number(body.baseRev);
        if (cur.rev !== 0 && Number.isFinite(base) && base !== cur.rev) {
          // someone else saved since this browser last read — never overwrite them
          return json({
            error: "conflict", rev: cur.rev, updated_at: cur.updated_at,
            updated_by: cur.updated_by, state: JSON.parse(cur.json),
          }, 409, origin);
        }
        const out = await writeState(env, JSON.stringify(body.state), who, body.meta || {});
        return json({ ok: true, ...out }, 200, origin);
      }

      if (path === "/api/history") {
        const rs = await env.DB.prepare(
          `SELECT id, rev, ts, who, days, payments, net FROM snapshots
           ORDER BY id DESC LIMIT 60`).all();
        return json({ ok: true, items: rs.results || [] }, 200, origin);
      }

      if (path === "/api/restore") {
        const id = Number(body.id);
        if (!Number.isFinite(id)) return json({ error: "which version?" }, 400, origin);
        const snap = await env.DB.prepare(
          "SELECT json, days, payments, net FROM snapshots WHERE id = ?").bind(id).first();
        if (!snap) return json({ error: "that version is gone" }, 404, origin);
        const st = JSON.parse(snap.json);
        // the restored bytes are identical to the snapshot, so carry its own summary across.
        // (Taking the figure from the client labelled the restore with the balance it replaced.)
        const out = await writeState(env, snap.json, who + " (restored #" + id + ")", {
          days: snap.days, payments: snap.payments, net: snap.net,
        });
        return json({ ok: true, state: st, ...out }, 200, origin);
      }

      return json({ error: "not found" }, 404, origin);
    } catch (e) {
      return json({ error: "server error", detail: String(e && e.message || e) }, 500, origin);
    }
  },
};
