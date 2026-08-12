#!/usr/bin/env python3
"""
Verify the LIVE web ledger at https://opheliaclarke.github.io/lead-ledger/

Every number the page shows is compared against the same exact-Fraction model that
validated the Excel workbook (verify.py), so the app and the workbook are proved to agree.

Covers: the UI typing path, the calculation engine over 40 randomised days, persistence
across a reload, the date window, payments and the outstanding balance, backup export ->
wipe -> restore, and a render pass (console errors, contrast, horizontal overflow).
"""
import asyncio
import json
import os
import random
import subprocess
import sys
from fractions import Fraction as Fr

try:
    from patchright.async_api import async_playwright
except ImportError:
    from playwright.async_api import async_playwright

URL = "https://opheliaclarke.github.io/lead-ledger/"
API = "https://lead-ledger.fleet-fefsba.workers.dev"
CODE = json.load(open("/root/.config/lead-ledger/access.json"))["codes"]["Bob"]
WORKER_DIR = "/root/workspace/lead-ledger-bob-olivia/worker"
FAILS = []


def api_curl(path, body):
    """Cloudflare 403s python-urllib; curl gets through."""
    r = subprocess.run(["curl", "-s", "-X", "POST", API + path,
                        "-H", f"Authorization: Bearer {CODE}",
                        "-H", "Content-Type: application/json",
                        "--data-binary", json.dumps(body)],
                       capture_output=True, text=True, timeout=60)
    return r.stdout


def wipe_server():
    cf = json.load(open("/root/.config/cloudflare/osanix-fleetview.json"))
    env = {**os.environ, "CLOUDFLARE_API_KEY": cf["api_key"],
           "CLOUDFLARE_EMAIL": cf["email"], "CLOUDFLARE_ACCOUNT_ID": cf["account_id"]}
    for sql in ["DELETE FROM state", "DELETE FROM snapshots"]:
        subprocess.run(["npx", "wrangler", "d1", "execute", "lead_ledger", "--remote",
                        "--config", "./wrangler.json", "--command", sql],
                       cwd=WORKER_DIR, env=env, capture_output=True, timeout=180)


async def sign_in(page):
    await page.goto(URL, wait_until="networkidle")
    try:
        await page.wait_for_selector("#gate:not([hidden])", timeout=8000)
    except Exception:
        return                     # already signed in on this context
    await page.fill("#gateCode", CODE)
    await page.click("#gateGo")
    await page.wait_for_selector("#gate", state="hidden", timeout=25000)


async def wait_saved(page, timeout=25000):
    await page.wait_for_function(
        "() => /^saved/.test(document.getElementById('syncTxt').textContent)", timeout=timeout)


def fail(msg):
    FAILS.append(msg)


def check(name, got, want):
    if str(got) != str(want):
        fail(f"{name}: got {got!r}, want {want!r}")


# ---------------------------------------------------------------- model
def model_day(counts, p_bo, p_gs, p_pf, amt, by, settle):
    b, c, d, e, f, g = (Fr(x) for x in counts)
    L = (b + c) * Fr(str(p_bo))
    M = g * Fr(str(p_gs)) + (e + f + g) * Fr(str(p_pf))
    Q = Fr(0)
    if settle == "YES" and amt not in ("", None):
        Q = Fr(str(amt)) if by == "Bob" else -Fr(str(amt))
    return dict(leads=b + c + d + e + f + g, L=L, M=M, Q=Q, net=L - M + Q)


def money(fr):
    """Format a Fraction the way the page does: en-US, 2dp, comma thousands."""
    v = float(fr)
    return f"{abs(v):,.2f}"


def signed(fr):
    v = float(fr)
    if v > 0:
        return "+" + f"{v:,.2f}"
    if v < 0:
        return "−" + f"{abs(v):,.2f}"
    return "0.00"


# ---------------------------------------------------------------- helpers
async def seed(page, rates, days, payments=None, win=None):
    """Write state straight into localStorage, then reload. Exercises the engine."""
    state = {
        "rates": rates,
        "days": [dict(id=f"d{i}", date=d["date"],
                      b=str(d["counts"][0]), c=str(d["counts"][1]), d=str(d["counts"][2]),
                      e=str(d["counts"][3]), f=str(d["counts"][4]), g=str(d["counts"][5]),
                      pbo=d.get("pbo", ""), pgsx=d.get("pgsx", ""), ppanel=d.get("ppanel", ""),
                      amt=d.get("amt", ""), by=d.get("by", "Bob"), settle=d.get("settle", "NO"))
                 for i, d in enumerate(days)],
        "payments": [dict(id=f"p{i}", date=p["date"], by=p["by"], amt=str(p["amt"]), note="")
                     for i, p in enumerate(payments or [])],
        "win": win or {"from": "", "to": ""},
        "showPrices": False,
    }
    cur = json.loads(api_curl("/api/load", {}) or "{}")
    api_curl("/api/save", {"state": state, "baseRev": cur.get("rev", 0),
                           "meta": {"days": len(state["days"]),
                                    "payments": len(state["payments"]), "net": "seed"}})
    # drop the local copy so the page adopts the server version instead of flagging a
    # conflict between the two (which is the correct behaviour when both have data)
    await page.evaluate("localStorage.removeItem('leadLedger.v1');"
                        "localStorage.removeItem('leadLedger.rev')")
    await page.reload(wait_until="networkidle")
    await page.wait_for_function(
        "n => document.querySelectorAll('#daysBody tr').length === n",
        arg=len(state["days"]), timeout=25000)


async def row_nets(page):
    return await page.eval_on_selector_all(
        "#daysBody tr td[data-c='net']",
        "els => els.map(e => e.firstChild.textContent.trim())")


async def row_dirs(page):
    return await page.eval_on_selector_all(
        "#daysBody tr td[data-c='net'] .dir", "els => els.map(e => e.textContent.trim())")


async def txt(page, sel):
    return (await page.text_content(sel)).strip()


async def goview(page, v):
    """Click a nav link and wait for that view to actually be rendered."""
    await page.click(f"nav a[data-v='{v}']")
    await page.wait_for_selector(f"#v-{v}.on", state="visible")
    await page.wait_for_timeout(120)


# ---------------------------------------------------------------- tests
async def t_ui_typing(page):
    """Type into the real inputs and read the answer back."""
    wipe_server()
    await page.goto(URL, wait_until="networkidle")
    await page.evaluate("localStorage.clear()")
    await sign_in(page)

    # set the rates through the UI
    await goview(page, "rates")
    for k, v in [("pbo", "8"), ("pgsx", "6"), ("ppanel", "0")]:
        await page.fill(f"#r-{k}", v)
    await goview(page, "daily")

    # one blank day exists on a fresh ledger; fill it
    cases = [
        ([40, 0, 25, 10, 0, 12], "", "Bob", "NO"),
        ([0, 35, 18, 0, 8, 20], "30", "Bob", "YES"),
        ([5, 0, 30, 2, 0, 45], "50", "Olivia", "YES"),
    ]
    for i, (counts, amt, by, settle) in enumerate(cases):
        if i > 0:
            await page.click("#addDay")
        rows = page.locator("#daysBody tr")
        r = rows.nth(i)
        for k, v in zip("bcdefg", counts):
            await r.locator(f"[data-k='{k}']").fill(str(v))
        if amt:
            await r.locator("[data-k='amt']").fill(amt)
            await r.locator("[data-k='by']").select_option(by)
            await r.locator("[data-k='settle']").select_option(settle)

    nets = await row_nets(page)
    dirs = await row_dirs(page)
    want = []
    for counts, amt, by, settle in cases:
        m = model_day(counts, 8, 6, 0, amt or None, by, settle)
        want.append(signed(m["net"]))
    check("UI typing — three day nets", nets, want)
    check("UI typing — directions", dirs, ["Olivia → Bob", "Olivia → Bob", "Bob → Olivia"])

    await goview(page, "summary")
    total = sum(model_day(c, 8, 6, 0, a or None, b, s)["net"] for c, a, b, s in cases)
    check("UI typing — period net", await txt(page, "#sNet"), signed(total))
    check("UI typing — answer", await txt(page, "#ansBig"), money(total))
    await wait_saved(page)
    print(f"  UI typing path: 3 days entered by hand -> "
          f"{', '.join(nets)}  period {signed(total)}")
    return len(cases) * 2 + 2


async def t_engine(page, rng):
    """40 randomised days straight into the engine, including edge cases."""
    rates = dict(pbo=8.25, pgsx=6.5, ppanel=0.5, bobcost=0, scost=0, gscost=6.5, gluprice=0)
    fixed = [
        ([0, 0, 0, 0, 0, 0], "", "Bob", "NO"),          # nothing at all
        ([10, 0, 0, 0, 0, 0], "", "Bob", "NO"),         # only Bob's own to Olivia
        ([0, 0, 99, 0, 0, 0], "", "Bob", "NO"),         # only GS to Olivia -> must be 0
        ([0, 0, 0, 7, 4, 0], "", "Bob", "NO"),          # only Bob's leads to his own buyer
        ([0, 0, 0, 0, 0, 25], "", "Bob", "NO"),         # only GS to GLU -> Bob owes
        ([3, 2, 1, 1, 1, 1], "40", "Bob", "YES"),
        ([3, 2, 1, 1, 1, 1], "40", "Olivia", "YES"),
        ([3, 2, 1, 1, 1, 1], "40", "Bob", "NO"),        # logged, not settled
        ([0, 0, 0, 0, 0, 0], "12.5", "Olivia", "YES"),  # cost only, no leads
    ]
    days, want = [], []
    for i in range(40):
        if i < len(fixed):
            counts, amt, by, settle = fixed[i]
        else:
            counts = [rng.choice([0, 0, rng.randint(1, 150)]) for _ in range(6)]
            amt = f"{rng.uniform(5, 300):.2f}" if rng.random() < .4 else ""
            by = rng.choice(["Bob", "Olivia"])
            settle = rng.choice(["YES", "NO"])
        date = f"2026-09-{i+1:02d}" if i < 30 else f"2026-10-{i-29:02d}"
        days.append(dict(date=date, counts=counts, amt=amt, by=by, settle=settle))
        want.append(model_day(counts, rates["pbo"], rates["pgsx"], rates["ppanel"],
                              amt or None, by, settle))

    await seed(page, rates, days)
    nets = await row_nets(page)
    check("engine — row count", len(nets), len(days))
    bad = [i for i, (g, w) in enumerate(zip(nets, want)) if g != signed(w["net"])]
    if bad:
        fail(f"engine — {len(bad)} row net mismatches, first at row {bad[0]}: "
             f"{nets[bad[0]]!r} vs {signed(want[bad[0]]['net'])!r}")

    Ls = await page.eval_on_selector_all("#daysBody tr td[data-c='L']",
                                         "els => els.map(e=>e.textContent.trim())")
    Ms = await page.eval_on_selector_all("#daysBody tr td[data-c='M']",
                                         "els => els.map(e=>e.textContent.trim())")
    for i, w in enumerate(want):
        check(f"engine — row {i} 'Olivia owes Bob'", Ls[i], money(w["L"]))
        check(f"engine — row {i} 'Bob owes Olivia'", Ms[i], money(w["M"]))

    await goview(page, "summary")
    tot = dict(L=sum(w["L"] for w in want), M=sum(w["M"] for w in want),
               Q=sum(w["Q"] for w in want), net=sum(w["net"] for w in want))
    check("engine — summary Olivia owes Bob", await txt(page, "#sL"), money(tot["L"]))
    check("engine — summary Bob owes Olivia", await txt(page, "#sM"), money(tot["M"]))
    check("engine — summary other cost", await txt(page, "#sQ"), signed(tot["Q"]))
    check("engine — summary net", await txt(page, "#sNet"), signed(tot["net"]))

    lanes = await page.eval_on_selector_all("#laneBody tr td:nth-child(2)",
                                            "els => els.map(e=>e.textContent.trim())")
    for i in range(6):
        check(f"engine — lane {i} total", lanes[i], str(sum(d["counts"][i] for d in days)))
    check("engine — total leads", await txt(page, "#laneTotal"),
          str(sum(sum(d["counts"]) for d in days)))
    print(f"  engine: {len(days)} days x (net, L, M) + 4 summary figures + 7 lane totals")
    return len(days) * 3 + 11, rates, days, want


async def t_window(page, rates, days, want):
    """The date window must filter days AND payments."""
    first10 = [w for d, w in zip(days, want) if d["date"] <= "2026-09-10"]
    pays = [dict(date="2026-09-05", by="Olivia", amt=500),
            dict(date="2026-09-09", by="Bob", amt=120),
            dict(date="2026-10-02", by="Olivia", amt=75)]
    await seed(page, rates, days, pays, win={"from": "2026-09-01", "to": "2026-09-10"})
    await goview(page, "summary")

    net10 = sum(w["net"] for w in first10)
    check("window — net over first 10 days", await txt(page, "#sNet"), signed(net10))
    check("window — paid by Olivia in window", await txt(page, "#pO"), money(Fr(500)))
    check("window — paid by Bob in window", await txt(page, "#pB"), money(Fr(120)))
    outstanding = net10 - Fr(500) + Fr(120)
    check("window — outstanding", await txt(page, "#ansBig"), money(outstanding))
    check("window — day count", await txt(page, "#winCount"),
          f"{len(first10)} days in range · {sum(int(w['leads']) for w in first10)} leads")

    # all time
    await page.click("#winAll")
    allnet = sum(w["net"] for w in want)
    allout = allnet - Fr(500) - Fr(75) + Fr(120)
    check("all time — net", await txt(page, "#sNet"), signed(allnet))
    check("all time — outstanding", await txt(page, "#ansBig"), money(allout))
    top = await txt(page, "#topBal")
    check("top bar — outstanding matches", top, signed(allout))
    print(f"  window + payments: windowed net {signed(net10)}, "
          f"all-time outstanding {signed(allout)} (top bar agrees)")
    return 8


async def t_persist(page):
    before = await row_nets(page)
    await page.reload(wait_until="networkidle")
    after = await row_nets(page)
    check("persistence — rows survive a reload", after, before)
    print(f"  persistence: {len(after)} rows identical after reload")
    return 1


async def t_backup(page, tmp):
    """Export a backup and a CSV, change the ledger, then restore from the file."""
    before_top = await txt(page, "#topBal")
    n_before = len(await row_nets(page))

    await goview(page, "backup")
    async with page.expect_download() as dl:
        await page.click("#expJson")
    path = await (await dl.value).path()
    blob = json.load(open(path))
    check("backup — day count in file", len(blob["data"]["days"]), n_before)

    async with page.expect_download() as dl2:
        await page.click("#expCsv")
    csv = open(await (await dl2.value).path(), newline="").read()
    lines = csv.split("\r\n")
    check("backup — CSV is CRLF", csv.count("\r\n") > 0, True)
    check("backup — CSV header", lines[0].split(",")[0], "Date")
    check("backup — CSV columns", len(lines[0].split(",")), 19)
    body = lines[1:lines.index("")] if "" in lines else lines[1:]
    check("backup — CSV data rows", len(body), n_before)

    # change the ledger, then put the file back
    await goview(page, "daily")
    await page.locator("#daysBody tr").first.locator("[data-k='b']").fill("777")
    await page.wait_for_timeout(400)
    changed = await txt(page, "#topBal")
    check("backup — the ledger really changed", changed != before_top, True)
    await wait_saved(page)

    await goview(page, "backup")
    await page.set_input_files("#impFile", path)
    await page.wait_for_timeout(900)
    await goview(page, "daily")
    check("restore from file — balance back to what it was", await txt(page, "#topBal"), before_top)
    check("restore from file — day count back", len(await row_nets(page)), n_before)
    print(f"  backup: JSON round-trip restored {n_before} days, CSV has 19 columns")
    return 8


async def t_render(page, ctx):
    errs = []
    for name, (w, h) in {"desktop": (1500, 950), "mobile": (390, 844)}.items():
        p = await ctx.new_page()
        p.on("console", lambda m: errs.append(m.text) if m.type == "error" else None)
        p.on("pageerror", lambda e: errs.append(str(e)))
        await sign_in(p)
        await p.set_viewport_size({"width": w, "height": h})
        for view in ["daily", "summary", "rates", "backup", "guide"]:
            await goview(p, view)
            ow = await p.evaluate("document.documentElement.scrollWidth")
            iw = await p.evaluate("window.innerWidth")
            if ow > iw + 1:
                fail(f"render — horizontal overflow on {name}/{view}: {ow} > {iw}")
        bad = await p.evaluate(CONTRAST_JS)
        if bad:
            fail(f"render — {len(bad)} contrast failures on {name}: {bad[:4]}")
        await p.screenshot(path=f"app-{name}.png", full_page=(name == "desktop"))
        await p.close()
        print(f"  render {name}: 5 views, no overflow, {len(bad)} contrast failures")
    if errs:
        fail(f"render — console errors: {errs[:5]}")
    return 9


CONTRAST_JS = """() => {
  const lum=c=>{const s=c.map(v=>{v/=255;return v<=0.03928?v/12.92:Math.pow((v+0.055)/1.055,2.4)});
    return 0.2126*s[0]+0.7152*s[1]+0.0722*s[2]};
  const parse=s=>{const m=String(s).match(/[\\d.]+/g);
    return m?m.slice(0,3).map(Number).concat([m[3]===undefined?1:+m[3]]):null};
  const comp=(f,b)=>f.slice(0,3).map((v,i)=>v*f[3]+b[i]*(1-f[3]));
  const bgOf=el=>{let n=el;const st=[];
    while(n&&n!==document.documentElement){const c=parse(getComputedStyle(n).backgroundColor);
      if(c&&c[3]>0)st.push(c); n=n.parentElement;}
    let base=[255,255,255]; for(let i=st.length-1;i>=0;i--) base=comp(st[i],base); return base;};
  const vis=el=>{const r=el.getBoundingClientRect();
    if(r.width===0&&r.height===0) return false;
    let n=el; while(n&&n!==document.documentElement){const cs=getComputedStyle(n);
      if(cs.display==='none'||cs.visibility==='hidden'||+cs.opacity===0) return false;
      if(n.hasAttribute&&n.hasAttribute('hidden')) return false; n=n.parentElement;} return true;};
  const out=[];
  document.querySelectorAll('*').forEach(el=>{
    const t=[...el.childNodes].filter(n=>n.nodeType===3&&n.textContent.trim())
                              .map(n=>n.textContent.trim()).join(' ');
    if(!t||!vis(el)) return;
    const cs=getComputedStyle(el); const fg=parse(cs.color); if(!fg) return;
    const bg=bgOf(el); const f=comp(fg,bg);
    const l1=lum(f), l2=lum(bg);
    const r=(Math.max(l1,l2)+0.05)/(Math.min(l1,l2)+0.05);
    const size=parseFloat(cs.fontSize), bold=+cs.fontWeight>=700;
    const need=(size>=24||(size>=18.66&&bold))?3:4.5;
    if(r<need-0.005) out.push({t:t.slice(0,40), r:+r.toFixed(2), need, size});
  });
  return out;
}"""


async def main():
    rng = random.Random(20260812)
    total = 0
    async with async_playwright() as pw:
        b = await pw.chromium.launch(args=["--no-sandbox"])
        ctx = await b.new_context(accept_downloads=True, viewport={"width": 1500, "height": 950})
        page = await ctx.new_page()
        page.on("pageerror", lambda e: fail(f"page error: {e}"))
        page.on("dialog", lambda d: asyncio.ensure_future(d.accept()))

        total += await t_ui_typing(page)
        n, rates, days, want = await t_engine(page, rng)
        total += n
        total += await t_window(page, rates, days, want)
        total += await t_persist(page)
        total += await t_backup(page, "/tmp")
        total += await t_render(page, ctx)

        await ctx.close()
        await b.close()

    print()
    if FAILS:
        print(f"FAIL — {len(FAILS)} problem(s):")
        for f in FAILS[:25]:
            print("   ", f)
        sys.exit(1)
    print(f"PASS — {total} checks against the live page, all matching the exact model.")


if __name__ == "__main__":
    asyncio.run(main())
