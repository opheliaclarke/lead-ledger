#!/usr/bin/env python3
"""
End-to-end test of the server-backed ledger, against the LIVE page and the LIVE Worker.

Two independent browser contexts (Bob and Olivia, each with their own access code) drive
the real site. Checks sign-in, that a change made by one is picked up by the other, the
conflict guard when both edit, and that every earlier version can be put back.

Run after `wrangler deploy`. It wipes the server ledger first, so do not run it once real
data is in there.
"""
import asyncio
import json
import subprocess
import sys
from fractions import Fraction as Fr

try:
    from patchright.async_api import async_playwright
except ImportError:
    from playwright.async_api import async_playwright

URL = "https://opheliaclarke.github.io/lead-ledger/"
API = "https://lead-ledger.fleet-fefsba.workers.dev"
CFG = "/root/.config/lead-ledger/access.json"
WORKER_DIR = "/root/workspace/lead-ledger-bob-olivia/worker"
FAILS = []


def fail(m):
    FAILS.append(m)


def check(name, got, want):
    if str(got) != str(want):
        fail(f"{name}: got {got!r}, want {want!r}")


def codes():
    d = json.load(open(CFG))
    return d["codes"]["Bob"], d["codes"]["Olivia"]


def wipe_server():
    import os
    cf = json.load(open("/root/.config/cloudflare/osanix-fleetview.json"))
    env = {**os.environ, "CLOUDFLARE_API_KEY": cf["api_key"],
           "CLOUDFLARE_EMAIL": cf["email"], "CLOUDFLARE_ACCOUNT_ID": cf["account_id"]}
    for sql in ["DELETE FROM state", "DELETE FROM snapshots"]:
        subprocess.run(["npx", "wrangler", "d1", "execute", "lead_ledger", "--remote",
                        "--config", "./wrangler.json", "--command", sql],
                       cwd=WORKER_DIR, env=env, capture_output=True, timeout=180)


async def sign_in(page, code, expect_name):
    await page.goto(URL, wait_until="networkidle")
    await page.wait_for_selector("#gate:not([hidden])", timeout=15000)
    await page.fill("#gateCode", code)
    await page.click("#gateGo")
    await page.wait_for_selector("#gate", state="hidden", timeout=20000)
    check(f"sign-in as {expect_name} — name in top bar",
          (await page.text_content("#whoAmI")).strip(), expect_name)


async def sync_state(page):
    return (await page.text_content("#syncTxt")).strip()


async def wait_saved(page, timeout=25000):
    await page.wait_for_function(
        "() => /^saved/.test(document.getElementById('syncTxt').textContent)",
        timeout=timeout)


async def force_poll(page):
    # the page listens for visibilitychange on the shared DOM, so this works even
    # from patchright's isolated world
    await page.evaluate("document.dispatchEvent(new Event('visibilitychange'))")
    await page.wait_for_timeout(1500)


async def goview(page, v):
    await page.click(f"nav a[data-v='{v}']")
    await page.wait_for_selector(f"#v-{v}.on", state="visible")
    await page.wait_for_timeout(200)


async def txt(page, sel):
    return (await page.text_content(sel)).strip()


def api_curl(path, code, body):
    """Cloudflare 403s python-urllib; curl gets through."""
    r = subprocess.run(["curl", "-s", "-X", "POST", API + path,
                        "-H", f"Authorization: Bearer {code}",
                        "-H", "Content-Type: application/json",
                        "--data-binary", json.dumps(body)],
                       capture_output=True, text=True, timeout=60)
    return r.stdout


async def main():
    bob_code, oli_code = codes()
    print("  wiping the server ledger for a clean run…")
    wipe_server()

    async with async_playwright() as pw:
        br = await pw.chromium.launch(args=["--no-sandbox"])
        bob_ctx = await br.new_context(viewport={"width": 1500, "height": 950})
        oli_ctx = await br.new_context(viewport={"width": 1500, "height": 950})
        bob = await bob_ctx.new_page()
        oli = await oli_ctx.new_page()
        for p in (bob, oli):
            p.on("dialog", lambda d: asyncio.ensure_future(d.accept()))
            p.on("pageerror", lambda e: fail(f"page error: {e}"))

        # ---- 1. a wrong code is refused ------------------------------------
        await bob.goto(URL, wait_until="networkidle")
        await bob.wait_for_selector("#gate:not([hidden])", timeout=15000)
        await bob.fill("#gateCode", "bob-notarealcode")
        await bob.click("#gateGo")
        await bob.wait_for_selector("#gateErr:not([hidden])", timeout=20000)
        err = await txt(bob, "#gateErr")
        check("wrong code — refused", "not recognised" in err, True)
        check("wrong code — gate stays shut",
              await bob.is_hidden("#gate"), False)
        print(f"  wrong code refused: {err!r}")

        # ---- 2. Bob signs in and puts data in ------------------------------
        await sign_in(bob, bob_code, "Bob")
        await goview(bob, "guide")
        await bob.click("#loadEx")
        await bob.wait_for_timeout(900)
        await goview(bob, "daily")
        nets = await bob.eval_on_selector_all(
            "#daysBody tr td[data-c='net']",
            "e => e.map(x => x.firstChild.textContent.trim())")
        check("Bob — three day nets", nets, ["+248.00", "+190.00", "−280.00"])
        await wait_saved(bob)
        print(f"  Bob signed in, entered 3 days, sync says {(await sync_state(bob))!r}")

        # ---- 3. Olivia opens and sees Bob's ledger -------------------------
        await sign_in(oli, oli_code, "Olivia")
        await goview(oli, "daily")
        onets = await oli.eval_on_selector_all(
            "#daysBody tr td[data-c='net']",
            "e => e.map(x => x.firstChild.textContent.trim())")
        check("Olivia — sees Bob's three days", onets, nets)
        check("Olivia — top balance", await txt(oli, "#topBal"), "+158.00")
        print(f"  Olivia opened on a different browser and sees the same 3 days, +158.00")

        # ---- 4. Olivia logs a payment; Bob picks it up ---------------------
        await goview(oli, "summary")
        await oli.click("#addPay")
        await oli.wait_for_timeout(300)
        row = oli.locator("#payBody tr").first
        await row.locator("[data-k='amt']").fill("100")
        await row.locator("[data-k='by']").select_option("Olivia")
        await oli.wait_for_timeout(400)
        await wait_saved(oli)
        check("Olivia — outstanding after her 100 payment", await txt(oli, "#topBal"), "+58.00")

        await force_poll(bob)
        check("Bob — picked up Olivia's payment", await txt(bob, "#topBal"), "+58.00")
        print("  Olivia logged a 100.00 payment; Bob's tab updated to +58.00 on its own")

        # ---- 5. conflict: the server moves while Bob is mid-edit -----------
        # NB: curl, not urllib — Cloudflare 403s python-urllib's user agent.
        cur = json.loads(api_curl("/api/load", oli_code, {}))
        state = cur["state"]
        state["days"][0]["b"] = "41"          # Olivia changes a lead count behind Bob's back
        api_curl("/api/save", oli_code, {"state": state, "baseRev": cur["rev"],
                                         "meta": {"days": 3, "payments": 1, "net": "+66.00"}})

        await goview(bob, "daily")
        await bob.locator("#daysBody tr").first.locator("[data-k='c']").fill("7")
        await bob.wait_for_selector("#conflict:not([hidden])", timeout=25000)
        ctext = await txt(bob, "#conflictTxt")
        check("conflict — names who saved", "Olivia" in ctext, True)
        check("conflict — sync chip", await sync_state(bob), "not saved")
        print(f"  conflict caught: {ctext[:64]}…")

        # Bob takes their version
        await bob.click("#cfTheirs")
        await bob.wait_for_selector("#conflict", state="hidden", timeout=15000)
        b0 = await bob.locator("#daysBody tr").first.locator("[data-k='b']").input_value()
        c0 = await bob.locator("#daysBody tr").first.locator("[data-k='c']").input_value()
        check("conflict — took Olivia's 41", b0, "41")
        check("conflict — Bob's unsaved 7 discarded", c0, "")
        print("  Bob chose their version: lead count is Olivia's 41, his unsaved edit dropped")

        # ---- 6. history and restore ---------------------------------------
        await goview(bob, "backup")
        await bob.wait_for_function(
            "() => document.querySelectorAll('#histBody tr').length > 1", timeout=25000)
        rows = await bob.eval_on_selector_all(
            "#histBody tr", "e => e.map(r => r.children[1].textContent.trim())")
        check("history — one version per real save", len(rows), 3)
        check("history — records both names", set(rows) >= {"Bob", "Olivia"}, True)
        print(f"  backup list: {len(rows)} versions, saved by {sorted(set(rows))}")
        n_versions = len(rows)
        oldest_net = await bob.eval_on_selector(
            "#histBody tr:last-child td:nth-child(5)", "e => e.textContent.trim()")

        # restore the oldest listed version (the 3-day, no-payment one)
        n_before = await txt(bob, "#topBal")
        await bob.click("#histBody tr:last-child [data-restore]")
        await bob.wait_for_timeout(2500)
        await goview(bob, "daily")
        after = await txt(bob, "#topBal")
        check("restore — balance changed back", after != n_before, True)
        check("restore — back to the first saved version", after, "+158.00")
        print(f"  restored the oldest version: outstanding went {n_before} -> {after}")

        # the restore is itself just another version, so it can be undone —
        # and it must add exactly ONE, or something is saving on its own
        await goview(bob, "backup")
        await bob.wait_for_timeout(1200)
        await bob.click("#histRefresh")
        await bob.wait_for_timeout(1800)
        rows2 = await bob.eval_on_selector_all(
            "#histBody tr", "e => e.map(r => r.children[1].textContent.trim())")
        check("restore — adds exactly one version, nothing saves by itself",
              len(rows2), n_versions + 1)
        check("restore — logged as its own version", "restored" in rows2[0], True)
        top_net = await bob.eval_on_selector("#histBody tr:first-child td:nth-child(5)",
                                             "e => e.textContent.trim()")
        check("restore — labelled with the balance it restored, not the one it replaced",
              top_net, oldest_net)
        print(f"  the restore is version {rows2[0]!r} labelled {top_net} — undoable")

        # ---- 7. Olivia sees the restore ------------------------------------
        await force_poll(oli)
        check("Olivia — sees the restored ledger", await txt(oli, "#topBal"), "+158.00")
        print("  Olivia's tab followed the restore")

        await bob_ctx.close()
        await oli_ctx.close()
        await br.close()

    print()
    if FAILS:
        print(f"FAIL — {len(FAILS)} problem(s):")
        for f in FAILS[:20]:
            print("   ", f)
        sys.exit(1)
    print("PASS — sign-in, sharing, conflict guard and version restore all work live.")


if __name__ == "__main__":
    asyncio.run(main())
