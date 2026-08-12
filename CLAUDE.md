# lead-ledger-bob-olivia — Bob ↔ Olivia daily lead settlement

**Status: DELIVERED 2026-08-12.** ⭐ **Bob's follow-up the same day: "put it on a git so an excel is
not needed" → the LIVE WEB LEDGER is now the product**, the workbook is the fallback.
**Awaiting Bob on 4 assumptions** (all four are rate fields or one-line changes — nothing is blocked).

**Deliverable:** https://opheliaclarke.github.io/lead-ledger/ = **the app** (noindex + robots
disallow; public repo because Pages needs it on the free plan — same posture as `lead-split`,
flagged because it holds partner financials). Repo `opheliaclarke/lead-ledger`.
`/about.html` = the written explanation (this was `index.html` until the app replaced it).
`lead-ledger-bob-olivia.xlsx` = the same model as a file, 5 tabs, 300 pre-formulated rows.

## ⭐ THE WEB LEDGER (index.html) — one file, no build, no server, no dependencies

Four views (`#daily` `#summary` `#rates` `#guide`, hash-routed). Type date + six lead counts, the NET
lands in the last column. Saves to **`localStorage` key `leadLedger.v1`** on every keystroke.
Top bar always shows all-time outstanding. CSV export, JSON backup/restore, wipe (auto-downloads a
backup first), "load three example days".

⭐ **Money is exact integer arithmetic, not floats** — everything converts to **1/10000 units**
(`toU`), so `(b+c)*price` is an integer product and nothing drifts. This is *better* than the
workbook, which relies on `ROUND(...,6)` guards.
⭐ **Per-day price override kept from the sheet**: prices default from Rates, and the per-day columns
are behind a toggle (default off) so the table stays readable. An overridden price is highlighted.
⭐ **Responsive by card-flip**: `table.resp` becomes labelled cards under 980px via `td::before{
content:attr(data-l)}` — one markup, two layouts, no duplicate render path.
⚠ **Only wrappers holding a `.resp` table may drop their scroll box on mobile** (`.tw-resp`). My
first version killed `overflow-x` on *every* `.tw`, which made the guide's plain table push the page
sideways (464 > 390). Caught by the automated overflow check, not by eye.
⚠ **The blank starter row is not data.** `loadEx` filters with `isEmptyDay()` before asking to
confirm — otherwise a fresh ledger asks "add to the 1 already here?" about an empty placeholder.

🛑 **The one real limit, stated on the Rates view: it is one browser.** Not shared with Olivia, gone
if browser data is cleared. Backup/restore is the mitigation. **If Bob wants both of them typing into
one live ledger that is a CF Worker + D1** (we have the pattern: `ai-visibility-collect`, `deploy-bot`,
`fleetview` on the Osanix account) — offered, not built.

## Verification of the app — `verify_app.py`, 163 checks against the LIVE URL

Drives the deployed page with Playwright and compares **every displayed number against the same
exact-`Fraction` model that validated the workbook**, so app and spreadsheet are proved to agree.
Covers: the UI typing path (3 days typed into real inputs), the engine over **40 seeded days**
(9 deliberate edge fixtures + randomised) checking net/L/M per row + 4 summary figures + 7 lane
totals, the date window filtering **days and payments**, persistence across reload, JSON
backup→wipe→restore, CSV shape, and a render pass (4 views × desktop+mobile: overflow, contrast,
console errors). All green.
⚠ **patchright evaluates in an ISOLATED world** — `window.LEDGER` set by page script reads as
`undefined`, while DOM queries and `localStorage` work fine. Don't debug page globals with it.
⚠ **Wait for `#v-<view>.on` after a nav click.** Reading a summary cell straight after the click
returns the *boot-time* value — my first run reported `0.00` vs `158.00` and it was a test race, not
a bug. It only showed up in the typing test because the seeded tests reload (so boot values are
already correct) — a race that hides itself in 4 of 5 tests.
⚠ **Read a downloaded CSV with `newline=""`** or Python's universal newlines eat the `\r\n` and a
CRLF assertion silently reads one line.

⚠ **This is NOT lead-split.** `lead-split` is Tyson & Berry, a **50/50 partnership**
(`settlement = share of net − collected + paid`, zero-sum). This is Bob & Olivia, a **buy/sell chain**
— a running account, not a profit split. Do not carry one model into the other. (Coincidence to note:
"Olivia" is also a buyer column in the Tyson/Berry sheet. Unconfirmed whether it's the same person.)

## ⭐ THE MODEL — six lanes, three of them create a debt

A debt appears **only when "who paid for the lead" and "who collects the money" come apart.**

| Source | Sold to | Paid for by | Collected by | Effect |
|---|---|---|---|---|
| Bob's own campaign | Olivia | Bob (ad spend) | Olivia buys it | **Olivia owes Bob** × `P_BO` |
| S | Olivia | Bob (pays S) | Olivia buys it | **Olivia owes Bob** × `P_BO` |
| GS | Olivia | Olivia (pays GS) | Olivia's buyer | nothing |
| Bob's own campaign | GLU (Bob's buyer) | Bob | Bob, from GLU | nothing + panel fee |
| S | GLU (Bob's buyer) | Bob (pays S) | Bob, from GLU | nothing + panel fee |
| GS | GLU (Bob's buyer) | **Olivia** (pays GS) | **Bob**, from GLU | **Bob owes Olivia** × `P_GSX` |

**`NET = (B+C)·P_BO − [ G·P_GSX + (E+F+G)·P_PANEL ] ± other cost`**, signed **positive = Olivia pays Bob**.
⭐ Column **D (GS → Olivia) appears in NO money formula** — that lane is entirely Olivia's, both ends.
Verified by a 99-lead GS→Olivia fixture producing 0.00.

Daily nets are **additive** — settling daily, weekly or monthly gives the identical figure.

## Layout

`Daily` A date · B–G the six lane counts · H total · **I/J/K the three prices** · L Olivia owes Bob ·
M Bob owes Olivia · N/O/P other cost (amount / paid by / YES-NO) · Q adjustment · R plain-English ·
**S = NET, the last column.** Yellow = type here, grey = computed, green = the answer.
`Rates` C6 `P_BO` · C7 `P_GSX` (defaults `=C15`, the GS cost) · C8 `P_PANEL` (0). C13–C16 reference only.
`Summary` date window → lane totals, money, NET, 4 checks, **payments-already-made log**, STILL OUTSTANDING.

⭐ **Prices live in COLUMNS I/J/K, defaulted from Rates, never hardcoded in a formula.** Direct fix for
the lead-split failure where the price went $8→$9 and was buried in four formulas per row. A per-day
override is legal and doesn't rewrite history.

⭐ **Blank-date rows return `""`, not 0.00** (`IF($A3="","",…)` on all 9 computed columns) so 300
pre-built rows don't read as 300 settled days.

## ⭐ The checks — v1 was a tautology, do not regress

**First version caught 1 of 4 breakages.** It compared `ΣL−ΣM+ΣQ` against `ΣS` — but S *is* L−M+Q on
each row, so damaging L moved both sides together. Same class of error as lead-split's `P2+Q2−M2`.
**Replaced with four checks that each rebuild a figure by a genuinely different route:**

1. **Lead money** — `SUMPRODUCT` of counts × prices, never reading L or M. Catches L/M damage.
2. **Other cost** — re-derives Q by `SUMIFS` on N/O/P.
3. **NET column** — catches a number typed straight into S.
4. **Missing price** — counts dated rows where the price column is empty.

**6/6 deliberate breakages caught, and the residual equals the size of the error** (80.00, 68.00,
−162.50, −512.50…). H (total leads) is display-only — damaging it moves no settlement, correctly silent.

⚠ **Measured blind spots — 0 does NOT mean the numbers are right:** a lead count typed wrong
(**+7,992**), leads in the wrong lane (**−253.50**), the wrong payer named (**−80.00**), a price typed
wrong on Rates (**+64,368**). All silent. Only the panel's own report validates the data.

⚠ **`SUMPRODUCT(a*b, c)` (comma form) treats `""` as 0; `SUMPRODUCT(a*b*c)` (all-`*`) returns
`#VALUE!`.** Proven here — the blank-row guard puts `""` in the price columns, so check 1 only works in
the comma form.

## Verification

`verify.py` — builds a 30-row copy (10 deliberate edge fixtures + randomised), evaluates the **real
workbook formulas** with the `formulas` package (a genuine Excel engine, `pip install formulas`), and
compares against an **exact-`Fraction` model written from the spec, not from the sheet**.
**302 cell checks, all pass**, covering: every row's H/L/M/Q/R/S + the three prices, the blank-row
guard on 9 columns, all 6 lane totals, the summary money block, all 4 checks reading 0, the payments
block, STILL OUTSTANDING, the plain-English sentence, and the **date window** (first-10-days totals
*and* the windowed payments). Panel fee set to **0.50 in the test** so that term is really exercised.
`verify_checks_fire.py` — the adversarial pass above.
⚠ The `formulas` solution key is `'[filename.xlsx]SHEETNAME'!REF` — **filename keeps its case, sheet
name is upper-cased.** Uppercasing both gives `KeyError`.
⚠ A fixture that lands on a **zero** value proves nothing — my first "MISSED" was deleting `L6` on a
row whose L was legitimately 0.00.

## Worked example (on the `Example` tab, hardcoded so edits to Rates can't move it)

`P_BO` 8.00 · `P_GSX` 6.00. **10 Aug +248.00 · 11 Aug +190.00 · 12 Aug −280.00 · three days +158.00.**
Day 3 flips the direction (45 GS leads → GLU) — deliberately chosen so both signs are demonstrated.

## 🛑 Open — needs Bob (all four are rate cells or one-liners, nothing is blocked)

1. **`P_GSX` — what does Bob pay Olivia for a GS lead that reaches his private buyer?** The one real
   gap. Defaulted to **pass-through at Olivia's GS cost, no margin**.
2. **Panel fee?** Bob's private-buyer leads run through Olivia's panel. Set to **0**.
3. **"Other cost = YES" — full reimbursement or 50/50?** Built as **full**.
4. **Is GLU the only private buyer?** All of Bob's private buyers are bucketed together — the
   settlement only cares *whose* buyer it was. Per-buyer breakout = a wider sheet.

Also assumed, from his own words: Olivia's price is the same whether the lead came from Bob's
campaign or from S ("cost from Bob to Olivia is also same always").

## Not done / not offered

- **No margin or P&L columns** — he said "don't add any extra figures". The reference rates (Bob's own
  cost, S cost, GS cost, GLU price) are recorded on the Rates tab but feed nothing. Offer a margin line
  rather than adding it unasked. ⚠ In lead-split the unit economics (**leads cost $33.17, sold for
  $8**) were the real story, so this is worth offering once — not twice.
- Nothing was written to any Google Sheet (no write path exists — see [[no-google-sheets-write-access]]).

## Files

`index.html` (**the app**) · `about.html` (the write-up) · `build_sheet.py` (generator, the single
source of the workbook) · `verify.py` · `verify_checks_fire.py` · `verify_app.py` (live-page tests) ·
`lead-ledger-bob-olivia.xlsx` · screenshots `app-daily/-summary/-mobile-daily/-desktop/-mobile.png`,
`qa-desktop.png` / `qa-mobile.png`.
Page QA: 0 contrast failures (runtime walker compositing alpha), 0 console errors, no horizontal
overflow desktop + mobile. Live download md5 `eb70e420a4b8c297f0c47e646e4f0378` == local.
