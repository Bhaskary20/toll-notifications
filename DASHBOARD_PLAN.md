# Toll Operations Desk — Premium Dashboard Plan

> **Status (15 Oct 2026):** steps 7.0–7.9 are built and checked on the mock bench (`py tools/mock_server.py --shots`).
> Decisions taken (section 9): headline = **pending**, stale after **7 days**, **dark only**.
> Still to do: 7.10 real run against the portal (connect, captcha, each job), and 7.11 `py build_exe.py`.
> Deviations from this plan: the KPI band reveals a brass top line on hover instead of lifting 1px (cells share one
> frame); earlier history snapshots hold totals, so the first pending delta reads "first pending snapshot".
>
> **Superseded look (15 Oct 2026, evening):** the page was redesigned twice more at the user's request, so sections
> 1–6 below describe the look only historically. The current page is a "classic" dark-and-gold design:
> Cormorant Garamond serif for headings and numbers, Manrope for text, a labelled sidebar (pages 1–6, session card,
> Motion switch, shortcuts), a hero with a greeting and a guilloché seal showing tasks done this week, a "This week"
> row with one tile per job (status from the saved results or the newest file in `downloads\`), then the KPI cards,
> trend, attention, files and activity. Sign-in gates appear on pages that need the portal.
> Motion: a one-time intro per browser session, gold silk lines on a canvas, cards that rise into view, a cursor
> light, a travelling border on running tasks and gold sparks when a task finishes. The sidebar Motion switch turns it
> all off (stored as `calm` in localStorage). The page no longer follows the Windows "reduce motion" setting, because
> Windows Server turns animation effects off by default and that hid every animation.

A step-by-step plan to rebuild `static/index.html` into a dark, quiet, premium dashboard with real KPI cards at the
top. Each phase is small enough to finish, check and screenshot before the next one starts.

---

## 0. Where we are and what is wrong

The current redesign works, but it looks like a **template**:

| Looks generic because of… | Replace with… |
|---|---|
| Orange → coral gradient buttons and glow shadows | One flat **brass** accent, no glow, pressed-metal feel |
| Glassmorphism cards + drifting colour blobs | Solid warm-black surfaces, **hairline** borders, one slow light sweep |
| Geist / Inter type everywhere | **Instrument Serif** numerals + **Hanken Grotesk** text + **JetBrains Mono** meta |
| "Good afternoon, GMC" hero banner | No hero. **The KPI band is the first thing on screen** |
| An icon in every card, 18px radius on everything | Icons only where they carry meaning, 6–10px radii |
| Count-up numbers, spinning loaders | **Odometer digit roll**, brass hairline progress, breathing status dot |
| Same visual weight on every card | Clear hierarchy: 4 primary KPIs → secondary strip → work area |

**Target feeling:** the instrument cluster of a luxury car at night, or a watch dial. Calm, dark and exact. Motion
is slow and soft, and nothing flashes.

---

## 1. Design rules (read before every phase)

1. **Data first.** The KPI band sits above the fold on a 1366×768 laptop.
2. **One accent.** Brass (`--brass`) for the active item, the main action and focus. Status colours appear only on
   status, always with a word next to them.
3. **Hairlines over boxes.** Separate with 1px lines at 6–8% white. Avoid filled panels inside panels.
4. **Big numbers, small labels.** KPI value 56–64px serif. Labels 11px mono, uppercase, letter-spacing .14em.
5. **Slow motion.** 600–1200ms, ease `cubic-bezier(.16,1,.3,1)`. Nothing bounces, and nothing loops faster than 4s.
6. **No decoration without data.** Every line, ring or bar on a card encodes a real number.
7. **Respect `prefers-reduced-motion`.** Everything still reads correctly with all animation off.
8. **Stays self-contained.** One HTML file with no JS libraries. Fonts come from Google Fonts with system fallbacks
   (some government networks block CDNs).

### Banned list
Gradient buttons · glow shadows · glass blur on cards · colour blobs · emoji · "AI sparkle" icons · rainbow charts ·
dual-axis charts · numbers in series colours · toasts that bounce · 3D tilt on hover.

---

## 2. Design tokens

Put these in `:root` and use only these values. Light theme is out of scope for v1 because the brief is dark premium.
Keep the toggle out until v1 ships.

### Colour

```css
:root{
  /* surfaces: warm black, not blue-black */
  --bg:        #0b0a09;   /* page */
  --surface-1: #121110;   /* cards */
  --surface-2: #181715;   /* raised: drawer, modal, tooltip */
  --surface-3: #201e1b;   /* hover row, input */
  --hair:      rgba(255,244,228,.07);  /* hairline */
  --hair-2:    rgba(255,244,228,.13);  /* hover hairline / divider */

  /* ink: warm off-white */
  --ink:   #f2ede4;
  --ink-2: #b4ada1;
  --ink-3: #7a746a;   /* labels, axis */

  /* the single accent */
  --brass:      #c9a86a;
  --brass-dim:  rgba(201,168,106,.14);
  --brass-line: rgba(201,168,106,.45);

  /* status (muted, always with text) */
  --ok:   #7fb89a;
  --warn: #d9b26a;
  --bad:  #d98579;

  /* chart series (validated, dataviz validator, dark, surface #0f0e0d: all checks PASS) */
  --s-fmc:    #b9862c;
  --s-bg:     #109884;
  --s-notice: #8a7cf0;
}
```

Re-run the validator whenever the series colours change:
`node <dataviz skill>/scripts/validate_palette.js "#b9862c,#109884,#8a7cf0" --mode dark --surface "#0f0e0d"`

### Type

| Role | Font | Size / weight | Notes |
|---|---|---|---|
| KPI value | Instrument Serif | 60px / 400, `font-variant-numeric: lining-nums tabular-nums` | the hero of every card |
| Page title | Instrument Serif | 34px / 400 | one per page |
| Section / card title | Hanken Grotesk | 14px / 600 | |
| Body | Hanken Grotesk | 14px / 400, line-height 1.55 | |
| Label / eyebrow | JetBrains Mono | 10.5–11px / 500, uppercase, `.14em` | |
| Data in tables | JetBrains Mono | 12.5px / 400 | |

Google Fonts URL:
`https://fonts.googleapis.com/css2?family=Instrument+Serif&family=Hanken+Grotesk:wght@400;500;600&family=JetBrains+Mono:wght@400;500&display=swap`

### Space, radius, motion

```css
:root{
  --s1:4px; --s2:8px; --s3:12px; --s4:16px; --s5:24px; --s6:32px; --s7:48px; --s8:64px;
  --r-sm:6px; --r:10px; --r-lg:14px;
  --ease: cubic-bezier(.16,1,.3,1);
  --t-fast: 180ms; --t: 600ms; --t-slow: 1100ms;
}
```

---

## 3. Layout

Desktop (≥1200px). Collapsed icon rail instead of the wide sidebar, so the data gets the width.

```
┌──┬──────────────────────────────────────────────────────────────────────────┐
│◆ │ TOLL OPERATIONS DESK            Thu 15 Oct · 12:40     ● Connected  ⌘K   │  top bar 64px
│  ├──────────────────────────────────────────────────────────────────────────┤
│▣ │ FMC PENDING      │ BG PENDING       │ NOTICES OPEN     │ ONGOING CONTRACTS │  KPI band
│◷ │ 32               │ 45               │ 5 / 14           │ 742               │  (primary, 4)
│⚑ │ ▁▂▅▆▇  +1 wk     │ ▁▃▆▇▇  ±0 wk     │ ◔ 64% rectified  │ ▬▬▬▬▬▬▬▭▭ split   │
│▤ │ PIU|AG|RO|HQ bar │ PIU|AG|RO|HQ bar │ 12 d avg wait    │ 180 agencies      │
│⌕ ├──────────────────────────────────────────────────────────────────────────┤
│  │ avg FMC wait 21d · avg BG wait 18d · Toll notif. 37 under process ·      │  secondary strip
│  │ Windfall 112 plazas ≥40% · data as on 14 Oct (1 day old)                 │  (one line, mono)
│  ├────────────────────────────────────────────┬─────────────────────────────┤
│  │ Register trend (FMC / BG / Notices)        │ Run                         │
│  │  hairline chart, hover crosshair           │  ▸ Weekly KPI report   ~1m  │
│  │                                            │  ▸ Toll notification  ~30s  │
│  │                                            │  ▸ Windfall            ~1m  │
│  │                                            │  ▸ BG / FMC exports         │
│  ├────────────────────────────────────────────┼─────────────────────────────┤
│  │ Latest files                               │ Activity (ticker)           │
└──┴────────────────────────────────────────────┴─────────────────────────────┘
```

- Grid: 12 columns, 24px gutters, max content width 1440px, 32px page padding (16px on mobile).
- KPI band: 4 equal columns separated by **vertical hairlines**, not 4 boxed cards. It reads as one instrument.
- Under 1100px: KPI band 2×2. Under 640px: 1 column, KPI value drops to 48px.
- Rail icons: Overview, Weekly KPI, BG & FMC, CO Division, Library, Explorer. Tooltip on hover, brass 2px bar on the
  active item.

---

## 4. KPI cards: the centrepiece

### 4.1 Which KPIs and where the numbers come from

| Card | Value | Micro-visual | Footnote | Source field |
|---|---|---|---|---|
| **FMC pending** | `fmc.total` | stacked hairline bar PIU / Agency / RO / HQ | `avg wait {pendency} d` + week delta | `/api/report/build` → `fmc.*` |
| **BG pending** | `bg.total` | same stacked bar | `avg wait {pendency} d` + week delta | `bg.*` |
| **Notices open** | `notice.open` / `notice.total` | thin ring = rectified share | `{pendency} d avg wait` | `notice.*` |
| **Ongoing contracts** | `ongoing.total` | split bar Regular / EQ / Manpower | `{agencies} agencies` | `ongoing.*` |
| Secondary strip | FMC & BG avg wait, Toll notification under process, Windfall plazas ≥40%, data age | none | none | build + `/api/co/*` |

Week delta and sparklines come from `report_history.json` (`/api/history`).

### 4.2 Card anatomy

```
 FMC PENDING                       ⓘ        ← mono label, ink-3; info tooltip explains the metric
 32                                         ← Instrument Serif 60px, odometer roll
 ▲ 1 since 12 Oct          ▁▂▅▆▇            ← delta (ink-2, arrow, no red/green); 5-point sparkline in series colour
 ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━            ← stacked bar, 3px, 2px gaps: PIU | Agency | RO | HQ
 PIU 12 · Agency 8 · RO 7 · HQ 5            ← mono 11px legend (identity not by colour alone)
 avg wait 21 d · as on 14 Oct               ← ink-3
```

### 4.3 States

| State | Rule | Look |
|---|---|---|
| Loading | fetch in flight | value replaced by a brass shimmer bar 40% wide, 1.6s sweep |
| Empty | no snapshot yet | value `—`, footnote "Run the weekly report to start tracking" with a link-button |
| Fresh | data ≤ 7 days old | normal |
| Stale | data > 7 days old | footnote turns `--warn` with the word "stale", plus a small "Refresh" link |
| Updated | new value arrives | odometer roll + 1.2s brass hairline sweeping under the value once |

### 4.4 Interaction
- Hover: border goes from `--hair` to `--hair-2`, card lifts **1px**, sparkline end dot grows from 3 to 4px. No glow.
- Click: opens a **right-hand drawer** (480px, `--surface-2`) with the full breakdown. RO-wise table, stage bars
  and Top 10 agencies for contracts. Esc or a click outside closes it.
- Keyboard: cards are `<button>`s. Enter opens the drawer, and the focus ring is 1px brass with a 3px offset.

---

## 5. Motion language

| Moment | Animation | Duration / ease |
|---|---|---|
| First load | Top bar fades in → KPI cells rise 8px + un-blur (4px → 0), staggered 80ms → rest follows | 900ms `--ease` |
| Number appears / changes | **Odometer**: each digit is a vertical strip 0–9 that rolls to its value; digits settle right-to-left with 40ms stagger | 1100ms `--ease` |
| Sparklines / chart lines | stroke draws from left (`stroke-dashoffset`) once per page view | 1200ms |
| Stacked bars | segments grow from left with 60ms stagger | 800ms |
| Ambient | one diagonal **light sweep** (2% white) across the KPI band every 14s; 1.5% film grain over the page | 3s sweep |
| Status dot | "breathing" opacity .5 ↔ 1 | 4s loop |
| Job running | 1px brass line crawling under the row that started it + mono elapsed timer `00:43` | continuous |
| Job done | row background flashes `--brass-dim` once, fades | 1.4s |
| Page switch | outgoing fades 0 in 150ms, incoming rises 6px + fades in | 450ms |
| Drawer / modal | slides 24px + fades; backdrop dims to 60% | 500ms |
| Toast | fades up 8px from bottom-right, stays 5s, fades out. One line, no icon colour shout | 400ms |

Implementation notes:
- CSS only where possible: `@keyframes`, `transition`, `animation-delay` via `--i` custom property on each cell.
- Odometer: a small function `roll(el, value)` that builds `<span class="digit"><span>0…9</span></span>` per
  character and sets `transform: translateY(-N0%)`. Commas and spaces stay static.
- Wrap all of it in `@media (prefers-reduced-motion: no-preference)`.

---

## 6. Backend changes needed (small)

The KPI cards need the **last full report summary**, not only the totals in `report_history.json`.

1. **Persist the last build summary.** In `app.py` → `report_build()`, after a successful build, write the JSON
   response body (minus `file`) to `BASE / "kpi_latest.json"` with a `"date"` field.
2. **Persist the last CO results.** In `co_notification()` and `co_remittance()`, write their response bodies to
   `BASE / "co_latest.json"` under keys `notification` and `windfall`, each with a `"date"`.
3. **New endpoint** `GET /api/dashboard`. It returns `{kpi: <kpi_latest or null>, co: <co_latest or null>,
   history: [...]}` so the page loads with one request.
4. Add `kpi_latest.json` and `co_latest.json` to `.gitignore`. Do **not** add them to `build_exe.py` `DATA`, because
   they are per-computer.
5. Optional: extend the history snapshot in `report_builder.build()` with `fmc_wait`, `bg_wait`, `notice_open` so
   the secondary strip gets trends too. The existing key names stay unchanged so older reports still read.

Acceptance: after one KPI build, restart the server, reload. All four KPI cards show the build's numbers without
running anything.

---

## 7. Step-by-step build order

Finish each step, screenshot it with the mock server (step 7.0), and only then go on.

### 7.0 Set up a test bench
- [ ] Copy `scratchpad/mock_shot.py` into `tools/mock_server.py`: serves `static/` + fake `/api/*` data built from
      the real files in `downloads/`, never contacts the portal.
- [ ] Add `--shots` flag: saves screenshots at 1440×900, 1366×768, 390×844 to `tools/shots/`.
- [ ] Add fake states: `idle`, `captcha_required`, `logged_in`, `job_running`, `stale_data`.

### 7.1 Backend (Section 6)
- [ ] `kpi_latest.json`, `co_latest.json`, `GET /api/dashboard`.
- [ ] Test: build once, restart, `curl http://127.0.0.1:5055/api/dashboard` returns the numbers.

### 7.2 Tokens + shell
- [ ] Replace the `<style>` tokens with Section 2. Remove blobs, glass, gradients, Geist.
- [ ] Icon rail + top bar (title, date, status dot, ⌘K).
- [ ] Check: screenshot is calm and warm-black, and nothing glows.

### 7.3 KPI band
- [ ] Markup: one `<section class="kpis">` with 4 `<button class="kpi">` cells separated by hairlines.
- [ ] `renderKpis(data)` fills value, delta, sparkline, stacked bar, footnote.
- [ ] Odometer `roll()` and loading / empty / stale states.
- [ ] Secondary strip (one mono line, `·` separators).
- [ ] Check at 1366×768: whole band visible without scrolling. Check empty state with no `kpi_latest.json`.

### 7.4 KPI drawer
- [ ] Right drawer with: stage bars + avg waits (FMC/BG), rectified vs open (Notices), Top 10 agencies + split
      (Contracts).
- [ ] Esc / outside click closes, and focus returns to the card.

### 7.5 Trend chart
- [ ] Hairline grid (2 lines max), mono axis labels, 2px series lines, 8px end markers with surface ring, direct
      end labels, crosshair + tooltip, legend toggles, **Table** view.
- [ ] One y-axis only. Ongoing contracts are not on this chart (different scale); they keep their sparkline.

### 7.6 Run panel (actions)
- [ ] One list of jobs: Weekly KPI report, Toll notification, Windfall, BG/FMC exports.
- [ ] Each row: name, one-line description, typical time, **Run** button. While running: crawling brass hairline,
      elapsed timer, live step text from the activity log.
- [ ] Weekly KPI keeps the 6-step pipeline (FMC → BG → Notices → Contracts → Build → Saved) as a thin segmented line.
- [ ] Locked rows (need sign-in) show "Sign in" instead of Run, which opens the sign-in modal.

### 7.7 Sign-in modal
- [ ] Same flow as now (credentials → working → captcha → done), restyled: `--surface-2`, brass focus, captcha in a
      framed inset, mono code input with wide tracking.
- [ ] Error text in `--bad` with a word, never colour only.

### 7.8 Pages: BG & FMC, CO Division, Library, Explorer
- [ ] Same features as today, restyled with hairline rows instead of boxed cards.
- [ ] Library: search + type filter + grouped by day. File type as a 3-letter mono tag, not coloured tiles.
- [ ] Explorer: sections as mono chips, documents list, live view in a thin frame.

### 7.9 Motion pass (Section 5)
- [ ] Add every animation in the table, one at a time, screenshot/record each.
- [ ] Verify reduced-motion: everything instant, nothing hidden.

### 7.10 QA
- [ ] No console errors (mock server prints `pageerror`).
- [ ] No horizontal scroll at 390px.
- [ ] Contrast: `--ink-3` on `--surface-1` ≥ 4.5:1 for text (adjust if not).
- [ ] Keyboard: Tab through rail → KPIs → Run buttons; ⌘K / Ctrl+K palette; Esc closes overlays.
- [ ] Real run: connect, captcha, weekly report, toll notification, windfall, one BG export, library download.

### 7.11 Ship
- [ ] Keep `static/index.classic.html` as fallback.
- [ ] `py build_exe.py` so the shared folder gets the new page.
- [ ] Update `READ ME FIRST.txt` text in `build_exe.py` if button names changed.

---

## 8. Definition of done

- The first screen shows the four KPIs, their trend and their freshness, with no clicks needed.
- A person seeing it for the first time can say what each number means (labels + ⓘ tooltips).
- Every feature of the current page still works (connect, captcha, BG/FMC exports with agency filter, weekly
  report, toll notification, windfall, library, explorer, live view, activity log, disconnect).
- No item from the banned list is on the page.
- Screenshots at 1440, 1366 and 390 widths look intentional, with nothing clipped and nothing overlapping.

---

## 9. Open questions (decide before 7.3)

1. Should the main FMC/BG number be **total in register** (what `report_history.json` stores today) or **pending**
   (PIU + Agency + RO + HQ)? The plan assumes pending is the headline and total goes in the drawer.
2. "Stale" threshold: 7 days (weekly report) or something else?
3. Keep a light theme for printing / projector use later, or dark only?
