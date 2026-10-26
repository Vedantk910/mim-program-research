---
name: mim-program-research
description: Research and rank Master's programs (MiM / MSc Finance / Quant Finance) for a specific candidate across regions, with ROI, visa, work-while-studying, placement quality, and Indian-community data. Use when asked to shortlist or compare international master's programs for a candidate profile, or to rebuild/update the ranked program table and Excel deliverable.
---

# MiM / MSc Finance Program Research

Rank 25-40 Master's programs for ONE candidate profile. The output is a tiered, weighted ranking with per-program data across 13 dimensions, plus an Excel deliverable. This skill encodes the full pipeline that produced `MIM_TOP25_REPORT.md` + `MIM_TOP25_AND_MORE.xlsx` for Vedant Konde (GMAT Focus 655, Sep 2027 intake, India/US/Australia excluded).

## 1. Inputs (ask for these if missing)

- Candidate: nationality, age, degree + CGPA, test score (GMAT Focus or GRE), work experience, extracurricular/founder narrative, target intake month
- Constraints: countries to EXCLUDE, hard requirements (visa length, work-while-studying, budget), ranking priority (prestige vs ROI vs living)
- Target tier sizes: usually Top 25 + 15 near-misses

Read the candidate's existing strategy file first if present (e.g. `STRATEGY_2027_APPLY.md`) — it carries prior decisions that must not be contradicted silently.

## 2. Research: 6 parallel regional agents

Spawn one general-purpose agent per region (all in a single message so they run concurrently):

1. **APAC** — Singapore (NUS/NTU/SMU), HK (HKUST/HKU), Japan/other if candidate wants
2. **UK/Ireland** — Imperial, LSE, Warwick, Oxford SAID, Cambridge, DCU
3. **Nordics** — SSE, KTH, GSE, CBS, Lund, Aalto
4. **DACH** — TUM, WHU, Frankfurt School, UZH, St. Gallen, WU, LMU, EBS
5. **France/CH/NL** — HEC, ESSEC, EDHEC, ETH, Maastricht, Tilburg, Erasmus, NEOMAP
6. **Southern EU/Middle East** — IE, ESADE, Bocconi, Ca' Foscari, CEU, NYU Abu Dhabi

Each agent gets: the full candidate profile, its region's program list, and the **13-field template** below. Tell it explicitly: *"DELIVER YOUR FULL REPORT AS PLAIN TEXT IN YOUR FINAL MESSAGE"* and *"mark unverified figures, never invent numbers"*.

### 13-field template (per program)

1. Program + university + duration + intake deadline
2. Test-score odds (low/med/high + the class average if findable)
3. Tuition + all-in cost (tuition + living, in local currency AND INR equivalent)
4. % of class that is Indian (forum estimates if unpublished — say so)
5. Median starting salary + top 4 employers
6. ROI: years to reclaim ALL cost (show the math)
7. Work rights WHILE studying (hours/week, exact rule)
8. Post-study visa (name + length, and whether a search period exists at all)
9. Political/economic scenario for the target year
10. "Indian verdict" from forums (net positive/negative + 1-2 specific complaints)
11. Best scholarships + realistic size for this profile
12. Spoon-fed-ness rating (1-10: does the program run campus recruiting/onsites/conversions?)
13. Ease of Indian living (food/veg, weather, safety, community size, alcohol)

## 3. Failure handling (this is where most of the cost went last time)

- Agents frequently **finish without handing back** their report. On each completion notification, if no report arrived, immediately `SendMessage` to that agent: *"Deliver your full report now as your final plain-text output — no more tool calls; mark anything unverified."* If it still fails twice, **relaunch** that region with the same brief.
- Agents hit search throttling mid-run. Their numbers then default to 2024-25 knowledge — tag them `[unverified]`.
- **Never** accept an agent figure for a load-bearing claim (fees, visa length, placement %) without one of: the agent's own source URL, a direct WebFetch of the official page, or an independent search. Verified > agent-asserted > prior knowledge.

## 4. Verification protocol (do this after agents return, before compiling)

Verify against official sources (gov/uni pages; not blogs):

- Post-study visa lengths: `irishimmigration.ie` (Stamp 1G), `Migrationsverket.se`, `Berlin Welcome Center` (18-mo), `UvA.nl`/IND (NL 3-yr renewal), `mom.gov.sg` (EP/COMPASS), `immd.gov.hk` (IANG), `Migrationsverket`/canton pages (CH = none)
- Fees + deadlines + placement stats: each university's admissions/fees page
- Work-while-studying rules: national rules change (e.g. Spain went to 30h/wk; France = 964 h/yr)
- Class averages: only if the school publishes them

Then tag every table cell: **CONFIRMED** (official source, this cycle) / **EST** (2024-25 vintage or agent estimate) / **VERIFY** (could not confirm — including program existence).

## 5. Ranking weights (in descending order)

| Weight | Factor |
|---|---|
| ~25% | Admission odds at the candidate's test score (the gate) |
| ~20% | Placement ease: spoon-fed-ness, target firms, conversion rate |
| ~20% | ROI: years to reclaim tuition + living at local post-grad salary |
| ~15% | Post-study visa safety net + political stability |
| ~10% | Reclaim-while-studying: part-time/full-time work rights + realistic earnings |
| ~5% | Indian verdict + living ease |
| ~4% | Prestige / Indian-credential equivalence (tiebreaker only) |
| ~1% | Scholarship access (accelerator, not tier-decider) |

Decision rule: **admission fit + reclaim speed set the tier; prestige orders within a tier.** Exceptions: a low-odds school stays in the list only if its salary/ROI payoff justifies the gamble (tag it "reach"). A no-search-visa country (CH, UAE) costs one tier regardless of brand.

## 6. Compilation

- One master table, 25 rows + near-miss block, columns: rank, program, university, country, tier, duration, deadline, odds, tuition, all-in, % Indian, salary, ROI, work-while-studying, post-study visa, spoon-fed /10, Indian living, data confidence, one-line verdict
- End with: anchor set (Tier 1), application-wave calendar sorted by deadline, scholarship-leverage list, reclaim-while-studying ladder, post-study-visa ladder, and a **pre-fee re-verify checklist** (every VERIFY/EST fact with the exact page to check)
- Mark tier changes that came from corrections (e.g. "HEC dropped a tier after fee/salary correction")

## 7. Excel deliverable

Use `scripts/build_workbook_template.py` (openpyxl). Structure:
- Sheet "Top 25 Ranked" — master table, frozen panes + autofilter, tier color-coding
- Sheet "15 Near-Misses (26-40)" — same columns
- Sheet "Strategy & Deadlines" — waves, anchor sets, ladders
- Sheet "Read-Me & Flags" — FX rates used, confidence legend, confirmed-source list, known conflicts, open items
- Run it, verify with `load_workbook` (row counts), then send the .xlsx to the user.

## 8. Known traps from the 2027 cycle (do not repeat)

- "CBS Amsterdam" does not exist — CBS = Copenhagen Business School (Denmark); Amsterdam = UvA
- HEC MiM is a **2-year** Grande École program (€58,970), not a 10-month MiM at ~€27k
- NEOMAP = EPFL Lausanne (Switzerland), not France's NEOMA
- Maastricht's program was renamed to "MSc in Financial Economics"
- Bocconi 2-yr all-in is ~€115-125k, not €30-35k (that was the 1-yr misestimate)
- NYU Abu Dhabi: unconfirmed whether a finance MSc exists at all; campus wind-down ~2028; UAE = zero work rights while studying
- Switzerland (and UAE) allow **zero** work hours during study for non-EU/3rd-country students — this single fact moves Swiss programs a full tier down for a "reclaim money while studying" candidate
- France work rights = 964 h/yr (≈48 h/wk), not 20 h/wk; Spain = 30 h/wk; Germany = 140 full days/yr; Ireland = full-time (EU)
- Never let an agent's "unverified" pass into the final table as if verified; the confidence column exists for a reason
