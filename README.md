# mim-program-research

A Claude Code skill that turns "rank the best master's programs for me" into a repeatable pipeline: parallel regional research agents → verification against official pages → weighted tier ranking → markdown report + Excel workbook.

Built from the 2027 MiM/MSc-Finance research cycle for Vedant Konde (GMAT Focus 655, India/US/Australia excluded). The full 25+15 program dataset lives in this repo's root: `MIM_TOP25_REPORT.md` (narrative + sources) and `MIM_TOP25_AND_MORE.xlsx` (the deliverable).

## Use it

In Claude Code, invoke with the skill name or the slash command:

    /mim-program-research GMAT Focus 655, 2027 intake, exclude India/US/Australia

or let Claude pick it up when you ask to "research/rank master's programs for [profile]".

## What it does

1. Collects candidate profile + constraints (reads any existing strategy file first).
2. Spawns 6 regional research agents in parallel (APAC / UK-IE / Nordics / DACH / FR-CH-NL / Southern-EU-ME), each returning the 13-field program template.
3. Handles the failure modes that actually happened: agents finishing without handing back (demand report via message, relaunch after 2 failures), search throttling (tag everything `[unverified]`).
4. Verifies load-bearing claims against official sources (gov immigration pages, uni fee/admissions pages) and tags every number CONFIRMED / EST / VERIFY.
5. Ranks with fixed weights (admission odds ~25%, placement ease ~20%, ROI ~20%, visa safety ~15%, work-while-studying ~10%, living + Indian verdict ~5%, prestige ~4% tiebreaker, scholarships ~1%).
6. Outputs the master table + application-wave calendar + scholarship leverage + pre-fee re-verify checklist, and builds the 4-sheet Excel workbook.

## Files

- `SKILL.md` — the playbook (13-field template, weights, verification protocol, known traps)
- `scripts/build_workbook_template.py` — openpyxl builder, a WORKING EXAMPLE with the Vedant dataset embedded; regenerate the data rows per candidate. Needs `pip install openpyxl`.

## Known traps (from the 2027 cycle)

See SKILL.md §8 — includes the CBS-Amsterdam mixup, HEC being a 2-yr program, Maastricht's rename, Bocconi's real 2-yr cost, NYUAD's unconfirmed finance MSc, and CH/UAE zero-work-hours rules.
