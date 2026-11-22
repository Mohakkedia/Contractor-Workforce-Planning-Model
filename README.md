# 📈 Contractor Workforce Planning Model

An 8-quarter contractor headcount and spend forecast against budget, a
contractor-vs-FTE cost comparison for 5 roles, and a **live scenario
toggle** (dropdown: Base Case / 10% Rate Increase / Hiring Freeze) that
recalculates every downstream number — headcount, cost, cumulative
budget, and exactly which quarter the budget runs out — with zero
copy-pasting between scenarios.

## A bug this model's own verification caught

The first version rounded headcount to a whole number every quarter
(`ROUND(prior_headcount x 1.025, 0)`). At a 2.5%/quarter growth rate,
that's fatal for small headcounts: 12 x 1.025 = 12.3, which rounds
**right back down to 12** — the growth signal silently disappeared and
every quarter showed identical, flat spend. Caught by checking the
actual computed output rather than trusting that "zero formula errors"
meant "correct model" (a recalc engine confirms formulas evaluate; it
says nothing about whether the *model* is doing what it's supposed to).
Fixed by keeping headcount as a continuous "FTE-equivalent contractor
count" (standard financial-modeling practice for blended mid-quarter
joins) and only formatting it to one decimal for display, rather than
rounding the actual value every quarter.

## Real results, all three scenarios (verified by editing the dropdown and recalculating)

| Scenario | Headcount trajectory | Budget outcome |
|---|---|---|
| **Base Case** | 34.0 to 40.4 (2.5%/qtr growth) | Cumulative surplus peaks ~Q4 (~$294K cushion), then compounding growth erodes it — **budget exceeded starting Q8** |
| **10% Rate Increase** | Same growth as Base Case | **Over budget from Q1** — the one-time rate jump alone exceeds quarterly budget before growth even factors in |
| **Hiring Freeze** | Flat at 34.0 for all 8 quarters | **Never exceeds budget** — surplus grows every quarter, reaching $1.02M cushion by Q8 |

This is a genuinely useful "which lever matters more" comparison: a 10%
rate increase is *immediately* worse than unmanaged 2.5%/quarter organic
growth compounding for two years — a finding a procurement/finance
partner would actually want to know before a rate renegotiation.

## Contractor vs. FTE cost comparison (same 5 roles)

| Role | Contractor Annual | FTE Annual (loaded) | Premium |
|---|---|---|---|
| Software Engineer | $197,600 | $191,400 | +3.2% |
| Data Analyst | $156,000 | $145,200 | +7.4% |
| Business Analyst | $145,600 | $138,600 | +5.1% |
| Project Manager | $176,800 | $165,000 | +7.2% |
| QA Engineer | $135,200 | $125,400 | +7.8% |

Contractor cost = hourly rate x 2,080 hours/year. FTE cost = salary x
(1 + 32% benefits/overhead loading). A 3-8% contractor premium across
roles is a realistic range — high enough to matter at scale, not so
high it reads as a rigged comparison either direction.

## Quickstart

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt

python scripts/build_workforce_model.py   # -> outputs/Contractor_Workforce_Planning_Model.xlsx
```

Open the workbook, go to `Assumptions!C4`, and switch the dropdown
between the three scenarios — the `Forecast` sheet's headcount, cost,
cumulative budget, "Within Budget?" row, and chart all recalculate live.

## Key Technologies
Microsoft Excel (scenario toggle via data validation + IF-branching formulas, conditional formatting, native charts, what-if analysis)

## Disclaimer
Educational/portfolio project using illustrative role/rate/salary figures.
