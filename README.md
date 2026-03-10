# xlsx-ultimate

Autonomous Excel agent that builds, verifies, and delivers spreadsheets with zero errors. Every formula is independently backtested in Python.

## Install

```bash
npx skills add SLFMR1/xlsx-ultimate
```

Or manually copy the folder to `~/.claude/skills/xlsx-ultimate/`.

## What It Does

Tell it what you need in plain language. It handles the rest:

- Asks the right questions upfront
- Plans the sheet architecture
- Builds with real Excel formulas (never hardcoded values)
- Verifies every calculation independently in Python
- Delivers with a verification report

Works for financial models, engineering calculations, dashboards, budgets, payroll, and anything else that belongs in a spreadsheet.

## Domains

**Finance** — SaaS metrics, P&L, DCF, cap tables, multi-currency, Monte Carlo, bookkeeping

**Engineering** — Bolt sizing, HVAC, electrical panels, pipe flow, structural, solar ROI

**Business** — KPI dashboards, sales pipelines, project budgets, inventory, marketing funnels

## How Verification Works

```
Build xlsx → Read inputs → Recalculate in Python → Compare → Report
```

Every key value is computed twice: once as an Excel formula, once in Python. If they don't match, the agent fixes the formula and re-verifies. Nothing ships with a failed check.

## Test Results

27 tests across 15 domains. Final iteration: **150/150** (15/15 tests at 10/10).

Validated against real-world spreadsheets from NYU Stern (Damodaran FCFF), SCORE/SBA, and EPA.

## Structure

```
xlsx-ultimate/
├── SKILL.md              # Agent instructions
├── evals/evals.json      # 18 test cases
├── references/            # Domain knowledge (finance, engineering, formulas)
├── scripts/               # Verification & recalc scripts
├── README.md
└── LICENSE
```

## License

MIT
