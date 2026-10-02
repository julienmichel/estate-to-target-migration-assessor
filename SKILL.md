---
name: estate-to-target
description: Assess a SQL Server estate and produce an Azure migration plan. Use when a customer sends a SQL inventory, Azure Migrate export, Data Migration Assistant (DMA) report, or sp_Blitz output and you need target platform recommendations (Azure SQL Database / Managed Instance / SQL on VM / Fabric), migration blockers, vCore sizing, effort bands, a wave plan, or a BoQ hand-off. Triggers include "assess this SQL estate", "which Azure SQL target", "SQL migration assessment", "size these databases for Azure", "migration wave plan", "DMA report", "Azure Migrate export", "can this go to Managed Instance".
---

# Estate-to-Target Migration Assessor

Turns whatever SQL Server inventory a customer sends into a defensible Azure
migration recommendation. Every verdict carries its evidence and its reasoning,
because an assessment the customer can argue with is worthless.

## When to use this

Use it when you have a **real artifact** describing a SQL estate. Do not use it
to speculate about a migration with no data — the whole point is grounding.

| You have | It works |
|---|---|
| Azure Migrate assessment export (`.xlsx`) | Best case — carries performance telemetry |
| Data Migration Assistant report (`.json` / `.csv`) | Strong on compatibility blockers, no sizing data |
| `sp_Blitz` output (`.csv`) | Strong on configuration findings |
| A spreadsheet of instances someone typed by hand | Works; sizing confidence will be flagged low |
| Several of the above together | **Preferred** — they merge into one richer record per database |

## Quick start

```bash
cd ~/.scout/m-skills/estate-to-target/scripts

# Single artifact
python assess.py --input inventory.xlsx --customer "Contoso" --region uaenorth

# Several artifacts - they are merged per database, not double-counted
python assess.py \
  --input azure_migrate.xlsx dma_report.json spblitz.csv \
  --customer "Contoso" --region uaenorth --out ./out

# Offline / no pricing lookup
python assess.py --input inventory.csv --no-pricing
```

Run `python make_samples.py` once to generate demo artifacts in `../samples/`
covering all four input formats. Useful for showing the skill without customer data.

## Options

| Flag | Default | Notes |
|---|---|---|
| `--input` / `-i` | required | One or more files, mixed formats fine |
| `--out` / `-o` | `./assessment-output` | Output directory |
| `--customer` / `-c` | `Customer` | Used in titles and filenames |
| `--region` / `-r` | `uaenorth` | Azure region for live pricing |
| `--currency` | `USD` | Passed to the retail price API |
| `--hours` | `730` | Billable hours/month; lower it for non-production |
| `--no-pricing` | off | Skip the price lookup entirely |
| `--formats` | `xlsx,html,csv,json` | Any subset |
| `--rules` | `reference/rules.json` | Point at a customised rule pack |
| `--quiet` / `-q` | off | Suppress progress output |

## What it produces

| File | Who it is for |
|---|---|
| `<customer>_assessment.xlsx` | The working document. Eight sheets, filterable, every finding traceable to its source file. |
| `<customer>_assessment.html` | The thing you screen-share. Self-contained, no dependencies. |
| `<customer>_boq.csv` | Machine-readable hand-off to BoQ / pricing tooling. |
| `<customer>_assessment.json` | Structured output so you can reason over the result conversationally. |

Workbook sheets: Executive Summary, Target Recommendation, Blockers & Findings,
**Path to PaaS**, Sizing & BoQ, Wave Plan, Source Inventory, Method & Assumptions.

## How it decides

**Target selection** is elimination, not scoring. Each detected finding removes the
platforms it is incompatible with; the cheapest-to-operate survivor wins, in the
order Azure SQL Database → SQL Managed Instance → SQL Server on Azure VM. SQL on VM
is the universal fallback and only appears when every PaaS option is gone.

**Microsoft Fabric is never auto-assigned.** It is flagged as a re-platform
*candidate* when naming and size suggest an analytical workload, because moving an
OLTP database to a warehouse changes the application contract and needs a human
decision.

**Residual versus resolved findings** is the distinction that makes the effort
numbers credible. SQL Agent jobs "block" Azure SQL Database — but SQL Managed
Instance runs Agent natively, so once MI is recommended there is nothing to
re-architect, only to migrate. The tool reports those separately. Counting all
blockers as remediation is the standard way these assessments overstate effort by
a factor of three.

**Sizing** derives vCores from source cores × CPU utilisation × peak ratio ×
headroom, rounded up to a real SKU on the right ladder for the target platform.
Provisioned storage is checked against the *platform ceiling* after growth
allowance — a 14 TB database fits under the 16 TB MI limit until you add log and
growth, at which point it does not, and the tool re-targets it.

**Pricing** is live Azure retail list price from the public pricing API for the
requested region, cached 24h. Azure Hybrid Benefit figures are indicative and must
be validated against the customer's Software Assurance position. SQL on VM is
deliberately not priced, because the answer depends on VM family, OS licensing and
disk configuration chosen at design time.

**Non-production** workloads (Developer/Express edition, or dev/test/UAT naming)
are detected, held at General Purpose, capped at Moderate effort and scheduled into
Wave 1 as the proving ground.

## Tuning

Everything adjustable lives in `reference/rules.json` — blocker catalogue and
remediation days, platform ceilings, vCore and storage ladders, Business Critical
thresholds, Fabric heuristics, effort bands, wave definitions. Edit that rather
than the code. Each rule carries a `remediation` string that flows straight into
the report, so improving the rule pack improves the customer-facing output.

To add a blocker, append to `blockers[]` with `detect` tokens matched
case-insensitively against all source text for a workload, and `blocks` listing the
platform IDs it eliminates.

## Presenting the result

1. **Lead with the Path to PaaS sheet.** "44 of your 200 databases need IaaS, and
   here is exactly what it costs to change that" is a far better conversation than
   a target column.
2. **Check the confidence flags before quoting.** Anything marked low confidence was
   sized on assumed values. Say so.
3. **Push for performance data.** A static inventory gives compatibility truth but
   sizing guesswork. A performance-based Azure Migrate assessment over 7–30 days
   turns the sizing from indicative into defensible.
4. **Treat effort bands as planning estimates, not a statement of work.**

## Limitations — state these to the customer

- Infers only from the artifacts supplied; it does not connect to the estate.
- Application compatibility, ISV support statements and licence entitlements are
  out of scope.
- List pricing excludes MACC, EA and CSP discounting.
- Advisory output. A data platform specialist should review before it reaches a
  customer.

## Layout

```
estate-to-target/
  SKILL.md
  reference/rules.json        tunable rule pack - edit this, not the code
  samples/                    demo artifacts in all four input formats
  scripts/
    assess.py                 CLI orchestrator
    ingest.py                 format detection, fuzzy header mapping, merging
    engine.py                 blockers, target selection, sizing, waves
    pricing.py                live Azure retail prices with offline fallback
    report.py                 xlsx / html / csv / json output
    make_samples.py           regenerates the sample artifacts
```

Requires `pandas` and `openpyxl`. No authentication and no Azure subscription
needed — the pricing API is public.
