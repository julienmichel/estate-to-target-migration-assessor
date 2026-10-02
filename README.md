# Estate-to-Target Migration Assessor

Turns whatever SQL Server inventory a customer sends into a defensible Azure migration plan: target platform per database, blockers with evidence, vCore sizing, effort bands, a sequenced wave plan, and a BoQ hand-off priced against live Azure retail rates.

Built as a [Microsoft Scout](https://aka.ms/scout) skill, but the CLI runs standalone anywhere Python does.

```bash
python scripts/assess.py --input inventory.xlsx --customer "Contoso" --region uaenorth
```

## The problem

A customer sends a SQL estate inventory. Someone hand-reads it and rebuilds the same recommendation deck: which databases can go to PaaS, what's blocking the rest, how big the targets need to be, what it will cost, and in what order to move. It takes most of a day, and the numbers are rarely defensible because they're assembled by eye.

This does it in seconds, and shows its working.

## Input formats

Mixed formats in a single run. Records describing the same database are merged, not double-counted.

| Format | Strength |
|---|---|
| Azure Migrate assessment export (`.xlsx`) | Best case — carries performance telemetry |
| Data Migration Assistant report (`.json` / `.csv`) | Compatibility blockers |
| `sp_Blitz` output (`.csv`) | Configuration findings |
| Hand-built inventory spreadsheet (`.csv` / `.xlsx`) | Works; sizing confidence flagged low |

Header matching is fuzzy and case-insensitive, because no two exports in the wild name these columns the same way.

## Output

| File | Audience |
|---|---|
| `<customer>_assessment.xlsx` | Working document — 8 sheets, filterable, every finding traceable to its source file |
| `<customer>_assessment.html` | Self-contained report for screen-sharing |
| `<customer>_boq.csv` | Machine-readable hand-off to BoQ / pricing tooling |
| `<customer>_assessment.json` | Structured output for downstream reasoning |

Sheets: Executive Summary · Target Recommendation · Blockers & Findings · **Path to PaaS** · Sizing & BoQ · Wave Plan · Source Inventory · Method & Assumptions.

## How it decides

**Target selection is elimination, not scoring.** Each detected finding removes the platforms it's incompatible with. The cheapest-to-operate survivor wins, in the order Azure SQL Database → SQL Managed Instance → SQL Server on Azure VM. SQL on VM is the universal fallback and only appears once every PaaS option is gone.

**Microsoft Fabric is never auto-assigned.** It's flagged as a re-platform *candidate* when naming and size suggest an analytical workload, because moving an OLTP database to a warehouse changes the application contract and needs a human decision.

**Residual vs. resolved findings** — the distinction that makes the effort numbers credible. SQL Agent jobs "block" Azure SQL Database, but SQL Managed Instance runs Agent natively. Once MI is the recommendation there's nothing to re-architect, only to migrate. Those are reported separately, with a 25% carry-over for the move-and-test work that remains. Counting every blocker as remediation is the standard way these assessments overstate effort by ~3x.

**Capacity is checked after growth allowance.** A 14.5 TB database sits under the 16 TB Managed Instance ceiling — until you add log and growth, at which point it doesn't. The engine catches that, re-targets, and explains why rather than silently clamping.

**Sizing** derives vCores from source cores × CPU utilisation × peak ratio × headroom, rounded up to a real SKU on the correct ladder for the target platform. Missing telemetry downgrades the confidence flag and says so on every surface.

**Pricing** is live Azure retail list price from the [public pricing API](https://learn.microsoft.com/rest/api/cost-management/retail-prices/azure-retail-prices) for the requested region, cached 24h. No authentication or subscription required. Azure Hybrid Benefit figures are indicative. SQL on VM is deliberately unpriced, because the answer depends on VM family, OS licensing and disk configuration chosen at design time.

**Non-production** workloads (Developer/Express edition, or dev/test/UAT naming) are detected, held at General Purpose, capped at Moderate effort and scheduled into Wave 1 as the proving ground.

## Path to PaaS

For every IaaS recommendation, the report states exactly what stands in the way and what removing it costs:

| Database | Blocker | To unblock |
|---|---|---|
| ERP_Core | CLR assemblies marked UNSAFE | 6 days |
| LegacyFinance | FileStream / FileTable | 12 days |
| EnterpriseDW | Exceeds MI storage ceiling after growth | 15 days |

That's a roadmap conversation, not a target column.

## Try it

```bash
pip install pandas openpyxl

python scripts/make_samples.py          # generates fixtures in samples/
python scripts/assess.py \
  --input samples/sample_inventory.csv \
          samples/sample_azure_migrate.xlsx \
          samples/sample_dma_report.json \
          samples/sample_spblitz.csv \
  --customer "Contoso Group" --region uaenorth
```

Sample data is entirely fictional and deliberately messy — inconsistent header spellings, mixed MB/GB units, missing telemetry — because that's what real exports look like.

## Options

| Flag | Default | Notes |
|---|---|---|
| `--input` / `-i` | required | One or more files, mixed formats fine |
| `--out` / `-o` | `./assessment-output` | Output directory |
| `--customer` / `-c` | `Customer` | Used in titles and filenames |
| `--region` / `-r` | `uaenorth` | Azure region for live pricing |
| `--currency` | `USD` | Passed to the retail price API |
| `--hours` | `730` | Billable hours/month |
| `--no-pricing` | off | Skip the price lookup |
| `--formats` | `xlsx,html,csv,json` | Any subset |
| `--rules` | `reference/rules.json` | Point at a customised rule pack |
| `--quiet` / `-q` | off | Suppress progress output |

## Tuning

Everything adjustable lives in [`reference/rules.json`](reference/rules.json) — blocker catalogue and remediation days, platform ceilings, vCore and storage ladders, Business Critical thresholds, Fabric heuristics, effort bands, wave definitions. Edit the rule pack, not the code.

To add a blocker, append to `blockers[]` with `detect` tokens (matched case-insensitively against all source text for a workload) and `blocks` listing the platform IDs it eliminates. The `remediation` string flows straight into the customer-facing report.

## Performance

200 workloads in 3.7s, including live pricing. Degrades gracefully: unreadable, empty and headerless files produce clear errors with non-zero exit codes; an unreachable pricing API produces an unpriced BoQ rather than a crash.

## Limitations

State these to the customer:

- Infers only from the artifacts supplied. It does not connect to the estate.
- A static inventory gives compatibility truth but sizing guesswork. A performance-based Azure Migrate assessment over 7–30 days is strongly recommended before any commercial commitment.
- Application compatibility, ISV support statements and licence entitlements are out of scope.
- List pricing excludes MACC, EA and CSP discounting. Suitable for shaping and comparison; it is not a quote.
- Advisory output. A data platform specialist should review before it reaches a customer.

## Layout

```
SKILL.md                 Microsoft Scout skill manifest
reference/rules.json     tunable rule pack
samples/                 demo artifacts in all four input formats
scripts/
  assess.py              CLI orchestrator
  ingest.py              format detection, fuzzy header mapping, merging
  engine.py              blockers, target selection, sizing, waves
  pricing.py             live Azure retail prices with offline fallback
  report.py              xlsx / html / csv / json output
  make_samples.py        regenerates the sample artifacts
```

## Install as a Scout skill

```bash
git clone https://github.com/julienmichel/estate-to-target-migration-assessor.git \
  ~/.scout/m-skills/estate-to-target
```

Then invoke with `/estate-to-target`.

## Licence

MIT — see [LICENSE](LICENSE).

Not an official Microsoft product. Recommendations are advisory and should be validated against current Azure documentation and the customer's agreement before use.
