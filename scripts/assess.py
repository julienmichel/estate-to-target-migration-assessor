#!/usr/bin/env python3
"""
Estate-to-Target Migration Assessor - command line entry point.

Turns whatever SQL Server inventory a customer sends into a defensible migration
recommendation: target platform per workload, blockers with evidence, Azure sizing,
effort bands, a sequenced wave plan and a BoQ hand-off.

    python assess.py --input inventory.xlsx --customer "Contoso" --region uaenorth

    python assess.py --input azure_migrate.xlsx dma_report.json blitz.csv \\
                     --customer "Contoso" --region uaenorth --out ./out

    python assess.py --input inventory.csv --no-pricing --quiet
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

import engine  # noqa: E402
import ingest  # noqa: E402
import pricing  # noqa: E402
import report  # noqa: E402


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="assess.py",
        description="Assess a SQL Server estate and produce an Azure migration plan.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Accepts Azure Migrate exports (.xlsx), DMA reports (.json/.csv), "
               "sp_Blitz output (.csv) and plain inventory spreadsheets.",
    )
    p.add_argument("--input", "-i", nargs="+", required=True,
                   help="One or more source artifacts. Mixed formats are fine and are merged.")
    p.add_argument("--out", "-o", default="./assessment-output",
                   help="Output directory (default: ./assessment-output)")
    p.add_argument("--customer", "-c", default="Customer",
                   help="Customer name used in titles and filenames")
    p.add_argument("--region", "-r", default="uaenorth",
                   help="Azure region for pricing, e.g. uaenorth, westeurope (default: uaenorth)")
    p.add_argument("--currency", default="USD", help="Pricing currency (default: USD)")
    p.add_argument("--hours", type=int, default=730,
                   help="Billable hours per month (default: 730 - continuous running)")
    p.add_argument("--rules", default=None,
                   help="Override the rule pack path (default: reference/rules.json)")
    p.add_argument("--no-pricing", action="store_true",
                   help="Skip the Azure Retail Prices lookup and emit an unpriced BoQ")
    p.add_argument("--formats", default="xlsx,html,csv,json",
                   help="Comma-separated outputs (default: xlsx,html,csv,json)")
    p.add_argument("--quiet", "-q", action="store_true", help="Suppress progress output")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    def say(msg: str = "") -> None:
        if not args.quiet:
            print(msg)

    paths = [Path(p).expanduser().resolve() for p in args.input]
    rules = engine.load_rules(Path(args.rules).resolve() if args.rules else None)
    formats = {f.strip().lower() for f in args.formats.split(",") if f.strip()}

    say("=" * 74)
    say("  ESTATE-TO-TARGET MIGRATION ASSESSOR")
    say("=" * 74)

    # 1. Ingest -----------------------------------------------------------
    say(f"\n[1/5] Ingesting {len(paths)} artifact(s)...")
    workloads, warnings = ingest.load(paths)
    for w in warnings:
        say(f"      ! {w}")
    if not workloads:
        print("\nERROR: No workloads could be read from the supplied files.\n"
              "       Expected an instance or database column. Check --input, or run with a\n"
              "       plain CSV containing at minimum: Instance, Database, Size (GB).",
              file=sys.stderr)
        return 2

    by_format: dict[str, int] = {}
    for w in workloads:
        by_format[w.source_format] = by_format.get(w.source_format, 0) + 1
    say(f"      {len(workloads)} workload(s) across "
        + ", ".join(f"{v} from {k}" for k, v in by_format.items()))

    # 2. Assess -----------------------------------------------------------
    say("\n[2/5] Detecting blockers and selecting targets...")
    assessments = engine.assess(workloads, rules)
    hard = sum(len(a.blockers) for a in assessments)
    soft = sum(len(a.warnings) for a in assessments)
    say(f"      {hard} hard blocker(s), {soft} advisory finding(s)")

    # 3. Price ------------------------------------------------------------
    if args.no_pricing:
        say("\n[3/5] Pricing skipped (--no-pricing)")
        book = None
    else:
        say(f"\n[3/5] Retrieving live Azure retail prices for '{args.region}'...")
        book = pricing.apply_pricing(assessments, args.region, args.currency, args.hours)
        if book.available:
            priced = sum(1 for a in assessments if a.est_monthly_usd)
            say(f"      Priced {priced}/{len(assessments)} workload(s) at list")
        else:
            say(f"      ! {book.error}")
            warnings.append(book.error)

    # 4. Summarise --------------------------------------------------------
    say("\n[4/5] Building wave plan and rolling up...")
    summary = engine.summarise(assessments)

    meta = {
        "customer": args.customer,
        "region": args.region,
        "currency": args.currency,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "sources": ", ".join(p.name for p in paths),
        "rule_version": rules["_meta"]["version"],
    }

    # 5. Write ------------------------------------------------------------
    outdir = Path(args.out).expanduser().resolve()
    outdir.mkdir(parents=True, exist_ok=True)
    stem = "".join(c if c.isalnum() or c in "-_" else "_" for c in args.customer)

    say(f"\n[5/5] Writing output to {outdir}")
    written: list[Path] = []

    if "xlsx" in formats:
        p = outdir / f"{stem}_assessment.xlsx"
        report.write_xlsx(p, assessments, summary, meta, warnings, rules)
        written.append(p)
    if "csv" in formats:
        p = outdir / f"{stem}_boq.csv"
        report.write_boq(p, assessments, meta)
        written.append(p)
    if "html" in formats:
        p = outdir / f"{stem}_assessment.html"
        report.write_html(p, assessments, summary, meta, warnings, rules)
        written.append(p)
    if "json" in formats:
        p = outdir / f"{stem}_assessment.json"
        report.write_json(p, assessments, summary, meta, warnings)
        written.append(p)

    for p in written:
        say(f"      + {p.name}  ({p.stat().st_size / 1024:.1f} KB)")

    # Console summary -----------------------------------------------------
    if not args.quiet:
        s = summary
        pct = (s["paas_eligible"] / s["total_workloads"] * 100) if s["total_workloads"] else 0
        print("\n" + "=" * 74)
        print("  ASSESSMENT SUMMARY")
        print("=" * 74)
        print(f"  Workloads              {s['total_workloads']:,}")
        print(f"  Data volume            {s['total_size_gb']:,.0f} GB")
        print(f"  PaaS eligible          {s['paas_eligible']:,} ({pct:.0f}%)")
        print(f"  IaaS required          {s['iaas_required']:,}")
        print(f"  Fabric candidates      {s['fabric_candidates']:,}")
        print(f"  Non-production         {s.get('non_production', 0):,}  (scheduled first as proving ground)")
        print(f"  Out of support         {s['eos_workloads']:,}")
        print(f"  With hard blockers     {s['workloads_with_hard_blockers']:,}"
              f"  -> {s.get('workloads_with_residual_work', 0):,} still need work after targeting")
        print(f"  Findings resolved      {s.get('findings_resolved_by_target', 0):,}  by the target choice alone")
        print(f"  Effort                 {s['effort_days_low']:,}-{s['effort_days_high']:,} person-days"
              f"  (plus {s['remediation_days']:,} remediation)")
        if s["est_monthly_usd"]:
            print(f"  Est. run-rate          {s['est_monthly_usd']:,.0f} {args.currency}/mo list"
                  f"  |  {s['est_monthly_ahb_usd']:,.0f} with AHB")

        print("\n  Target split")
        for t, c in sorted(s["by_target"].items(), key=lambda kv: -kv[1]):
            print(f"    {c:>4}  {t}")

        print("\n  Waves")
        for wid, c in s["by_wave"].items():
            name = next((w["name"] for w in rules["waves"] if w["id"] == wid), f"Wave {wid}")
            print(f"    {c:>4}  {name}")

        if s["top_blockers"]:
            print("\n  Top findings")
            for label, count in s["top_blockers"][:6]:
                print(f"    {count:>4}  {label[:62]}")

        if s["low_confidence"]:
            print(f"\n  CAVEAT: {s['low_confidence']} workload(s) sized on assumed values "
                  f"(missing CPU/size telemetry).")
            print("          Recommend a performance-based Azure Migrate assessment before quoting.")
        print("=" * 74 + "\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
