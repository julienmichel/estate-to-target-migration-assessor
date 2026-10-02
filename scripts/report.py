"""
Output generation for the Estate-to-Target Migration Assessor.

Three artifacts, each with a different reader in mind:

  * <customer>_assessment.xlsx - the working document. Seven sheets, filterable,
    every finding traceable back to the source file.
  * <customer>_boq.csv         - machine-readable hand-off to BoQ / pricing tooling.
  * <customer>_assessment.html - the thing you actually screen-share, self-contained.
"""

from __future__ import annotations

import csv
import html
import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Palette
NAVY = "1F3864"
BLUE = "2E75B6"
LIGHT = "DDEBF7"
RED = "C00000"
AMBER = "BF8F00"
GREEN = "375623"
GREY = "7F7F7F"

HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(color=NAVY, bold=True, size=16)
SUB_FONT = Font(color=GREY, size=10, italic=True)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TARGET_COLOURS = {
    "Azure SQL Database": "C6E0B4",
    "Azure SQL Managed Instance": "BDD7EE",
    "SQL Server on Azure VM": "FFE699",
    "Microsoft Fabric (Warehouse / Lakehouse)": "E4DFEC",
}


# --------------------------------------------------------------------------
# Shared helpers
# --------------------------------------------------------------------------

def _style_header(ws, row: int, ncols: int) -> None:
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def _autosize(ws, min_w: int = 10, max_w: int = 55) -> None:
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        longest = 0
        for cell in col:
            if cell.value is not None:
                longest = max(longest, max(len(line) for line in str(cell.value).split("\n")))
        ws.column_dimensions[letter].width = max(min_w, min(max_w, longest + 2))


def _add_table(ws, name: str, header_row: int, nrows: int, ncols: int) -> None:
    if nrows <= 0:
        return
    ref = f"A{header_row}:{get_column_letter(ncols)}{header_row + nrows}"
    table = Table(displayName=name, ref=ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showColumnStripes=False)
    ws.add_table(table)


def _fmt_findings(findings: list[dict]) -> str:
    if not findings:
        return ""
    return "\n".join(f"[{f['severity'].upper()}] {f['label']}" for f in findings)


# --------------------------------------------------------------------------
# XLSX
# --------------------------------------------------------------------------

def write_xlsx(path: Path, assessments, summary: dict, meta: dict,
               warnings: list[str], rules: dict) -> None:
    wb = Workbook()

    _sheet_summary(wb.active, summary, meta, warnings)
    _sheet_recommendations(wb.create_sheet("Target Recommendation"), assessments)
    _sheet_blockers(wb.create_sheet("Blockers & Findings"), assessments)
    _sheet_unblock(wb.create_sheet("Path to PaaS"), assessments)
    _sheet_sizing(wb.create_sheet("Sizing & BoQ"), assessments, meta)
    _sheet_waves(wb.create_sheet("Wave Plan"), assessments, rules)
    _sheet_inventory(wb.create_sheet("Source Inventory"), assessments)
    _sheet_assumptions(wb.create_sheet("Method & Assumptions"), rules, meta, warnings)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def _sheet_summary(ws, s: dict, meta: dict, warnings: list[str]) -> None:
    ws.title = "Executive Summary"
    ws["A1"] = "SQL Estate Migration Assessment"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = (f"{meta['customer']}  |  generated {meta['generated']}  |  "
                f"target region {meta['region']}")
    ws["A2"].font = SUB_FONT

    row = 4
    paas_pct = (s["paas_eligible"] / s["total_workloads"] * 100) if s["total_workloads"] else 0

    headline = [
        ("Workloads assessed", f"{s['total_workloads']:,}"),
        ("Total data volume", f"{s['total_size_gb']:,.0f} GB"),
        ("PaaS eligible (SQL DB or MI)", f"{s['paas_eligible']:,}  ({paas_pct:.0f}%)"),
        ("Require IaaS (SQL on VM)", f"{s['iaas_required']:,}"),
        ("Fabric re-platform candidates", f"{s['fabric_candidates']:,}"),
        ("Non-production (migrate first)", f"{s.get('non_production', 0):,}"),
        ("Workloads with hard blockers", f"{s['workloads_with_hard_blockers']:,}"),
        ("  still needing work after targeting", f"{s.get('workloads_with_residual_work', 0):,}"),
        ("  findings resolved by target choice", f"{s.get('findings_resolved_by_target', 0):,}"),
        ("Clean - no findings at all", f"{s['clean_workloads']:,}"),
        ("Out-of-support versions (ESU exposure)", f"{s['eos_workloads']:,}"),
        ("Migration effort estimate", f"{s['effort_days_low']:,} - {s['effort_days_high']:,} person-days"),
        ("  plus remediation", f"{s['remediation_days']:,} person-days"),
    ]
    if s["est_monthly_usd"]:
        headline.append(("Estimated Azure run-rate (list)",
                         f"{s['est_monthly_usd']:,.0f} {meta['currency']}/month"))
        headline.append(("  with Azure Hybrid Benefit",
                         f"{s['est_monthly_ahb_usd']:,.0f} {meta['currency']}/month"))

    ws.cell(row=row, column=1, value="Measure")
    ws.cell(row=row, column=2, value="Value")
    _style_header(ws, row, 2)
    row += 1
    for label, value in headline:
        ws.cell(row=row, column=1, value=label).border = BORDER
        c = ws.cell(row=row, column=2, value=value)
        c.border = BORDER
        c.font = Font(bold=not label.startswith("  "))
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Recommended target platform").font = Font(bold=True, color=NAVY, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Target")
    ws.cell(row=row, column=2, value="Workloads")
    _style_header(ws, row, 2)
    row += 1
    for target, count in sorted(s["by_target"].items(), key=lambda kv: -kv[1]):
        c1 = ws.cell(row=row, column=1, value=target)
        c1.border = BORDER
        if target in TARGET_COLOURS:
            c1.fill = PatternFill("solid", fgColor=TARGET_COLOURS[target])
        ws.cell(row=row, column=2, value=count).border = BORDER
        row += 1

    row += 1
    ws.cell(row=row, column=1, value="Most common findings across the estate").font = Font(bold=True, color=NAVY, size=12)
    row += 1
    ws.cell(row=row, column=1, value="Finding")
    ws.cell(row=row, column=2, value="Workloads affected")
    _style_header(ws, row, 2)
    row += 1
    for label, count in s["top_blockers"]:
        ws.cell(row=row, column=1, value=label).border = BORDER
        ws.cell(row=row, column=2, value=count).border = BORDER
        row += 1

    if s["low_confidence"]:
        row += 1
        c = ws.cell(row=row, column=1,
                    value=f"NOTE: {s['low_confidence']} workload(s) were sized on assumed values because the "
                          f"source data lacked CPU or size telemetry. Treat those sizings as indicative only - "
                          f"a performance-based Azure Migrate assessment is recommended before quoting.")
        c.font = Font(color=AMBER, bold=True)
        c.alignment = Alignment(wrap_text=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.row_dimensions[row].height = 45

    if warnings:
        row += 2
        ws.cell(row=row, column=1, value="Ingestion warnings").font = Font(bold=True, color=RED)
        for w in warnings:
            row += 1
            ws.cell(row=row, column=1, value=w).font = Font(color=RED, size=9)

    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 34
    ws.sheet_view.showGridLines = False


def _sheet_recommendations(ws, assessments) -> None:
    cols = ["Instance", "Database", "Version", "Edition", "Size (GB)", "Non-Prod",
            "Recommended Target", "Service Tier", "vCores", "Storage (GB)",
            "Confidence", "Fabric Candidate", "Residual Work", "Resolved by Target",
            "Effort Band", "Wave", "Rationale"]
    ws.append(cols)
    _style_header(ws, 1, len(cols))

    for a in assessments:
        w = a.workload
        ws.append([
            w.instance, w.database, w.sql_version, w.sql_edition,
            round(w.db_size_gb, 1) if w.db_size_gb else None,
            "YES" if a.is_non_production else "",
            a.target_label, a.service_tier, a.vcores, a.storage_gb,
            a.confidence, "YES" if a.fabric_candidate else "",
            len(a.residual), len(a.resolved_by_target),
            a.effort_band, a.wave,
            " ".join(a.rationale + a.capacity_notes),
        ])
        r = ws.max_row
        fill = TARGET_COLOURS.get(a.target_label)
        if fill:
            ws.cell(row=r, column=7).fill = PatternFill("solid", fgColor=fill)
        conf = ws.cell(row=r, column=11)
        conf.font = Font(color={"high": GREEN, "medium": AMBER, "low": RED}[a.confidence], bold=True)
        if a.resolved_by_target:
            ws.cell(row=r, column=14).font = Font(color=GREEN, bold=True)
        ws.cell(row=r, column=17).alignment = Alignment(wrap_text=True, vertical="top")

    _add_table(ws, "Recommendations", 1, len(assessments), len(cols))
    _autosize(ws)
    ws.column_dimensions["Q"].width = 85
    ws.freeze_panes = "C2"


def _sheet_unblock(ws, assessments) -> None:
    """What it would take to move an IaaS recommendation onto PaaS.

    This is the conversation that turns an assessment into a roadmap, so it gets
    its own sheet rather than being buried in a rationale column.
    """
    ws["A1"] = "Path to PaaS - what stands between these workloads and a managed platform"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = ("Each of these landed on SQL Server on Azure VM. Remove the blockers listed and the "
                "workload becomes eligible for Azure SQL Managed Instance, with the patching, backup "
                "and HA burden handled by the platform.")
    ws["A2"].font = SUB_FONT
    ws.merge_cells("A2:F2")

    row = 4
    cols = ["Instance", "Database", "Blocker", "What it takes", "Est. Days", "Unlocks"]
    for i, h in enumerate(cols, start=1):
        ws.cell(row=row, column=i, value=h)
    _style_header(ws, row, len(cols))
    row += 1

    count = 0
    for a in assessments:
        if a.target != "SQL_VM" or not a.paas_unblock:
            continue
        total = sum(f.get("effort_days", 0) for f in a.paas_unblock)
        for f in a.paas_unblock:
            ws.cell(row=row, column=1, value=a.workload.instance).border = BORDER
            ws.cell(row=row, column=2, value=a.workload.database).border = BORDER
            ws.cell(row=row, column=3, value=f["label"]).border = BORDER
            c = ws.cell(row=row, column=4, value=f["remediation"])
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(row=row, column=5, value=f["effort_days"]).border = BORDER
            ws.cell(row=row, column=6, value="Azure SQL Managed Instance").border = BORDER
            row += 1
            count += 1
        c = ws.cell(row=row, column=4, value=f"Total to unblock {a.workload.database}")
        c.font = Font(bold=True, italic=True)
        ws.cell(row=row, column=5, value=total).font = Font(bold=True)
        row += 2

    if count == 0:
        ws.cell(row=row, column=1,
                value="No workloads require IaaS - the whole estate is PaaS eligible.").font = Font(color=GREEN, bold=True)

    _autosize(ws)
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 80
    ws.sheet_view.showGridLines = False


def _sheet_blockers(ws, assessments) -> None:
    cols = ["Instance", "Database", "Finding ID", "Finding", "Severity", "Status",
            "Blocks", "Evidence", "Remediation", "Remediation Days"]
    ws.append(cols)
    _style_header(ws, 1, len(cols))

    count = 0
    for a in assessments:
        resolved_ids = {id(f) for f in a.resolved_by_target}
        for f in sorted(a.all_findings(), key=lambda x: {"hard": 0, "medium": 1, "low": 2}[x["severity"]]):
            is_resolved = id(f) in resolved_ids
            status = "Resolved by target" if is_resolved else "Needs work"
            ws.append([
                a.workload.instance, a.workload.database, f["id"], f["label"],
                f["severity"].upper(), status, ", ".join(f["blocks"]) or "-",
                f["evidence"], f["remediation"],
                0 if is_resolved else f["effort_days"],
            ])
            count += 1
            r = ws.max_row
            ws.cell(row=r, column=5).font = Font(
                color={"HARD": RED, "MEDIUM": AMBER, "LOW": GREY}[f["severity"].upper()], bold=True)
            ws.cell(row=r, column=6).font = Font(color=GREEN if is_resolved else AMBER, bold=True)
            for c in (8, 9):
                ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")

    if count == 0:
        ws.append(["-", "-", "-", "No findings detected across the estate", "-", "-", "-", "-", "-", 0])
        count = 1

    _add_table(ws, "Findings", 1, count, len(cols))
    _autosize(ws)
    ws.column_dimensions["H"].width = 60
    ws.column_dimensions["I"].width = 70
    ws.freeze_panes = "C2"


def _sheet_sizing(ws, assessments, meta: dict) -> None:
    cur = meta["currency"]
    cols = ["Instance", "Database", "Target", "SKU", "Service Tier", "vCores",
            "Storage (GB)", f"Est. {cur}/month (list)", f"Est. {cur}/month (with AHB)",
            "Sizing Confidence", "Sizing Basis", "Price Basis"]
    ws.append(cols)
    _style_header(ws, 1, len(cols))

    for a in assessments:
        ws.append([
            a.workload.instance, a.workload.database, a.target_label, a.sku_name,
            a.service_tier, a.vcores, a.storage_gb,
            a.est_monthly_usd, a.est_monthly_ahb_usd, a.sizing_confidence,
            " ".join(a.sizing_notes), a.price_source,
        ])
        r = ws.max_row
        for c in (8, 9):
            ws.cell(row=r, column=c).number_format = '#,##0.00'
        ws.cell(row=r, column=10).font = Font(
            color={"high": GREEN, "medium": AMBER, "low": RED}[a.sizing_confidence], bold=True)
        for c in (11, 12):
            ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")

    total_row = ws.max_row + 1
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    for c in (8, 9):
        letter = get_column_letter(c)
        cell = ws.cell(row=total_row, column=c,
                       value=f"=SUM({letter}2:{letter}{total_row - 1})")
        cell.font = Font(bold=True)
        cell.number_format = '#,##0.00'

    _add_table(ws, "Sizing", 1, len(assessments), len(cols))
    _autosize(ws)
    ws.column_dimensions["K"].width = 75
    ws.column_dimensions["L"].width = 75
    ws.freeze_panes = "C2"


def _sheet_waves(ws, assessments, rules: dict) -> None:
    ws["A1"] = "Migration Wave Plan"
    ws["A1"].font = TITLE_FONT
    row = 3

    for wave in rules["waves"]:
        members = [a for a in assessments if a.wave == wave["id"]]
        if not members:
            continue

        low = sum(m.effort_days_low for m in members)
        high = sum(m.effort_days_high for m in members)
        size = sum(m.workload.db_size_gb or 0 for m in members)

        c = ws.cell(row=row, column=1,
                    value=f"{wave['name']}   -   {len(members)} workload(s), "
                          f"{size:,.0f} GB, {low}-{high} person-days")
        c.font = Font(bold=True, size=12, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLUE)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        row += 1

        c = ws.cell(row=row, column=1, value=wave["intent"])
        c.font = Font(italic=True, color=GREY)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        row += 1

        hdr = ["Instance", "Database", "Target", "Size (GB)", "Effort", "Days", "Key findings"]
        for i, h in enumerate(hdr, start=1):
            ws.cell(row=row, column=i, value=h)
        _style_header(ws, row, len(hdr))
        row += 1

        for m in sorted(members, key=lambda x: -(x.workload.db_size_gb or 0)):
            ws.cell(row=row, column=1, value=m.workload.instance).border = BORDER
            ws.cell(row=row, column=2, value=m.workload.database).border = BORDER
            ws.cell(row=row, column=3, value=m.target_label).border = BORDER
            ws.cell(row=row, column=4,
                    value=round(m.workload.db_size_gb, 1) if m.workload.db_size_gb else None).border = BORDER
            ws.cell(row=row, column=5, value=m.effort_band).border = BORDER
            ws.cell(row=row, column=6,
                    value=f"{m.effort_days_low}-{m.effort_days_high}").border = BORDER
            fc = ws.cell(row=row, column=7, value=_fmt_findings(m.all_findings()))
            fc.border = BORDER
            fc.alignment = Alignment(wrap_text=True, vertical="top")
            row += 1

        row += 1

    _autosize(ws)
    ws.column_dimensions["G"].width = 70
    ws.sheet_view.showGridLines = False


def _sheet_inventory(ws, assessments) -> None:
    cols = ["Source File", "Format", "Instance", "Database", "Version", "Edition",
            "OS", "Cores", "CPU Util %", "Memory (GB)", "Data (GB)", "Log (GB)",
            "Max IOPS", "Throughput (MB/s)", "Compat Level", "Collation",
            "Agent Jobs", "Linked Servers", "Raw Signals Captured"]
    ws.append(cols)
    _style_header(ws, 1, len(cols))

    for a in assessments:
        w = a.workload
        ws.append([
            w.source_file, w.source_format, w.instance, w.database, w.sql_version,
            w.sql_edition, w.os_name, w.cores, w.cpu_util_pct, w.memory_gb,
            w.db_size_gb, w.log_size_gb, w.max_iops, w.throughput_mbps,
            w.compat_level, w.collation, w.agent_job_count, w.linked_server_count,
            "\n".join(w.signals),
        ])
        ws.cell(row=ws.max_row, column=len(cols)).alignment = Alignment(wrap_text=True, vertical="top")

    _add_table(ws, "Inventory", 1, len(assessments), len(cols))
    _autosize(ws)
    ws.column_dimensions[get_column_letter(len(cols))].width = 90
    ws.freeze_panes = "E2"


def _sheet_assumptions(ws, rules: dict, meta: dict, warnings: list[str]) -> None:
    ws["A1"] = "Method, Assumptions and Limitations"
    ws["A1"].font = TITLE_FONT
    row = 3

    def block(title: str, lines: list[str]) -> None:
        nonlocal row
        c = ws.cell(row=row, column=1, value=title)
        c.font = Font(bold=True, size=12, color=NAVY)
        row += 1
        for line in lines:
            cell = ws.cell(row=row, column=1, value=line)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            row += 1
        row += 1

    s = rules["sizing"]
    block("How the target platform is chosen", [
        "1. Every finding in the source artifacts is matched against the rule pack in reference/rules.json.",
        "2. Each finding eliminates the platforms it is incompatible with - for example SQL Agent jobs eliminate Azure SQL Database but not SQL Managed Instance.",
        "3. The cheapest-to-operate surviving platform wins, in preference order: SQL Database, then SQL Managed Instance, then SQL Server on Azure VM.",
        "4. SQL Server on Azure VM is the universal fallback and is only recommended when every PaaS option has been eliminated.",
        "5. Microsoft Fabric is never auto-assigned. It is flagged as a re-platform candidate for human confirmation, because moving an OLTP database to a warehouse changes the application contract.",
    ])

    block("How sizing is derived", [
        f"vCores = source cores x average CPU utilisation x {s['default_peak_to_avg_ratio']} peak-to-average ratio x {s['headroom_factor']} headroom, rounded up to the next available SKU.",
        f"Where CPU utilisation is absent from the source data, {s['default_cpu_util_pct']:.0f}% is assumed and the sizing confidence is downgraded.",
        f"Storage = (data + log) x {s['storage_growth_factor']} growth allowance, rounded up to the next provisioning step.",
        f"Business Critical is recommended above {s['business_critical_triggers']['min_iops']:,} IOPS or {s['business_critical_triggers']['min_throughput_mbps']} MB/s; otherwise General Purpose.",
        f"Memory is modelled at {s['memory_per_vcore_gb']} GB per vCore, consistent with the Gen5 hardware family.",
    ])

    block("Pricing basis", [
        f"Prices are live Azure retail list prices for region '{meta['region']}' in {meta['currency']}, retrieved from the public Azure Retail Prices API on {meta['generated']}.",
        f"Monthly figures assume {rules['pricing']['hours_per_month']} hours of continuous running.",
        "Azure Hybrid Benefit figures are indicative: they remove the SQL licence component from compute but must be validated against the customer's Software Assurance position.",
        "List prices exclude any MACC, EA or CSP discounting. Use these for shaping and relative comparison, not as a quote.",
        "SQL Server on Azure VM is deliberately not priced here, because the cost depends on VM family, OS licensing and storage configuration chosen at design time.",
    ])

    block("Known limitations - read before presenting", [
        "This assessment is only as good as the artifacts supplied. It infers from what the source data states and does not connect to the customer estate.",
        "Static assessments (an inventory spreadsheet or a DMA compatibility report) carry no performance telemetry, so sizing falls back on assumptions. A performance-based Azure Migrate assessment over a representative 7-30 day window is strongly recommended before any commercial commitment.",
        "Application-level compatibility, third-party ISV support statements and licence entitlements are out of scope and must be confirmed separately.",
        "Effort bands are planning estimates derived from workload complexity. They are not a statement of work.",
        "Every recommendation is advisory and should be reviewed by the account's data platform specialist before it reaches the customer.",
    ])

    if warnings:
        block("Ingestion warnings from this run", warnings)

    block("Provenance", [
        f"Rule pack version: {rules['_meta']['version']} (updated {rules['_meta']['updated']})",
        f"Source artifacts: {meta['sources']}",
        f"Generated: {meta['generated']}",
    ])

    ws.column_dimensions["A"].width = 140
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------
# BoQ CSV
# --------------------------------------------------------------------------

def write_boq(path: Path, assessments, meta: dict) -> None:
    """Flat, machine-readable hand-off for BoQ and pricing-calculator tooling."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow([
            "Customer", "Region", "Currency", "Instance", "Database", "Service",
            "SKU", "ServiceTier", "vCores", "StorageGB", "Quantity",
            "EstMonthlyList", "EstMonthlyWithAHB", "Wave", "EffortBand",
            "EffortDaysLow", "EffortDaysHigh", "HardBlockers", "Confidence",
        ])
        for a in assessments:
            writer.writerow([
                meta["customer"], meta["region"], meta["currency"],
                a.workload.instance, a.workload.database, a.target_label,
                a.sku_name, a.service_tier, a.vcores, a.storage_gb, 1,
                a.est_monthly_usd if a.est_monthly_usd is not None else "",
                a.est_monthly_ahb_usd if a.est_monthly_ahb_usd is not None else "",
                a.wave, a.effort_band, a.effort_days_low, a.effort_days_high,
                len(a.blockers), a.confidence,
            ])


# --------------------------------------------------------------------------
# HTML report
# --------------------------------------------------------------------------

def write_html(path: Path, assessments, summary: dict, meta: dict,
               warnings: list[str], rules: dict) -> None:
    e = html.escape
    s = summary
    paas_pct = (s["paas_eligible"] / s["total_workloads"] * 100) if s["total_workloads"] else 0

    cards = [
        ("Workloads assessed", f"{s['total_workloads']:,}", ""),
        ("PaaS eligible", f"{paas_pct:.0f}%", f"{s['paas_eligible']} of {s['total_workloads']}"),
        ("Total data volume", f"{s['total_size_gb']:,.0f} GB", ""),
        ("Needs real remediation", f"{s.get('workloads_with_residual_work', 0):,}",
         f"{s.get('findings_resolved_by_target', 0)} findings resolved by target choice"),
        ("Out of support", f"{s['eos_workloads']:,}", "ESU exposure today"),
        ("Effort", f"{s['effort_days_low']:,}-{s['effort_days_high']:,}", "person-days"),
    ]
    if s["est_monthly_usd"]:
        cards.append(("Est. run-rate", f"{s['est_monthly_usd']:,.0f}",
                      f"{meta['currency']}/mo list, {s['est_monthly_ahb_usd']:,.0f} with AHB"))

    card_html = "".join(
        f'<div class="card"><div class="card-label">{e(l)}</div>'
        f'<div class="card-value">{e(v)}</div>'
        f'<div class="card-sub">{e(sub)}</div></div>'
        for l, v, sub in cards
    )

    target_rows = "".join(
        f'<tr><td><span class="pill pill-{_slug(t)}">{e(t)}</span></td>'
        f'<td class="num">{c}</td>'
        f'<td class="bar-cell"><div class="bar" style="width:{(c / s["total_workloads"] * 100):.1f}%"></div></td></tr>'
        for t, c in sorted(s["by_target"].items(), key=lambda kv: -kv[1])
    )

    blocker_rows = "".join(
        f'<tr><td>{e(label)}</td><td class="num">{count}</td></tr>'
        for label, count in s["top_blockers"]
    ) or '<tr><td colspan="2">No findings detected.</td></tr>'

    wave_html = ""
    for wave in rules["waves"]:
        members = [a for a in assessments if a.wave == wave["id"]]
        if not members:
            continue
        low = sum(m.effort_days_low for m in members)
        high = sum(m.effort_days_high for m in members)
        size = sum(m.workload.db_size_gb or 0 for m in members)
        rows = "".join(
            f'<tr><td>{e(m.workload.instance)}</td><td><strong>{e(m.workload.database)}</strong></td>'
            f'<td><span class="pill pill-{_slug(m.target_label)}">{e(m.target_label)}</span></td>'
            f'<td class="num">{(m.workload.db_size_gb or 0):,.0f}</td>'
            f'<td class="num">{m.vcores}</td>'
            f'<td><span class="band band-{m.effort_band}">{m.effort_band}</span> '
            f'{m.effort_days_low}-{m.effort_days_high}d</td>'
            f'<td class="findings">{_findings_html(m)}</td></tr>'
            for m in sorted(members, key=lambda x: -(x.workload.db_size_gb or 0))
        )
        wave_html += f"""
        <div class="wave">
          <div class="wave-head">
            <h3>{e(wave['name'])}</h3>
            <div class="wave-stats">{len(members)} workloads &middot; {size:,.0f} GB &middot; {low}-{high} person-days</div>
          </div>
          <p class="wave-intent">{e(wave['intent'])}</p>
          <table>
            <thead><tr><th>Instance</th><th>Database</th><th>Target</th><th class="num">Size GB</th>
            <th class="num">vCores</th><th>Effort</th><th>Findings</th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>"""

    unblock_rows = ""
    for a in assessments:
        if a.target != "SQL_VM" or not a.paas_unblock:
            continue
        total = sum(f.get("effort_days", 0) for f in a.paas_unblock)
        items = "".join(f"<li>{e(f['label'])} &mdash; <span style='color:#6b7280'>{e(f['remediation'])}</span></li>"
                        for f in a.paas_unblock)
        unblock_rows += (
            f'<tr><td>{e(a.workload.instance)}</td><td><strong>{e(a.workload.database)}</strong></td>'
            f'<td><ul style="margin:0;padding-left:18px;font-size:12px">{items}</ul></td>'
            f'<td class="num">{total}d</td></tr>'
        )
    unblock_html = ""
    if unblock_rows:
        unblock_html = f"""
  <section><h2>Path to PaaS</h2>
    <p style="font-size:13px;color:var(--muted);margin-top:0">
      These workloads landed on SQL Server on Azure VM. Remove the blockers below and each becomes
      eligible for Azure SQL Managed Instance, handing patching, backup and high availability to the platform.
    </p>
    <table><thead><tr><th>Instance</th><th>Database</th><th>What stands in the way</th>
    <th class="num">Est. days</th></tr></thead><tbody>{unblock_rows}</tbody></table>
  </section>"""

    warn_html = ""
    if warnings or s["low_confidence"]:
        items = list(warnings)
        if s["low_confidence"]:
            items.insert(0, f"{s['low_confidence']} workload(s) were sized on assumed values because the source "
                            f"data lacked CPU or size telemetry. Treat those sizings as indicative only.")
        warn_html = ('<div class="warn"><strong>Caveats</strong><ul>'
                     + "".join(f"<li>{e(i)}</li>" for i in items) + "</ul></div>")

    doc = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{e(meta['customer'])} - SQL Estate Migration Assessment</title>
<style>
  :root {{
    --navy:#1f3864; --blue:#2e75b6; --light:#ddebf7; --ink:#1a1a1a;
    --muted:#6b7280; --line:#e5e7eb; --red:#c00000; --amber:#bf8f00; --green:#375623;
  }}
  * {{ box-sizing:border-box; }}
  body {{ margin:0; font-family:"Segoe UI",system-ui,-apple-system,sans-serif;
         color:var(--ink); background:#f6f8fb; line-height:1.55; }}
  header {{ background:linear-gradient(135deg,var(--navy),var(--blue));
            color:#fff; padding:36px 40px; }}
  header h1 {{ margin:0 0 6px; font-size:26px; font-weight:600; }}
  header .meta {{ opacity:.85; font-size:13px; }}
  main {{ max-width:1400px; margin:0 auto; padding:28px 40px 60px; }}
  section {{ background:#fff; border:1px solid var(--line); border-radius:10px;
             padding:24px 28px; margin-bottom:22px; }}
  h2 {{ font-size:17px; color:var(--navy); margin:0 0 16px;
        border-bottom:2px solid var(--light); padding-bottom:8px; }}
  h3 {{ margin:0; font-size:15px; color:var(--navy); }}
  .cards {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(170px,1fr)); gap:14px; }}
  .card {{ background:#fff; border:1px solid var(--line); border-left:4px solid var(--blue);
           border-radius:8px; padding:14px 16px; }}
  .card-label {{ font-size:11px; text-transform:uppercase; letter-spacing:.5px; color:var(--muted); }}
  .card-value {{ font-size:26px; font-weight:600; color:var(--navy); margin:4px 0 2px; }}
  .card-sub {{ font-size:11px; color:var(--muted); }}
  table {{ width:100%; border-collapse:collapse; font-size:13px; }}
  th {{ background:var(--navy); color:#fff; text-align:left; padding:9px 11px;
        font-weight:600; font-size:12px; }}
  td {{ padding:9px 11px; border-bottom:1px solid var(--line); vertical-align:top; }}
  tr:nth-child(even) td {{ background:#fafbfd; }}
  .num {{ text-align:right; font-variant-numeric:tabular-nums; }}
  .pill {{ display:inline-block; padding:3px 9px; border-radius:11px;
           font-size:11px; font-weight:600; white-space:nowrap; }}
  .pill-azure-sql-database {{ background:#c6e0b4; color:#2d4a1e; }}
  .pill-azure-sql-managed-instance {{ background:#bdd7ee; color:#1b3a5c; }}
  .pill-sql-server-on-azure-vm {{ background:#ffe699; color:#6b5200; }}
  .pill-microsoft-fabric-warehouse-lakehouse {{ background:#e4dfec; color:#432c6b; }}
  .band {{ display:inline-block; width:24px; text-align:center; border-radius:4px;
           font-size:11px; font-weight:700; color:#fff; margin-right:4px; }}
  .band-S {{ background:var(--green); }} .band-M {{ background:var(--blue); }}
  .band-L {{ background:var(--amber); }} .band-XL {{ background:var(--red); }}
  .findings {{ font-size:11px; color:var(--muted); max-width:380px; }}
  .f-hard {{ color:var(--red); font-weight:600; }}
  .f-medium {{ color:var(--amber); }}
  .bar-cell {{ width:40%; }}
  .bar {{ height:14px; background:var(--blue); border-radius:3px; min-width:2px; }}
  .wave {{ margin-bottom:26px; }}
  .wave-head {{ display:flex; justify-content:space-between; align-items:baseline;
                background:var(--light); padding:10px 14px; border-radius:6px 6px 0 0; }}
  .wave-stats {{ font-size:12px; color:var(--navy); font-weight:600; }}
  .wave-intent {{ font-size:12px; color:var(--muted); font-style:italic; margin:8px 0 10px; }}
  .warn {{ background:#fff8e6; border:1px solid #f0d58c; border-left:4px solid var(--amber);
           border-radius:8px; padding:14px 18px; font-size:13px; margin-bottom:22px; }}
  .warn ul {{ margin:8px 0 0; padding-left:20px; }}
  .two-col {{ display:grid; grid-template-columns:1fr 1fr; gap:22px; }}
  @media(max-width:900px) {{ .two-col {{ grid-template-columns:1fr; }} main {{ padding:20px; }} }}
  footer {{ text-align:center; color:var(--muted); font-size:11px; padding:20px; }}
</style></head><body>
<header>
  <h1>SQL Estate Migration Assessment</h1>
  <div class="meta">{e(meta['customer'])} &nbsp;&middot;&nbsp; target region {e(meta['region'])}
   &nbsp;&middot;&nbsp; generated {e(meta['generated'])}
   &nbsp;&middot;&nbsp; sources: {e(meta['sources'])}</div>
</header>
<main>
  {warn_html}
  <section><h2>Estate at a glance</h2><div class="cards">{card_html}</div></section>
  <div class="two-col">
    <section><h2>Recommended target platform</h2>
      <table><thead><tr><th>Target</th><th class="num">Workloads</th><th>Share</th></tr></thead>
      <tbody>{target_rows}</tbody></table></section>
    <section><h2>Most common findings</h2>
      <table><thead><tr><th>Finding</th><th class="num">Affected</th></tr></thead>
      <tbody>{blocker_rows}</tbody></table></section>
  </div>
  <section><h2>Migration wave plan</h2>{wave_html}</section>
  {unblock_html}
  <section><h2>Method and assumptions</h2>
    <p style="font-size:13px;color:var(--muted)">
      Target selection eliminates platforms that each detected finding rules out, then takes the
      cheapest-to-operate survivor in the order Azure SQL Database &rarr; SQL Managed Instance &rarr;
      SQL Server on Azure VM. Microsoft Fabric is flagged as a candidate only and never auto-assigned.
      Sizing derives vCores from source cores &times; CPU utilisation &times;
      {rules['sizing']['default_peak_to_avg_ratio']} peak ratio &times;
      {rules['sizing']['headroom_factor']} headroom. Prices are live Azure retail list prices for
      {e(meta['region'])} and exclude customer-specific discounting, so they are suitable for shaping
      and comparison but are not a quote. Full assumptions, limitations and per-finding evidence are in
      the <strong>Method &amp; Assumptions</strong> sheet of the accompanying workbook.
    </p></section>
</main>
<footer>Estate-to-Target Migration Assessor &middot; rule pack v{e(rules['_meta']['version'])}
 &middot; advisory output, review before sharing with a customer</footer>
</body></html>"""

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(doc, encoding="utf-8")


def _slug(text: str) -> str:
    return "".join(c if c.isalnum() else "-" for c in text.lower()).strip("-").replace("--", "-")


def _findings_html(a) -> str:
    items = a.all_findings()
    if not items:
        return '<span style="color:#375623">clean</span>'
    out = []
    for f in items[:4]:
        cls = "f-hard" if f["severity"] == "hard" else ("f-medium" if f["severity"] == "medium" else "")
        out.append(f'<span class="{cls}">{html.escape(f["label"])}</span>')
    if len(items) > 4:
        out.append(f"<span>+{len(items) - 4} more</span>")
    return "<br>".join(out)


def write_json(path: Path, assessments, summary: dict, meta: dict, warnings: list[str]) -> None:
    """Structured output so an agent can reason over the result conversationally."""
    payload = {
        "meta": meta,
        "summary": summary,
        "ingestion_warnings": warnings,
        "assessments": [
            {
                "instance": a.workload.instance,
                "database": a.workload.database,
                "source_file": a.workload.source_file,
                "source_format": a.workload.source_format,
                "sql_version": a.workload.sql_version,
                "size_gb": a.workload.db_size_gb,
                "target": a.target,
                "target_label": a.target_label,
                "confidence": a.confidence,
                "rationale": a.rationale,
                "eliminated_platforms": a.eliminated,
                "fabric_candidate": a.fabric_candidate,
                "non_production": a.is_non_production,
                "end_of_support": a.eos_flag,
                "blockers": a.blockers,
                "warnings": a.warnings,
                "residual_findings": a.residual,
                "resolved_by_target": a.resolved_by_target,
                "paas_unblock_path": a.paas_unblock,
                "capacity_notes": a.capacity_notes,
                "sku": a.sku_name,
                "vcores": a.vcores,
                "service_tier": a.service_tier,
                "storage_gb": a.storage_gb,
                "sizing_notes": a.sizing_notes,
                "sizing_confidence": a.sizing_confidence,
                "effort_band": a.effort_band,
                "effort_days": [a.effort_days_low, a.effort_days_high],
                "remediation_days": a.remediation_days,
                "wave": a.wave,
                "wave_name": a.wave_name,
                "est_monthly": a.est_monthly_usd,
                "est_monthly_ahb": a.est_monthly_ahb_usd,
            }
            for a in assessments
        ],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
