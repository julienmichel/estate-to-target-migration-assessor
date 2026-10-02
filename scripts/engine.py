"""
Decision engine for the Estate-to-Target Migration Assessor.

Four stages, in order, because each depends on the last:

  1. detect_blockers  - match source signals against the rule pack
  2. choose_target    - eliminate platforms the blockers rule out, then pick
                        the cheapest surviving option
  3. size             - derive vCores / tier / storage from observed telemetry
  4. plan_waves       - band the effort and sequence the estate

Every verdict carries its reasons. An assessment a customer cannot argue with
is worthless, so nothing here is a black box.
"""

from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

from ingest import Workload

RULES_PATH = Path(__file__).resolve().parent.parent / "reference" / "rules.json"


def load_rules(path: Path | None = None) -> dict:
    return json.loads((path or RULES_PATH).read_text(encoding="utf-8"))


# --------------------------------------------------------------------------
# Result record
# --------------------------------------------------------------------------

@dataclass
class Assessment:
    workload: Workload

    target: str = "SQL_MI"
    target_label: str = ""
    confidence: str = "medium"
    rationale: list[str] = field(default_factory=list)

    blockers: list[dict] = field(default_factory=list)
    warnings: list[dict] = field(default_factory=list)
    eliminated: dict[str, list[str]] = field(default_factory=dict)

    # Findings that still require work after the target is chosen, versus those
    # the target resolves natively. Conflating these is the classic way to
    # overstate migration effort by a factor of three.
    residual: list[dict] = field(default_factory=list)
    resolved_by_target: list[dict] = field(default_factory=list)
    paas_unblock: list[dict] = field(default_factory=list)

    is_non_production: bool = False
    capacity_notes: list[str] = field(default_factory=list)

    fabric_candidate: bool = False
    eos_flag: bool = False
    eos_date: str = ""

    vcores: int = 0
    service_tier: str = "GeneralPurpose"
    storage_gb: int = 0
    sku_name: str = ""
    sizing_notes: list[str] = field(default_factory=list)
    sizing_confidence: str = "medium"

    effort_band: str = "M"
    effort_days_low: int = 0
    effort_days_high: int = 0
    remediation_days: int = 0

    wave: int = 2
    wave_name: str = ""
    dependency_group: str = ""

    est_monthly_usd: float | None = None
    est_monthly_ahb_usd: float | None = None
    price_source: str = ""

    def all_findings(self) -> list[dict]:
        return self.blockers + self.warnings


# --------------------------------------------------------------------------
# Stage 1 - blocker detection
# --------------------------------------------------------------------------

def _haystack(w: Workload) -> str:
    """Everything we know about a workload, lower-cased, for pattern matching."""
    parts = [
        w.database, w.workload_type, w.collation, w.sql_edition,
        w.os_name, w.compat_level, *w.signals,
    ]
    parts.extend(str(v) for v in w.extra.values())
    return " ".join(str(p) for p in parts if p).lower()


def detect_blockers(w: Workload, rules: dict) -> tuple[list[dict], list[dict]]:
    """Returns (blockers, warnings) - hard findings vs advisory ones."""
    hay = _haystack(w)
    blockers: list[dict] = []
    warnings: list[dict] = []
    seen: set[str] = set()

    for rule in rules["blockers"]:
        if rule["id"] in seen:
            continue

        hit = any(token.lower() in hay for token in rule["detect"])

        # Structured counters are stronger evidence than free text.
        if rule["id"] == "AGENT_JOBS" and w.agent_job_count > 0:
            hit = True
        if rule["id"] == "LINKED_SERVER" and w.linked_server_count > 0:
            hit = True

        if not hit:
            continue
        seen.add(rule["id"])

        blocks = list(rule.get("blocks", []))

        # CLR escalates when the assembly is UNSAFE or EXTERNAL_ACCESS.
        esc = rule.get("escalate_if", {})
        if esc.get("unsafe") and re.search(r"unsafe|external[_ ]access", hay):
            blocks = list(esc["unsafe"])

        finding = {
            "id": rule["id"],
            "label": rule["label"],
            "severity": rule["severity"],
            "blocks": blocks,
            "remediation": rule["remediation"],
            "effort_days": rule.get("effort_days", 0),
            "evidence": _evidence_for(rule, w),
            "detaches": rule.get("detaches_rather_than_blocks", False),
            "dependency": rule.get("creates_dependency_group", False),
            "opportunity": rule.get("is_opportunity_signal", False),
        }

        if rule["severity"] == "hard" and blocks:
            blockers.append(finding)
        else:
            warnings.append(finding)

    # Version / end-of-support is derived, not pattern-matched.
    vinfo = rules["version_support"].get(w.sql_version)
    if vinfo and not vinfo["supported"]:
        eos = vinfo["eos"]
        passed = datetime.strptime(eos, "%Y-%m-%d").date() <= date.today()
        warnings.append({
            "id": "END_OF_SUPPORT",
            "label": f"SQL Server {w.sql_version} {'is out of support' if passed else 'reaches end of support'} ({eos})",
            "severity": "medium",
            "blocks": [],
            "remediation": "Out-of-support versions carry ESU cost and security exposure. "
                           "A PaaS target removes the patching obligation entirely.",
            "effort_days": 0,
            "evidence": f"Detected version: SQL Server {w.sql_version}",
            "detaches": False,
            "dependency": False,
            "opportunity": True,
        })

    # Platform ceilings become blockers in their own right.
    size = w.db_size_gb or 0
    if size > rules["platform_limits"]["SQL_DB"]["max_db_size_gb"]:
        blockers.append({
            "id": "SIZE_SQL_DB",
            "label": f"Database exceeds the Azure SQL Database single-database ceiling ({size:,.0f} GB)",
            "severity": "hard",
            "blocks": ["SQL_DB"],
            "remediation": "Use SQL MI, or Azure SQL Database Hyperscale which lifts the size ceiling.",
            "effort_days": 0,
            "evidence": f"Reported size {size:,.1f} GB",
            "detaches": False, "dependency": False, "opportunity": False,
        })
    if size > rules["platform_limits"]["SQL_MI"]["max_instance_storage_gb"]:
        blockers.append({
            "id": "SIZE_SQL_MI",
            "label": f"Database exceeds the SQL Managed Instance storage ceiling ({size:,.0f} GB)",
            "severity": "hard",
            "blocks": ["SQL_MI"],
            "remediation": "Move to SQL Server on Azure VM, or partition and archive cold data first.",
            "effort_days": 15,
            "evidence": f"Reported size {size:,.1f} GB",
            "detaches": False, "dependency": False, "opportunity": False,
        })

    return blockers, warnings


def _evidence_for(rule: dict, w: Workload) -> str:
    """Quote the source text that triggered a rule, so findings are traceable."""
    if rule["id"] == "AGENT_JOBS" and w.agent_job_count > 0:
        return f"{w.agent_job_count} SQL Agent job(s) reported"
    if rule["id"] == "LINKED_SERVER" and w.linked_server_count > 0:
        return f"{w.linked_server_count} linked server(s) reported"
    for sig in w.signals:
        low = sig.lower()
        if any(tok.lower() in low for tok in rule["detect"]):
            return sig[:300]
    return "Matched on inventory metadata"


def _is_non_production(w: Workload, rules: dict) -> bool:
    """Dev, test and UAT estates behave differently and must not inflate the plan."""
    np_rules = rules.get("non_production", {})
    edition = (w.sql_edition or "").lower()
    if any(v in edition for v in np_rules.get("edition_values", [])):
        return True
    wtype = (w.workload_type or "").lower()
    if any(v == wtype or v in wtype.split() for v in np_rules.get("workload_type_values", [])):
        return True
    tokens = set(re.split(r"[^a-z0-9]+", f"{w.instance} {w.database}".lower()))
    return any(p in tokens for p in np_rules.get("name_patterns", []))


def partition_findings(a: Assessment) -> None:
    """Split findings into work that remains versus work the target choice removes.

    This is the single most important correction in the engine. SQL Agent jobs
    'block' Azure SQL Database, but SQL Managed Instance runs Agent natively - so
    once MI is the recommendation there is no re-architecting to do, only the
    migration itself. Counting those as remediation would massively overstate
    effort and push clean workloads into the complex waves.
    """
    residual: list[dict] = []
    resolved: list[dict] = []

    for f in a.all_findings():
        blocks = f.get("blocks", [])
        if f.get("detaches"):
            # Separating SSIS/SSRS/SSAS is real work on any target.
            residual.append(f)
        elif not blocks:
            # Universal findings: deprecated syntax, collation, Windows auth, EOS.
            residual.append(f)
        elif a.target in blocks:
            residual.append(f)  # defensive - should be impossible by construction
        else:
            resolved.append(f)

    a.residual = residual
    a.resolved_by_target = resolved

    # For an IaaS recommendation, spell out what it would take to reach PaaS.
    # This is the upsell conversation, and it is the first thing an architect asks.
    if a.target == "SQL_VM":
        a.paas_unblock = [
            f for f in a.all_findings()
            if "SQL_MI" in f.get("blocks", []) and not f.get("detaches")
        ]


def validate_capacity(a: Assessment, rules: dict) -> bool:
    """Check PROVISIONED storage against the platform ceiling, not just raw size.

    A 14 TB database sits under the 16 TB SQL MI limit, but once log and growth
    allowance are added the provisioned requirement is ~20 TB and will not fit.
    Returns True when the target had to change and the workload needs re-sizing.
    """
    limits = rules["platform_limits"]

    if a.target == "SQL_MI":
        ceiling = limits["SQL_MI"]["max_instance_storage_gb"]
        if a.storage_gb >= ceiling:
            required = (a.workload.db_size_gb or 0) + (a.workload.log_size_gb or 0)
            required *= rules["sizing"]["storage_growth_factor"]
            if required > ceiling:
                a.blockers.append({
                    "id": "CAPACITY_SQL_MI",
                    "label": f"Provisioned storage requirement ({required:,.0f} GB) exceeds the "
                             f"SQL Managed Instance ceiling ({ceiling:,} GB)",
                    "severity": "hard",
                    "blocks": ["SQL_MI"],
                    "remediation": "Archive or partition cold data to bring the footprint under the instance "
                                   "ceiling, split across multiple instances, or move to SQL Server on Azure VM.",
                    "effort_days": 15,
                    "evidence": f"{a.workload.db_size_gb:,.0f} GB data plus log, with "
                                f"{rules['sizing']['storage_growth_factor']}x growth allowance",
                    "detaches": False, "dependency": False, "opportunity": False,
                })
                return True

    if a.target == "SQL_DB":
        ceiling = limits["SQL_DB"]["max_db_size_gb"]
        if a.storage_gb > ceiling:
            a.capacity_notes.append(
                f"Provisioned storage ({a.storage_gb:,} GB) exceeds the standard Azure SQL Database "
                f"ceiling ({ceiling:,} GB). Azure SQL Database Hyperscale lifts this limit and should "
                f"be specified for this workload."
            )
    return False


# --------------------------------------------------------------------------
# Stage 2 - target selection
# --------------------------------------------------------------------------

def _looks_analytical(w: Workload, rules: dict) -> bool:
    fh = rules["fabric_heuristics"]
    name = (w.database or "").lower()
    wtype = (w.workload_type or "").lower()

    if any(v in wtype for v in fh["workload_type_values"]):
        return True
    tokens = re.split(r"[^a-z0-9]+", name)
    return any(t in fh["name_patterns"] for t in tokens if t)


def choose_target(w: Workload, blockers: list[dict], warnings: list[dict],
                  rules: dict) -> dict:
    """Eliminate what the evidence rules out, then take the cheapest survivor."""
    candidates = ["SQL_DB", "SQL_MI", "SQL_VM"]
    eliminated: dict[str, list[str]] = {}
    rationale: list[str] = []

    for finding in blockers + warnings:
        if finding.get("detaches"):
            continue  # BI services get separated, they do not veto the engine target
        for platform in finding.get("blocks", []):
            if platform in candidates:
                candidates.remove(platform)
            eliminated.setdefault(platform, []).append(finding["label"])

    # SQL on VM is the universal fallback - it can always host the workload.
    if not candidates:
        candidates = ["SQL_VM"]
        rationale.append("Every PaaS option is ruled out by hard blockers; SQL Server on Azure VM is the only viable lift-and-shift target.")

    target = min(candidates, key=lambda c: rules["targets"][c]["preference"])

    if target == "SQL_DB":
        rationale.append("No instance-scoped dependencies found. The database is self-contained and suits a fully managed single database.")
    elif target == "SQL_MI":
        if "SQL_DB" in eliminated:
            rationale.append("Instance-level features rule out Azure SQL Database: " + "; ".join(eliminated["SQL_DB"][:3]) + ".")
        rationale.append("SQL Managed Instance preserves near-complete SQL Server surface area while remaining fully managed.")
    elif target == "SQL_VM":
        reasons = eliminated.get("SQL_MI", []) + eliminated.get("SQL_DB", [])
        if reasons:
            rationale.append("PaaS ruled out by: " + "; ".join(dict.fromkeys(reasons))[:400] + ".")
        rationale.append("SQL Server on Azure VM retains full OS and instance control. Revisit PaaS once these dependencies are remediated.")

    # Fabric is only ever a flagged candidate - never silently assigned.
    fabric = False
    fh = rules["fabric_heuristics"]
    if _looks_analytical(w, rules) and (w.db_size_gb or 0) >= fh["min_size_gb_for_fabric"]:
        fabric = True
        rationale.append(
            f"Naming and size ({w.db_size_gb:,.0f} GB) suggest an analytical rather than transactional workload. "
            "Flagged as a Microsoft Fabric re-platform candidate - confirm with the application owner before committing, "
            "as this changes the application contract."
        )

    hard_count = len(blockers)
    if hard_count == 0 and (w.cores or w.db_size_gb):
        confidence = "high"
    elif hard_count > 2 or not (w.cores or w.db_size_gb):
        confidence = "low"
    else:
        confidence = "medium"

    return {
        "target": target,
        "target_label": rules["targets"][target]["label"],
        "rationale": rationale,
        "eliminated": eliminated,
        "fabric_candidate": fabric,
        "confidence": confidence,
    }


# --------------------------------------------------------------------------
# Stage 3 - sizing
# --------------------------------------------------------------------------

def _ladder_up(value: float, ladder: list[int]) -> int:
    for step in ladder:
        if value <= step:
            return step
    return ladder[-1]


def size(w: Workload, target: str, rules: dict, non_production: bool = False) -> dict:
    s = rules["sizing"]
    notes: list[str] = []
    assumed = 0

    cores = w.cores
    if cores is None or cores <= 0:
        cores = 8.0
        assumed += 1
        notes.append("No core count in the source data; assumed 8 cores. Replace with real telemetry before quoting.")

    util = w.cpu_util_pct
    if util is None or util <= 0:
        util = s["default_cpu_util_pct"]
        assumed += 1
        notes.append(f"No CPU utilisation captured; assumed {util:.0f}% average. "
                     "An Azure Migrate performance-based assessment would tighten this considerably.")
    elif util > 100:
        util = 100.0

    peak = util * s["default_peak_to_avg_ratio"]
    peak = min(peak, 100.0)
    required = (cores * peak / 100.0) * s["headroom_factor"]

    if target == "SQL_MI":
        ladder = s["mi_vcore_ladder"]
    elif target == "SQL_VM":
        ladder = s.get("vm_vcore_ladder", s["vcore_ladder"])
    else:
        ladder = s["vcore_ladder"]
    vcores = _ladder_up(max(required, 2.0), ladder)
    notes.append(
        f"Sized from {cores:.0f} source cores at {util:.0f}% average "
        f"({peak:.0f}% modelled peak) with a {s['headroom_factor']:.2f}x headroom factor -> {vcores} vCores."
    )

    data = w.db_size_gb
    if data is None or data <= 0:
        data = 100.0
        assumed += 1
        notes.append("No database size in the source data; assumed 100 GB.")
    log = w.log_size_gb or (data * 0.25)
    storage = _ladder_up((data + log) * s["storage_growth_factor"], s["storage_ladder_gb"])
    notes.append(
        f"Storage from {data:,.0f} GB data plus {log:,.0f} GB log at "
        f"{s['storage_growth_factor']:.1f}x growth -> {storage:,} GB provisioned."
    )

    bc = s["business_critical_triggers"]
    tier = "GeneralPurpose"
    iops = w.max_iops or 0
    tput = w.throughput_mbps or 0

    if non_production:
        notes.append("Non-production workload: held at General Purpose regardless of observed IOPS. "
                     "Paying for Business Critical redundancy on a dev or test database is waste.")
    elif iops >= bc["min_iops"]:
        tier = "BusinessCritical"
        notes.append(f"{iops:,.0f} observed IOPS exceeds the {bc['min_iops']:,} threshold; Business Critical recommended for local NVMe storage.")
    elif tput >= bc["min_throughput_mbps"]:
        tier = "BusinessCritical"
        notes.append(f"{tput:,.0f} MB/s observed throughput exceeds the {bc['min_throughput_mbps']} MB/s threshold; Business Critical recommended.")
    else:
        gp_iops = vcores * s["iops_per_vcore"]["general_purpose"]
        notes.append(
            f"General Purpose selected. It delivers roughly {gp_iops:,} IOPS at {vcores} vCores"
            + (f", against {iops:,.0f} observed." if iops else "; no IOPS telemetry was supplied to challenge this.")
        )

    mem = vcores * s["memory_per_vcore_gb"]
    if w.memory_gb and w.memory_gb > mem * 1.5:
        notes.append(
            f"Source instance has {w.memory_gb:,.0f} GB RAM but {vcores} vCores only provides about {mem:,.0f} GB. "
            "If the workload is memory-bound, size up or consider a memory-optimised VM instead."
        )

    if target == "SQL_DB":
        sku = f"GP_Gen5_{vcores}" if tier == "GeneralPurpose" else f"BC_Gen5_{vcores}"
    elif target == "SQL_MI":
        sku = f"MI_{'GP' if tier == 'GeneralPurpose' else 'BC'}_Gen5_{vcores}"
    else:
        # Tiers are a PaaS concept. On IaaS the equivalent lever is the VM family
        # and disk type, so report that instead of a meaningless service tier.
        family = "E" if mem > vcores * 4 else "D"
        sku = f"Standard_{family}{vcores}ds_v5"
        tier = "IaaS - Premium SSD v2" if iops < bc["min_iops"] else "IaaS - Ultra Disk"
        notes.append(
            f"VM family is indicative ({sku}). Validate availability in the target region and confirm "
            f"the disk configuration delivers the required {iops:,.0f} IOPS."
            if iops else
            f"VM family is indicative ({sku}). Validate availability in the target region."
        )

    confidence = "high" if assumed == 0 else ("medium" if assumed == 1 else "low")

    return {
        "vcores": vcores,
        "service_tier": tier,
        "storage_gb": storage,
        "sku_name": sku,
        "sizing_notes": notes,
        "sizing_confidence": confidence,
    }


# --------------------------------------------------------------------------
# Stage 4 - effort and waves
# --------------------------------------------------------------------------

def band_effort(a: Assessment, rules: dict) -> None:
    """Band effort on work that actually remains, not on raw finding count."""
    bands = rules["effort_bands"]
    size_gb = a.workload.db_size_gb or 0

    residual_hard = sum(1 for f in a.residual if f["severity"] == "hard")
    residual_medium = sum(1 for f in a.residual if f["severity"] == "medium")

    # Full cost for work the target does not resolve.
    remediation = sum(f.get("effort_days", 0) for f in a.residual)
    # Carry-over for work the target absorbs: Agent jobs still have to be moved
    # and tested on SQL MI, they just do not need re-architecting.
    carry = sum(f.get("effort_days", 0) for f in a.resolved_by_target) * 0.25
    remediation = int(round(remediation + carry))

    if a.fabric_candidate or residual_hard >= 3 or size_gb > 5120:
        band = "XL"
    elif residual_hard >= 1 or size_gb > 1024:
        band = "L"
    elif residual_medium > 2 or size_gb > 200:
        band = "M"
    else:
        band = "S"

    # Non-production is low-risk by definition and is used to prove the factory.
    # Capping the band stops a dev database being scheduled like a core ERP system.
    if a.is_non_production and band in {"L", "XL"}:
        band = "M"
        a.rationale.append(
            "Non-production workload: effort capped at Moderate and scheduled early, "
            "to prove the migration runbook before production cutover."
        )

    a.effort_band = band
    a.effort_days_low = bands[band]["days_low"]
    a.effort_days_high = bands[band]["days_high"]
    a.remediation_days = remediation


def plan_waves(assessments: list[Assessment], rules: dict) -> None:
    """Assign waves, then pull dependency-linked workloads into a common wave."""
    wave_by_band = {}
    for wave in rules["waves"]:
        for band in wave["accepts_effort"]:
            wave_by_band[band] = (wave["id"], wave["name"])

    for a in assessments:
        wid, wname = wave_by_band.get(a.effort_band, (2, "Wave 2 - Standard"))
        a.wave = wid
        a.wave_name = wname

    # Non-production goes first regardless of size: it is the proving ground.
    first = rules["waves"][0]
    for a in assessments:
        if a.is_non_production:
            a.wave = first["id"]
            a.wave_name = first["name"]

    # Linked servers couple databases together: a database cannot cut over ahead
    # of something it queries across a link. But blanket-promoting every database
    # on an instance to the latest wave destroys the plan - on a densely linked
    # instance it drags simple databases into the re-architect wave for no reason.
    # So align only genuinely tight couplings, and flag broad meshes for a proper
    # dependency-mapping exercise instead of guessing.
    MAX_TIGHT_GROUP = 6

    groups: dict[str, list[Assessment]] = {}
    for a in assessments:
        if any(f.get("dependency") for f in a.all_findings()):
            groups.setdefault(a.workload.instance, []).append(a)

    for instance, members in groups.items():
        if len(members) < 2:
            continue

        for m in members:
            m.dependency_group = instance

        if len(members) > MAX_TIGHT_GROUP:
            names = ", ".join(sorted(x.workload.database for x in members)[:4])
            for m in members:
                m.rationale.append(
                    f"Instance {instance} has a broad linked-server mesh across {len(members)} databases "
                    f"({names}...). Wave assignment here reflects individual complexity only - run a "
                    f"dependency-mapping exercise before fixing the cutover sequence."
                )
            continue

        latest = max(m.wave for m in members)
        name = next(w["name"] for w in rules["waves"] if w["id"] == latest)
        for m in members:
            if m.wave != latest:
                m.rationale.append(
                    f"Moved to {name} to cut over alongside its linked-server dependants on "
                    f"instance {instance}."
                )
            m.wave = latest
            m.wave_name = name


# --------------------------------------------------------------------------
# Orchestration
# --------------------------------------------------------------------------

def assess(workloads: list[Workload], rules: dict) -> list[Assessment]:
    results: list[Assessment] = []

    for w in workloads:
        blockers, warnings = detect_blockers(w, rules)
        a = Assessment(workload=w, blockers=blockers, warnings=warnings)
        a.is_non_production = _is_non_production(w, rules)

        # Decide, size, then verify the sizing actually fits the chosen platform.
        # A capacity failure adds a blocker and forces one re-decision, which is
        # why this is a loop rather than a straight line.
        for attempt in range(2):
            decision = choose_target(w, a.blockers, a.warnings, rules)
            a.target = decision["target"]
            a.target_label = decision["target_label"]
            a.rationale = decision["rationale"]
            a.eliminated = decision["eliminated"]
            a.fabric_candidate = decision["fabric_candidate"]
            a.confidence = decision["confidence"]

            sizing = size(w, a.target, rules, non_production=a.is_non_production)
            a.vcores = sizing["vcores"]
            a.service_tier = sizing["service_tier"]
            a.storage_gb = sizing["storage_gb"]
            a.sku_name = sizing["sku_name"]
            a.sizing_notes = sizing["sizing_notes"]
            a.sizing_confidence = sizing["sizing_confidence"]

            if not validate_capacity(a, rules) or attempt == 1:
                break

        if a.is_non_production:
            a.rationale.append(
                "Detected as non-production from its edition, name or declared workload type. "
                "Confirm before planning, and treat as the pilot for this estate."
            )

        partition_findings(a)

        eos = next((f for f in a.warnings if f["id"] == "END_OF_SUPPORT"), None)
        if eos:
            a.eos_flag = True
            m = re.search(r"\((\d{4}-\d{2}-\d{2})\)", eos["label"])
            a.eos_date = m.group(1) if m else ""

        band_effort(a, rules)
        results.append(a)

    plan_waves(results, rules)
    return results


def summarise(assessments: list[Assessment]) -> dict:
    """Roll the estate up for the executive summary sheet."""
    total = len(assessments)
    by_target: dict[str, int] = {}
    by_wave: dict[int, int] = {}
    by_band: dict[str, int] = {}
    blocker_tally: dict[str, int] = {}

    days_low = days_high = remediation = 0
    size_gb = 0.0
    cost = 0.0
    cost_ahb = 0.0
    priced = 0

    for a in assessments:
        by_target[a.target_label] = by_target.get(a.target_label, 0) + 1
        by_wave[a.wave] = by_wave.get(a.wave, 0) + 1
        by_band[a.effort_band] = by_band.get(a.effort_band, 0) + 1
        days_low += a.effort_days_low
        days_high += a.effort_days_high
        remediation += a.remediation_days
        size_gb += a.workload.db_size_gb or 0
        for f in a.all_findings():
            blocker_tally[f["label"]] = blocker_tally.get(f["label"], 0) + 1
        if a.est_monthly_usd:
            cost += a.est_monthly_usd
            priced += 1
        if a.est_monthly_ahb_usd:
            cost_ahb += a.est_monthly_ahb_usd

    return {
        "total_workloads": total,
        "total_size_gb": round(size_gb, 1),
        "by_target": by_target,
        "by_wave": dict(sorted(by_wave.items())),
        "by_effort_band": by_band,
        "paas_eligible": sum(1 for a in assessments if a.target in {"SQL_DB", "SQL_MI"}),
        "iaas_required": sum(1 for a in assessments if a.target == "SQL_VM"),
        "fabric_candidates": sum(1 for a in assessments if a.fabric_candidate),
        "non_production": sum(1 for a in assessments if a.is_non_production),
        "eos_workloads": sum(1 for a in assessments if a.eos_flag),
        "workloads_with_hard_blockers": sum(1 for a in assessments if a.blockers),
        "workloads_with_residual_work": sum(
            1 for a in assessments if any(f.get("effort_days", 0) > 0 for f in a.residual)),
        "findings_resolved_by_target": sum(len(a.resolved_by_target) for a in assessments),
        "paas_unblockable": sum(1 for a in assessments if a.target == "SQL_VM" and a.paas_unblock),
        "clean_workloads": sum(1 for a in assessments if not a.blockers and not a.warnings),
        "effort_days_low": days_low,
        "effort_days_high": days_high,
        "remediation_days": remediation,
        "top_blockers": sorted(blocker_tally.items(), key=lambda kv: -kv[1])[:12],
        "est_monthly_usd": round(cost, 2) if priced else None,
        "est_monthly_ahb_usd": round(cost_ahb, 2) if priced else None,
        "priced_workloads": priced,
        "low_confidence": sum(1 for a in assessments if a.sizing_confidence == "low"),
    }
