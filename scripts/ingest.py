"""
Ingestion layer for the Estate-to-Target Migration Assessor.

Accepts the artifacts a customer actually sends, rather than demanding a template:

  * Azure Migrate assessment export  (.xlsx - multi-sheet)
  * Data Migration Assistant report  (.json or .csv)
  * sp_Blitz / sp_BlitzFirst output  (.csv)
  * Generic instance inventory       (.csv / .xlsx / .tsv)

Everything is normalised onto the Workload record so the rules engine only ever
sees one shape. Header matching is fuzzy and case-insensitive because every
export in the wild names these columns differently.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any

import pandas as pd


# --------------------------------------------------------------------------
# Normalised record
# --------------------------------------------------------------------------

@dataclass
class Workload:
    """One migratable unit - usually a database, sometimes a whole instance."""

    source_file: str = ""
    source_format: str = ""

    instance: str = ""
    database: str = ""

    sql_version: str = ""
    sql_edition: str = ""
    os_name: str = ""

    cores: float | None = None
    cpu_util_pct: float | None = None
    memory_gb: float | None = None

    db_size_gb: float | None = None
    log_size_gb: float | None = None
    max_iops: float | None = None
    throughput_mbps: float | None = None

    compat_level: str = ""
    collation: str = ""
    workload_type: str = ""

    agent_job_count: int = 0
    linked_server_count: int = 0
    db_count: int = 0

    # Raw feature / finding strings harvested from the source artifact.
    signals: list[str] = field(default_factory=list)

    # Anything the parser saw but did not map - kept for traceability.
    extra: dict[str, Any] = field(default_factory=dict)

    @property
    def key(self) -> str:
        return f"{self.instance}/{self.database}" if self.database else self.instance

    def add_signal(self, text: Any) -> None:
        if text is None:
            return
        s = str(text).strip()
        if s and s.lower() not in {"nan", "none", "null", "n/a", "-", ""}:
            if s not in self.signals:
                self.signals.append(s)

    def to_dict(self) -> dict:
        return asdict(self)


# --------------------------------------------------------------------------
# Header matching
# --------------------------------------------------------------------------

# Ordered longest-first within each group so that specific headers win over
# generic ones (e.g. "max iops" should not be eaten by a bare "iops" rule).
COLUMN_ALIASES: dict[str, list[str]] = {
    "instance": [
        "instance name", "sql instance", "server name", "servername",
        "machine name", "host name", "hostname", "source server", "instance",
        "server", "machine",
    ],
    "database": [
        "database name", "databasename", "db name", "dbname", "database", "db",
    ],
    "sql_version": [
        "sql server version", "source sql version", "sql version", "product version",
        "engine version", "version",
    ],
    "sql_edition": [
        "sql server edition", "source edition", "sql edition", "edition",
    ],
    "os_name": [
        "operating system", "os name", "os version", "os",
    ],
    "cores": [
        "number of cores", "logical processors", "cpu cores", "core count",
        "vcpus", "cores", "cpu",
    ],
    "cpu_util_pct": [
        "cpu utilization percentage", "percentage of cores utilized",
        "cpu utilisation", "cpu utilization", "cpu usage", "cpu %", "cpu percent",
    ],
    "memory_gb": [
        "memory in use (mb)", "memory (mb)", "memory in gb", "memory (gb)",
        "ram (gb)", "allocated memory", "memory", "ram",
    ],
    "db_size_gb": [
        "database size (gb)", "database size in mb", "database size (mb)",
        "total database size (gb)", "data size (gb)", "size (gb)", "size (mb)",
        "database size", "db size", "size",
    ],
    "log_size_gb": [
        "log size (gb)", "log size (mb)", "transaction log size", "log size",
    ],
    "max_iops": [
        "total disk iops", "max iops", "iops", "disk iops", "read iops",
    ],
    "throughput_mbps": [
        "total disk throughput (mbps)", "throughput (mbps)", "disk throughput",
        "throughput",
    ],
    "compat_level": [
        "compatibility level", "compat level", "compatibility",
    ],
    "collation": [
        "database collation", "instance collation", "collation",
    ],
    "workload_type": [
        "workload type", "application type", "workload", "purpose", "environment",
    ],
    "agent_job_count": [
        "number of agent jobs", "agent job count", "sql agent jobs", "agent jobs", "jobs",
    ],
    "linked_server_count": [
        "number of linked servers", "linked server count", "linked servers",
    ],
    "db_count": [
        "number of databases", "database count", "db count", "databases",
    ],
}

# Headers whose *values* are feature findings rather than measurements.
SIGNAL_HEADER_HINTS = [
    "feature", "finding", "blocker", "issue", "compatibility", "recommendation",
    "warning", "breaking change", "behavior change", "behaviour change",
    "readiness", "migration issue", "assessment", "impact", "rule", "category",
    "title", "description", "object type", "severity", "priority",
]


def _norm(text: Any) -> str:
    """Lower-case, collapse punctuation, squeeze whitespace."""
    return re.sub(r"[\s_\-]+", " ", str(text).strip().lower()).strip()


def _match_column(header: str) -> str | None:
    """Map a raw header onto a Workload field, or None."""
    h = _norm(header)
    if not h:
        return None
    # Exact alias hit first.
    for field_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if h == alias:
                return field_name
    # Then containment, longest alias first so specific beats generic.
    best: tuple[int, str] | None = None
    for field_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in h or h in alias:
                score = len(alias)
                if best is None or score > best[0]:
                    best = (score, field_name)
    return best[1] if best else None


def _is_signal_header(header: str) -> bool:
    h = _norm(header)
    return any(hint in h for hint in SIGNAL_HEADER_HINTS)


# --------------------------------------------------------------------------
# Value coercion
# --------------------------------------------------------------------------

def _num(value: Any) -> float | None:
    """Pull a float out of messy cell content. Returns None when not numeric."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        f = float(value)
        return None if f != f else f  # drop NaN
    s = str(value).strip().replace(",", "")
    if not s or s.lower() in {"nan", "none", "null", "n/a", "-", "unknown", "not available"}:
        return None
    m = re.search(r"-?\d+(?:\.\d+)?", s)
    return float(m.group()) if m else None


def _to_gb(value: Any, header: str) -> float | None:
    """Normalise a size to GB, honouring a unit hint in the header."""
    n = _num(value)
    if n is None:
        return None
    h = _norm(header)
    if "(mb)" in h or " mb" in h or h.endswith("mb"):
        return round(n / 1024.0, 3)
    if "(tb)" in h or " tb" in h or h.endswith("tb"):
        return round(n * 1024.0, 3)
    if "(kb)" in h or h.endswith("kb"):
        return round(n / (1024.0 * 1024.0), 4)
    return n


def _clean_version(value: Any) -> str:
    """Reduce any version spelling to a canonical release year token."""
    s = str(value or "").strip()
    if not s:
        return ""
    low = s.lower()
    for token in ["2008 r2", "2008r2"]:
        if token in low:
            return "2008R2"
    for year in ["2025", "2022", "2019", "2017", "2016", "2014", "2012", "2008", "2005", "2000"]:
        if year in s:
            return year
    # Fall back on the major build number.
    m = re.match(r"^(\d+)\.", s)
    if m:
        return {
            "16": "2022", "15": "2019", "14": "2017", "13": "2016",
            "12": "2014", "11": "2012", "10": "2008", "9": "2005",
        }.get(m.group(1), s)
    return s


# --------------------------------------------------------------------------
# Generic tabular parser - the workhorse
# --------------------------------------------------------------------------

def _parse_frame(df: pd.DataFrame, source_file: str, source_format: str,
                 sheet: str = "") -> list[Workload]:
    """Turn any instance/database-shaped table into Workload records."""
    if df.empty:
        return []

    df = df.dropna(how="all").dropna(axis=1, how="all")
    if df.empty:
        return []

    mapping: dict[str, str] = {}      # raw header -> Workload field
    signal_cols: list[str] = []
    for col in df.columns:
        matched = _match_column(col)
        if matched and matched not in mapping.values():
            mapping[col] = matched
        elif _is_signal_header(col):
            signal_cols.append(col)

    # A table with no identity column is not an inventory - skip it.
    if not any(v in {"instance", "database"} for v in mapping.values()):
        return []

    out: list[Workload] = []
    for _, row in df.iterrows():
        w = Workload(source_file=source_file, source_format=source_format)
        if sheet:
            w.extra["sheet"] = sheet

        for raw_col, fieldname in mapping.items():
            val = row.get(raw_col)
            if val is None or (isinstance(val, float) and val != val):
                continue

            if fieldname in {"db_size_gb", "log_size_gb", "memory_gb"}:
                setattr(w, fieldname, _to_gb(val, raw_col))
            elif fieldname in {"cores", "cpu_util_pct", "max_iops", "throughput_mbps"}:
                setattr(w, fieldname, _num(val))
            elif fieldname in {"agent_job_count", "linked_server_count", "db_count"}:
                n = _num(val)
                setattr(w, fieldname, int(n) if n is not None else 0)
            elif fieldname == "sql_version":
                w.sql_version = _clean_version(val)
            else:
                setattr(w, fieldname, str(val).strip())

        for col in signal_cols:
            val = row.get(col)
            if val is None or (isinstance(val, float) and val != val):
                continue
            text = str(val).strip()
            if text and text.lower() not in {"nan", "no", "false", "0", "none", "ready", "n/a"}:
                # Keep the header as context - "CLR: True" is clearer than "True".
                if text.lower() in {"yes", "true", "1"}:
                    w.add_signal(col)
                else:
                    w.add_signal(f"{col}: {text}" if _is_signal_header(col) else text)

        if w.instance or w.database:
            out.append(w)

    return out


# --------------------------------------------------------------------------
# Format-specific entry points
# --------------------------------------------------------------------------

def parse_excel(path: Path) -> list[Workload]:
    """Azure Migrate exports and hand-built inventories both land here."""
    workloads: list[Workload] = []
    try:
        book = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 - surfaced to the caller as a warning
        raise RuntimeError(f"Could not read workbook {path.name}: {exc}") from exc

    is_azure_migrate = any(
        kw in " ".join(book.keys()).lower()
        for kw in ["assessed", "azure migrate", "readiness", "sql instance"]
    )
    fmt = "azure_migrate" if is_azure_migrate else "inventory_xlsx"

    for sheet_name, df in book.items():
        # Azure Migrate puts a banner above the real header on some sheets.
        df = _reheader_if_needed(df)
        workloads.extend(_parse_frame(df, path.name, fmt, sheet_name))

    return _merge_duplicates(workloads)


def _reheader_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """Azure Migrate sheets often carry 1-3 banner rows before the real header."""
    if df.empty:
        return df
    unnamed = sum(1 for c in df.columns if str(c).startswith("Unnamed"))
    if unnamed < max(2, len(df.columns) // 2):
        return df

    for idx in range(min(6, len(df))):
        candidate = df.iloc[idx]
        named = sum(
            1 for v in candidate
            if v is not None and str(v).strip() and not str(v).startswith("Unnamed")
            and not (isinstance(v, float) and v != v)
        )
        if named >= max(2, len(df.columns) // 2):
            new = df.iloc[idx + 1:].copy()
            new.columns = [str(c).strip() for c in candidate]
            return new.reset_index(drop=True)
    return df


def parse_csv(path: Path) -> list[Workload]:
    """sp_Blitz output, DMA CSV exports and plain inventories."""
    df = None
    for enc in ("utf-8-sig", "utf-16", "latin-1"):
        for sep in (None, ",", ";", "\t"):
            try:
                df = pd.read_csv(path, encoding=enc, sep=sep, engine="python")
                if df.shape[1] > 1:
                    break
            except Exception:  # noqa: BLE001 - keep trying other dialects
                df = None
        if df is not None and df.shape[1] > 1:
            break
    if df is None:
        raise RuntimeError(f"Could not parse {path.name} as delimited text")

    headers = " ".join(str(c).lower() for c in df.columns)
    if "findingsgroup" in headers.replace(" ", "") or "priority" in headers and "finding" in headers:
        return _parse_spblitz(df, path.name)
    return _merge_duplicates(_parse_frame(df, path.name, "inventory_csv"))


def _parse_spblitz(df: pd.DataFrame, source_file: str) -> list[Workload]:
    """sp_Blitz is finding-per-row, so collapse it to one Workload per database."""
    cols = {_norm(c): c for c in df.columns}

    def col(*names: str) -> str | None:
        for n in names:
            if n in cols:
                return cols[n]
        for n in names:
            for k, v in cols.items():
                if n in k:
                    return v
        return None

    db_col = col("databasename", "database name", "database")
    finding_col = col("finding")
    group_col = col("findingsgroup", "findings group")
    detail_col = col("details", "detail")
    prio_col = col("priority")
    server_col = col("servername", "server name", "server")

    by_db: dict[str, Workload] = {}
    for _, row in df.iterrows():
        dbname = str(row.get(db_col, "") or "").strip() if db_col else ""
        server = str(row.get(server_col, "") or "").strip() if server_col else ""
        if not dbname and not server:
            continue
        key = f"{server}/{dbname}"
        w = by_db.get(key)
        if w is None:
            w = Workload(
                source_file=source_file,
                source_format="sp_blitz",
                instance=server or "UNKNOWN-INSTANCE",
                database=dbname,
            )
            by_db[key] = w

        parts = [
            str(row.get(c, "")).strip()
            for c in (group_col, finding_col, detail_col)
            if c and row.get(c) is not None
        ]
        text = " | ".join(p for p in parts if p and p.lower() != "nan")
        if text:
            prio = _num(row.get(prio_col)) if prio_col else None
            w.add_signal(f"[P{int(prio)}] {text}" if prio is not None else text)

    return list(by_db.values())


def parse_dma_json(path: Path) -> list[Workload]:
    """Data Migration Assistant JSON. Schema varies by DMA version, so walk it."""
    raw = json.loads(path.read_text(encoding="utf-8-sig"))
    workloads: dict[str, Workload] = {}

    def ensure(instance: str, database: str) -> Workload:
        key = f"{instance}/{database}"
        if key not in workloads:
            workloads[key] = Workload(
                source_file=path.name,
                source_format="dma",
                instance=instance or "UNKNOWN-INSTANCE",
                database=database,
            )
        return workloads[key]

    def walk(node: Any, instance: str = "", database: str = "") -> None:
        if isinstance(node, dict):
            lowered = {_norm(k): k for k in node}

            for hint in ("server name", "servername", "instance name", "source server", "server"):
                if hint in lowered and isinstance(node[lowered[hint]], str):
                    instance = node[lowered[hint]].strip() or instance
                    break
            for hint in ("database name", "databasename", "name", "database"):
                if hint in lowered and isinstance(node[lowered[hint]], str):
                    cand = node[lowered[hint]].strip()
                    if cand and len(cand) < 130:
                        database = cand
                    break

            # A compatibility / feature-parity finding.
            title = None
            for hint in ("title", "rulename", "rule name", "displayname", "changecategory"):
                if hint in lowered and isinstance(node[lowered[hint]], str):
                    title = node[lowered[hint]].strip()
                    break
            if title:
                sev = ""
                for hint in ("severity", "impact", "changecategory", "category"):
                    if hint in lowered and isinstance(node[lowered[hint]], str):
                        sev = node[lowered[hint]].strip()
                        break
                impacted = ""
                for hint in ("impactedobjects", "impacted objects", "objectcount"):
                    if hint in lowered:
                        val = node[lowered[hint]]
                        if isinstance(val, list):
                            impacted = f" ({len(val)} objects)"
                        elif isinstance(val, (int, float)):
                            impacted = f" ({int(val)} objects)"
                        break
                w = ensure(instance, database)
                w.add_signal(f"{sev + ': ' if sev else ''}{title}{impacted}")

            for key, val in node.items():
                nk = _norm(key)
                if nk in {"compatibilitylevel", "compatibility level"} and isinstance(val, (str, int)):
                    ensure(instance, database).compat_level = str(val)
                elif nk == "collation" and isinstance(val, str):
                    ensure(instance, database).collation = val
                elif nk in {"sqlserverversion", "sql server version", "sourceversion", "version"} \
                        and isinstance(val, (str, int, float)):
                    ensure(instance, database).sql_version = _clean_version(val)
                elif nk in {"edition", "sourceedition"} and isinstance(val, str):
                    ensure(instance, database).sql_edition = val
                elif nk in {"sizemb", "size mb", "databasesizemb"}:
                    n = _num(val)
                    if n is not None:
                        ensure(instance, database).db_size_gb = round(n / 1024.0, 3)
                elif nk in {"sizegb", "size gb", "databasesizegb"}:
                    n = _num(val)
                    if n is not None:
                        ensure(instance, database).db_size_gb = n
                walk(val, instance, database)

        elif isinstance(node, list):
            for item in node:
                walk(item, instance, database)

    walk(raw)
    return [w for w in workloads.values() if w.database or w.signals]


# --------------------------------------------------------------------------
# Post-processing
# --------------------------------------------------------------------------

def _merge_duplicates(workloads: list[Workload]) -> list[Workload]:
    """Azure Migrate spreads one database across several sheets - stitch them."""
    merged: dict[str, Workload] = {}
    for w in workloads:
        key = w.key.lower()
        if key not in merged:
            merged[key] = w
            continue
        base = merged[key]
        for f in (
            "sql_version", "sql_edition", "os_name", "compat_level",
            "collation", "workload_type",
        ):
            if not getattr(base, f) and getattr(w, f):
                setattr(base, f, getattr(w, f))
        for f in (
            "cores", "cpu_util_pct", "memory_gb", "db_size_gb",
            "log_size_gb", "max_iops", "throughput_mbps",
        ):
            bv, wv = getattr(base, f), getattr(w, f)
            if wv is not None and (bv is None or wv > bv):
                setattr(base, f, wv)
        for f in ("agent_job_count", "linked_server_count", "db_count"):
            setattr(base, f, max(getattr(base, f), getattr(w, f)))
        for s in w.signals:
            base.add_signal(s)
    return list(merged.values())


def _fold_instance_profiles(workloads: list[Workload]) -> list[Workload]:
    """Instance-level rows describe a host, not a migratable database.

    Azure Migrate exports an instance sheet and a database sheet. The instance
    rows carry the good telemetry (cores, CPU utilisation, memory, OS) but no
    database name. Left alone they become phantom workloads that get sized and
    priced as though they were databases, inflating both the count and the
    run-rate. So fold their telemetry down into the real databases on that
    instance, and only keep a profile standalone when no databases were found.
    """
    profiles: dict[str, Workload] = {}
    databases: list[Workload] = []

    for w in workloads:
        if w.database:
            databases.append(w)
        else:
            key = w.instance.lower()
            if key in profiles:
                existing = profiles[key]
                for f in ("cores", "cpu_util_pct", "memory_gb"):
                    if getattr(existing, f) is None and getattr(w, f) is not None:
                        setattr(existing, f, getattr(w, f))
                for s in w.signals:
                    existing.add_signal(s)
            else:
                profiles[key] = w

    by_instance: dict[str, list[Workload]] = {}
    for w in databases:
        by_instance.setdefault(w.instance.lower(), []).append(w)

    for key, profile in profiles.items():
        members = by_instance.get(key)
        if not members:
            # Nothing to fold into - the instance itself is the workload.
            profile.database = "(instance-level)"
            databases.append(profile)
            continue

        for w in members:
            # Host telemetry describes the instance, so it is shared across its
            # databases. Only fill gaps; never overwrite database-specific data.
            for f in ("cores", "cpu_util_pct", "memory_gb", "sql_version",
                      "sql_edition", "os_name", "collation"):
                if not getattr(w, f) and getattr(profile, f):
                    setattr(w, f, getattr(profile, f))
            for f in ("agent_job_count", "linked_server_count", "db_count"):
                if getattr(w, f) == 0 and getattr(profile, f) > 0:
                    setattr(w, f, getattr(profile, f))
            for s in profile.signals:
                w.add_signal(s)
            w.extra["instance_profile_merged"] = True

    return databases


def load(paths: list[Path]) -> tuple[list[Workload], list[str]]:
    """Ingest every supplied artifact. Returns (workloads, warnings)."""
    workloads: list[Workload] = []
    warnings: list[str] = []

    for path in paths:
        if not path.exists():
            warnings.append(f"Not found, skipped: {path}")
            continue
        suffix = path.suffix.lower()
        try:
            if suffix in {".xlsx", ".xlsm", ".xls"}:
                found = parse_excel(path)
            elif suffix == ".json":
                found = parse_dma_json(path)
            elif suffix in {".csv", ".tsv", ".txt"}:
                found = parse_csv(path)
            else:
                warnings.append(f"Unsupported extension {suffix}, skipped: {path.name}")
                continue
        except Exception as exc:  # noqa: BLE001 - one bad file must not kill the run
            warnings.append(f"Failed to parse {path.name}: {exc}")
            continue

        if not found:
            warnings.append(f"No recognisable workloads in {path.name}")
        workloads.extend(found)

    # Cross-file merge: a DMA report and an Azure Migrate export describing the
    # same database should become a single, richer record.
    workloads = _merge_duplicates(workloads)
    workloads = _fold_instance_profiles(workloads)

    for w in workloads:
        if not w.instance:
            w.instance = "UNKNOWN-INSTANCE"

    return workloads, warnings
