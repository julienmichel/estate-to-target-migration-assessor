#!/usr/bin/env python3
"""
Generates realistic sample artifacts in all four supported input formats.

These exist so the skill can be demonstrated and regression-tested without a real
customer estate. The data is entirely fictional and deliberately messy - different
column spellings, mixed units, missing telemetry - because that is what real
exports look like.

    python make_samples.py
"""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parent.parent / "samples"


def generic_inventory() -> None:
    """A hand-built inventory spreadsheet - the most common real-world input."""
    rows = [
        # instance, db, version, edition, cores, cpu%, mem GB, size GB, log GB, iops, mbps, collation, jobs, links, type
        ("SQLPRD01", "ERP_Core",        "2014", "Enterprise", 24, 68, 256, 1850, 320, 11200, 420, "SQL_Latin1_General_CP1_CI_AS", 42, 3, "OLTP"),
        ("SQLPRD01", "ERP_Archive",     "2014", "Enterprise", 24, 12, 256, 6200, 180,  1400,  90, "SQL_Latin1_General_CP1_CI_AS",  4, 0, "Archive"),
        ("SQLPRD02", "CRM_Live",        "2016", "Enterprise", 16, 55, 128,  940, 160,  7800, 310, "SQL_Latin1_General_CP1_CI_AS", 18, 1, "OLTP"),
        ("SQLPRD03", "Billing",         "2019", "Standard",    8, 41,  64,  420,  70,  3100, 140, "SQL_Latin1_General_CP1_CI_AS",  6, 0, "OLTP"),
        ("SQLPRD03", "BillingStaging",  "2019", "Standard",    8, 18,  64,   95,  20,   800,  40, "SQL_Latin1_General_CP1_CI_AS",  2, 0, "Staging"),
        ("SQLPRD04", "HR_Portal",       "2019", "Standard",    4, 22,  32,   38,   8,   450,  25, "SQL_Latin1_General_CP1_CI_AS",  0, 0, "OLTP"),
        ("SQLPRD04", "Timesheets",      "2019", "Standard",    4, 15,  32,   12,   3,   210,  12, "SQL_Latin1_General_CP1_CI_AS",  0, 0, "OLTP"),
        ("SQLDWH01", "EnterpriseDW",    "2016", "Enterprise", 32, 72, 512, 14500, 900, 18000, 780, "SQL_Latin1_General_CP1_CI_AS", 31, 4, "Reporting"),
        ("SQLDWH01", "SalesMart",       "2016", "Enterprise", 32, 35, 512,  2400, 260,  5200, 240, "SQL_Latin1_General_CP1_CI_AS", 12, 2, "Analytics"),
        ("SQLLEG01", "LegacyFinance",   "2008R2", "Enterprise", 8, 30,  48,  610, 120,  2200, 110, "Latin1_General_BIN",            9, 2, "OLTP"),
        ("SQLLEG01", "DocumentStore",   "2008R2", "Enterprise", 8, 25,  48, 3100, 140,  2900, 180, "Latin1_General_BIN",            3, 0, "Content"),
        ("SQLAPP05", "WebCatalogue",    "2022", "Standard",    8, 34,  64,  180,  30,  1900,  95, "SQL_Latin1_General_CP1_CI_AS",  0, 0, "OLTP"),
        ("SQLAPP05", "SessionState",    "2022", "Standard",    8, 61,  64,   22,  40,  6400, 260, "SQL_Latin1_General_CP1_CI_AS",  0, 0, "OLTP"),
        ("SQLDEV01", "DevSandbox",      "2019", "Developer",   4,  8,  16,   45,  10,   300,  18, "SQL_Latin1_General_CP1_CI_AS",  1, 0, "Dev"),
        ("SQLINT01", "IntegrationHub",  "2017", "Enterprise", 12, 48,  96,  260,  55,  4100, 190, "SQL_Latin1_General_CP1_CI_AS", 27, 5, "Integration"),
    ]

    # Deliberately varied header spellings - this is what the fuzzy matcher is for.
    df = pd.DataFrame(rows, columns=[
        "Server Name", "Database Name", "SQL Server Version", "Edition",
        "Number of Cores", "CPU Utilization Percentage", "Memory (GB)",
        "Database Size (GB)", "Log Size (GB)", "Max IOPS", "Throughput (MBps)",
        "Collation", "Number of Agent Jobs", "Linked Servers", "Workload Type",
    ])

    # Feature flags as a findings column.
    features = [
        "Service Broker enabled; CLR assemblies present (UNSAFE)",
        "",
        "SQL Agent jobs; cross-database queries to ERP_Core",
        "Cross database query detected",
        "",
        "",
        "",
        "SSIS packages hosted on instance; SSAS cube co-located; linked server to Oracle",
        "Cross database queries; SSRS reporting services",
        "FileStream enabled; deprecated feature: SETUSER",
        "FileStream enabled; FileTable in use",
        "",
        "",
        "",
        "xp_cmdshell enabled; linked servers to 5 external systems; Database Mail configured",
    ]
    df["Migration Issue"] = features

    df.to_csv(OUT / "sample_inventory.csv", index=False, encoding="utf-8")
    print(f"  + sample_inventory.csv ({len(df)} rows)")


def azure_migrate_export() -> None:
    """Mimics an Azure Migrate assessment workbook - multi-sheet, banner rows, MB units."""
    instances = pd.DataFrame([
        ("AZMIG-SQL-01", "SQLPRD01", "Microsoft SQL Server 2014", "Enterprise",
         "Windows Server 2012 R2", 24, 68.4, 262144, 8, "Ready", "Azure SQL Managed Instance"),
        ("AZMIG-SQL-02", "SQLPRD02", "Microsoft SQL Server 2016", "Enterprise",
         "Windows Server 2016", 16, 55.1, 131072, 3, "Ready with conditions", "Azure SQL Managed Instance"),
        ("AZMIG-SQL-03", "SQLAPP05", "Microsoft SQL Server 2022", "Standard",
         "Windows Server 2022", 8, 34.2, 65536, 2, "Ready", "Azure SQL Database"),
    ], columns=[
        "Assessment Name", "SQL instance name", "SQL Server version", "SQL Server edition",
        "Operating system", "Number of cores", "Percentage of cores utilized",
        "Memory (MB)", "Number of databases", "Azure SQL MI readiness", "Recommended target",
    ])

    databases = pd.DataFrame([
        ("SQLPRD01", "ERP_Core",       1894400, 327680, "Ready with conditions", 11200, 420, 110,
         "SQL_Latin1_General_CP1_CI_AS", "Service Broker is not supported on Azure SQL Database"),
        ("SQLPRD01", "ERP_Archive",    6348800, 184320, "Ready", 1400, 90, 110,
         "SQL_Latin1_General_CP1_CI_AS", ""),
        ("SQLPRD02", "CRM_Live",        962560, 163840, "Ready with conditions", 7800, 310, 130,
         "SQL_Latin1_General_CP1_CI_AS", "Cross-database queries are not supported on Azure SQL Database"),
        ("SQLAPP05", "WebCatalogue",    184320,  30720, "Ready", 1900, 95, 160,
         "SQL_Latin1_General_CP1_CI_AS", ""),
        ("SQLAPP05", "SessionState",     22528,  40960, "Ready", 6400, 260, 160,
         "SQL_Latin1_General_CP1_CI_AS", ""),
    ], columns=[
        "SQL instance name", "Database name", "Database size (MB)", "Log size (MB)",
        "Azure SQL Database readiness", "Total disk IOPS", "Total disk throughput (MBps)",
        "Compatibility level", "Database collation", "Migration issue",
    ])

    path = OUT / "sample_azure_migrate.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        instances.to_excel(writer, sheet_name="Assessed SQL instances", index=False)
        databases.to_excel(writer, sheet_name="Assessed SQL databases", index=False)
    print(f"  + sample_azure_migrate.xlsx ({len(instances)} instances, {len(databases)} databases)")


def dma_report() -> None:
    """Data Migration Assistant JSON - nested, finding-oriented."""
    payload = {
        "AssessmentName": "Contoso SQL Readiness",
        "AssessmentTarget": "Azure SQL Database",
        "Servers": [
            {
                "ServerName": "SQLLEG01",
                "SqlServerVersion": "10.50.6000",
                "Edition": "Enterprise Edition",
                "Databases": [
                    {
                        "Name": "LegacyFinance",
                        "CompatibilityLevel": "100",
                        "Collation": "Latin1_General_BIN",
                        "SizeMB": 624640,
                        "CompatibilityIssues": [
                            {"Title": "Deprecated syntax SETUSER is discontinued",
                             "Severity": "Breaking change",
                             "ChangeCategory": "BreakingChange",
                             "ImpactedObjects": [{"Name": "usp_LegacyImpersonate"},
                                                 {"Name": "usp_BatchPost"}]},
                            {"Title": "Linked servers are not supported on the target platform",
                             "Severity": "Blocking",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "ORACLE_FIN"}, {"Name": "LEGACY_AS400"}]},
                            {"Title": "SQL Server Agent jobs will not migrate",
                             "Severity": "Blocking",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "nightly_close"}]},
                        ],
                    },
                    {
                        "Name": "DocumentStore",
                        "CompatibilityLevel": "100",
                        "Collation": "Latin1_General_BIN",
                        "SizeMB": 3174400,
                        "CompatibilityIssues": [
                            {"Title": "FILESTREAM is not supported on Azure SQL Database or Managed Instance",
                             "Severity": "Blocking",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "DocumentBlobs"}]},
                            {"Title": "FileTable is not supported",
                             "Severity": "Blocking",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "ScannedDocs"}]},
                        ],
                    },
                ],
            },
            {
                "ServerName": "SQLINT01",
                "SqlServerVersion": "14.0.3456",
                "Edition": "Enterprise Edition",
                "Databases": [
                    {
                        "Name": "IntegrationHub",
                        "CompatibilityLevel": "140",
                        "Collation": "SQL_Latin1_General_CP1_CI_AS",
                        "SizeMB": 266240,
                        "CompatibilityIssues": [
                            {"Title": "xp_cmdshell is not available on Azure SQL Database or Managed Instance",
                             "Severity": "Blocking",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "usp_RunExtract"}]},
                            {"Title": "Linked servers reference non-SQL providers",
                             "Severity": "Warning",
                             "ChangeCategory": "MigrationBlocker",
                             "ImpactedObjects": [{"Name": "SAP_PROD"}, {"Name": "ODBC_WMS"}]},
                            {"Title": "Database Mail profile configured",
                             "Severity": "Information",
                             "ChangeCategory": "Informational",
                             "ImpactedObjects": [{"Name": "AlertProfile"}]},
                        ],
                    },
                ],
            },
        ],
    }
    (OUT / "sample_dma_report.json").write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print("  + sample_dma_report.json (2 servers, 3 databases)")


def spblitz_output() -> None:
    """sp_Blitz is one row per finding - the parser collapses it per database."""
    rows = [
        ("SQLPRD01", "ERP_Core", 10, "Reliability", "CLR is Enabled",
         "CLR assemblies are present and at least one is marked UNSAFE."),
        ("SQLPRD01", "ERP_Core", 50, "Performance", "Service Broker Enabled",
         "Service Broker queues are active on this database."),
        ("SQLPRD01", "ERP_Core", 100, "Informational", "Agent Jobs Owned By User",
         "42 SQL Server Agent jobs reference this database."),
        ("SQLPRD02", "CRM_Live", 50, "Reliability", "Cross Database Ownership Chaining",
         "Cross database queries detected referencing ERP_Core."),
        ("SQLPRD02", "CRM_Live", 110, "Performance", "Deprecated Feature In Use",
         "Deprecated feature usage recorded in the last 7 days."),
        ("SQLDWH01", "EnterpriseDW", 50, "Reliability", "Linked Server Configured",
         "Linked server ORACLE_FIN uses a non-SQL provider."),
        ("SQLDWH01", "EnterpriseDW", 100, "Informational", "SSAS Installed On Server",
         "Analysis Services is co-located with the database engine."),
        ("SQLDWH01", "EnterpriseDW", 100, "Informational", "SSIS Installed On Server",
         "Integration Services catalogue present on this instance."),
        ("SQLDWH01", "SalesMart", 50, "Reliability", "Cross Database Query",
         "Three-part naming references EnterpriseDW."),
        ("SQLAPP05", "SessionState", 200, "Informational", "No Issues Found",
         "No notable findings for this database."),
    ]
    df = pd.DataFrame(rows, columns=[
        "ServerName", "DatabaseName", "Priority", "FindingsGroup", "Finding", "Details",
    ])
    df.to_csv(OUT / "sample_spblitz.csv", index=False, encoding="utf-8")
    print(f"  + sample_spblitz.csv ({len(df)} findings)")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    print("Generating sample artifacts...")
    generic_inventory()
    azure_migrate_export()
    dma_report()
    spblitz_output()
    print(f"\nWritten to {OUT}")
