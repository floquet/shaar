#!/usr/bin/python3
"""
Interactive Round-Trip Demo for XL Loyalists:
1. Reads normalized source tabs (Requirements, Parts, Req_Part_Map, Failure_Modes)
   from an existing .xlsx workbook (or injects a sample Tier-4 sub-part / MTBF edit).
2. Validates referential integrity in SQLite (foreign keys ON).
3. Re-runs analytical & recursive CTE queries and writes a versioned -01.xlsx workbook
   (plus a Diff_Audit tab showing exact metric deltas from baseline).
"""

import argparse
import sqlite3
from pathlib import Path
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


SCHEMA_SQL = """
PRAGMA foreign_keys = ON;

DROP TABLE IF EXISTS Failure_Modes;
DROP TABLE IF EXISTS Req_Part_Map;
DROP TABLE IF EXISTS Parts;
DROP TABLE IF EXISTS Requirements;

CREATE TABLE Requirements (
    Req_ID TEXT PRIMARY KEY,
    Title TEXT NOT NULL,
    Threshold_Metric TEXT NOT NULL
);

CREATE TABLE Parts (
    Part_ID TEXT PRIMARY KEY,
    Part_Name TEXT NOT NULL,
    Subsystem TEXT NOT NULL,
    MTBF_Hours REAL NOT NULL CHECK(MTBF_Hours > 0),
    Parent_Part_ID TEXT REFERENCES Parts(Part_ID)
);

CREATE TABLE Req_Part_Map (
    Req_ID TEXT REFERENCES Requirements(Req_ID),
    Part_ID TEXT REFERENCES Parts(Part_ID),
    PRIMARY KEY (Req_ID, Part_ID)
);

CREATE TABLE Failure_Modes (
    Failure_ID TEXT PRIMARY KEY,
    Part_ID TEXT REFERENCES Parts(Part_ID),
    Mode_Name TEXT NOT NULL,
    Severity TEXT CHECK(Severity IN ('Critical', 'Major', 'Minor')) NOT NULL
);
"""

SQL_CRITICAL_ROLLUP = """
SELECT
    r.Req_ID,
    r.Title AS Requirement,
    COUNT(DISTINCT p.Part_ID) AS High_Risk_Parts,
    ROUND(SUM(1.0 / p.MTBF_Hours) * 1e6, 2) AS Combined_FITS,
    ROUND(1.0 / SUM(1.0 / p.MTBF_Hours), 1) AS Effective_Series_MTBF_Hrs
FROM Requirements r
JOIN Req_Part_Map rpm ON r.Req_ID = rpm.Req_ID
JOIN Parts p ON rpm.Part_ID = p.Part_ID
JOIN Failure_Modes fm ON p.Part_ID = fm.Part_ID
WHERE p.MTBF_Hours < 10000
  AND fm.Severity = 'Critical'
GROUP BY r.Req_ID, r.Title
ORDER BY Combined_FITS DESC;
"""


def build_blast_radius_sql(target_part: str) -> str:
    return f"""
    WITH RECURSIVE Assembly_Chain(Part_ID, Part_Name, Parent_Part_ID, Depth, Path) AS (
        SELECT Part_ID, Part_Name, Parent_Part_ID, 0, Part_ID || ' (' || Part_Name || ')'
        FROM Parts
        WHERE Part_ID = '{target_part}'
        UNION ALL
        SELECT p.Part_ID, p.Part_Name, p.Parent_Part_ID, ac.Depth + 1,
               ac.Path || ' -> ' || p.Part_ID || ' (' || p.Part_Name || ')'
        FROM Parts p
        JOIN Assembly_Chain ac ON p.Part_ID = ac.Parent_Part_ID
    )
    SELECT
        ac.Depth AS Assembly_Tier,
        ac.Part_ID AS Impacted_Part_ID,
        ac.Part_Name AS Impacted_Assembly,
        r.Req_ID AS Threatened_Req_ID,
        r.Title AS Threatened_Requirement,
        ac.Path AS Propagation_Chain
    FROM Assembly_Chain ac
    LEFT JOIN Req_Part_Map rpm ON ac.Part_ID = rpm.Part_ID
    LEFT JOIN Requirements r ON rpm.Req_ID = r.Req_ID
    ORDER BY ac.Depth, r.Req_ID;
    """


def load_from_excel(xlsx_path: Path) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(xlsx_path)
    required_tabs = ["Requirements", "Parts", "Req_Part_Map", "Failure_Modes"]
    missing = [t for t in required_tabs if t not in xls.sheet_names]
    if missing:
        raise ValueError(f"Missing required source tabs in {xlsx_path}: {missing}")
    return {tab: xls.parse(tab) for tab in required_tabs}


def ingest_and_query(
    source_tables: dict[str, pd.DataFrame],
    db_path: Path,
    target_part: str,
    baseline_risk: pd.DataFrame | None = None,
) -> dict[str, pd.DataFrame]:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.executescript(SCHEMA_SQL)

    parts_df = source_tables["Parts"].copy()
    parts_df["Parent_Part_ID"] = parts_df["Parent_Part_ID"].where(
        pd.notna(parts_df["Parent_Part_ID"]), None
    )

    inserted = set()
    remaining = parts_df.to_dict("records")
    ordered_parts = []
    while remaining:
        progressed = False
        next_remaining = []
        for row in remaining:
            parent = row["Parent_Part_ID"]
            if parent is None or parent in inserted:
                ordered_parts.append(row)
                inserted.add(row["Part_ID"])
                progressed = True
            else:
                next_remaining.append(row)
        if not progressed:
            bad = [r["Part_ID"] for r in next_remaining]
            raise ValueError(f"Orphaned or cyclic Parent_Part_ID detected in Parts: {bad}")
        remaining = next_remaining

    source_tables["Requirements"].to_sql("Requirements", conn, if_exists="append", index=False)
    pd.DataFrame(ordered_parts).to_sql("Parts", conn, if_exists="append", index=False)
    source_tables["Req_Part_Map"].to_sql("Req_Part_Map", conn, if_exists="append", index=False)
    source_tables["Failure_Modes"].to_sql("Failure_Modes", conn, if_exists="append", index=False)
    conn.commit()

    new_risk = pd.read_sql_query(SQL_CRITICAL_ROLLUP, conn)
    blast_radius = pd.read_sql_query(build_blast_radius_sql(target_part), conn)

    out_tables: dict[str, pd.DataFrame] = {
        "1_Critical_Risk_Query": new_risk,
        "2_Blast_Radius_CTE": blast_radius,
    }

    if baseline_risk is not None:
        diff = pd.merge(
            baseline_risk[["Req_ID", "Requirement", "Combined_FITS", "Effective_Series_MTBF_Hrs"]],
            new_risk[["Req_ID", "Requirement", "Combined_FITS", "Effective_Series_MTBF_Hrs"]],
            on=["Req_ID", "Requirement"],
            how="outer",
            suffixes=("_Baseline", "_Updated"),
        )
        diff["Delta_FITS"] = (diff["Combined_FITS_Updated"] - diff["Combined_FITS_Baseline"]).round(2)
        diff["Delta_MTBF_Hrs"] = (
            diff["Effective_Series_MTBF_Hrs_Updated"] - diff["Effective_Series_MTBF_Hrs_Baseline"]
        ).round(1)
        out_tables["0_Diff_Audit"] = diff

    for tab in ["Requirements", "Parts", "Req_Part_Map", "Failure_Modes"]:
        out_tables[tab] = pd.read_sql_query(f"SELECT * FROM {tab};", conn)

    conn.close()
    return out_tables


def export_styled_excel(tables: dict[str, pd.DataFrame], xlsx_path: Path) -> None:
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    query_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    diff_header_fill = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    zebra_fill = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=11)
    body_font = Font(color="111111", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for sheet_name, df in tables.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            if sheet_name.startswith("0_"):
                fill_to_use = diff_header_fill
            elif sheet_name[0].isdigit():
                fill_to_use = query_header_fill
            else:
                fill_to_use = header_fill

            for col_idx, col_name in enumerate(df.columns, start=1):
                cell = ws.cell(row=1, column=col_idx)
                cell.fill = fill_to_use
                cell.font = header_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
                cell.border = thin_border

                max_len = max(
                    len(str(col_name)),
                    *(len(str(val)) for val in df[col_name].fillna("")),
                )
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 68)

            for row_idx in range(2, len(df) + 2):
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = body_font
                    cell.border = thin_border
                    if row_idx % 2 == 0:
                        cell.fill = zebra_fill


def main() -> None:
    parser = argparse.ArgumentParser(description="Round-trip Excel -> SQLite -> Versioned Excel")
    parser.add_argument("--input", default="radar_ram_demo.xlsx", help="Baseline .xlsx file")
    parser.add_argument("--output", default="radar_ram_demo-01.xlsx", help="Output versioned .xlsx")
    parser.add_argument("--db", default="radar_ram-01.sqlite", help="Output versioned SQLite DB")
    parser.add_argument(
        "--simulate-edit",
        action="store_true",
        help="Automatically simulate friend's edits: degrade PN-102 MTBF to 5000 hrs and add Tier-4 PN-109 under PN-105",
    )
    parser.add_argument(
        "--trace-part",
        default=None,
        help="Part_ID to trace upward in 2_Blast_Radius_CTE (defaults to PN-109 if --simulate-edit else PN-105)",
    )
    args = parser.parse_args()

    in_path = Path(args.input)
    out_path = Path(args.output)
    db_path = Path(args.db)

    baseline_xls = pd.ExcelFile(in_path)
    baseline_risk = (
        baseline_xls.parse("1_Critical_Risk_Query")
        if "1_Critical_Risk_Query" in baseline_xls.sheet_names
        else None
    )

    tables = load_from_excel(in_path)

    if args.simulate_edit:
        tables["Parts"].loc[tables["Parts"]["Part_ID"] == "PN-102", "MTBF_Hours"] = 5000.0

        new_part = pd.DataFrame([{
            "Part_ID": "PN-109",
            "Part_Name": "AlGaAs Infrared Die",
            "Subsystem": "Power",
            "MTBF_Hours": 11000.0,
            "Parent_Part_ID": "PN-105",
        }])
        tables["Parts"] = pd.concat([tables["Parts"], new_part], ignore_index=True)

        new_fm = pd.DataFrame([{
            "Failure_ID": "FM-06",
            "Part_ID": "PN-109",
            "Mode_Name": "Quantum Well Dark Current",
            "Severity": "Major",
        }])
        tables["Failure_Modes"] = pd.concat([tables["Failure_Modes"], new_fm], ignore_index=True)

    target_part = args.trace_part or ("PN-109" if args.simulate_edit else "PN-105")
    updated_tables = ingest_and_query(tables, db_path, target_part, baseline_risk)
    export_styled_excel(updated_tables, out_path)
    print(f"Created {out_path} and {db_path} (traced {target_part})")


if __name__ == "__main__":
    main()
