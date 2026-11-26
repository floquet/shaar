#!/usr/bin/python3
"""
Radar RAM Relational Database -> Normalized & Analytical Excel Generator
Builds an in-memory SQLite database, executes relational & recursive CTE queries,
and exports a formatted .xlsx workbook via pandas + openpyxl without launching Excel.
"""

import sqlite3
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def build_and_query_db(db_path: str = "radar_ram.sqlite") -> dict[str, pd.DataFrame]:
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    cur.executescript("""
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
        MTBF_Hours REAL NOT NULL,
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
    """)

    requirements = [
        ("REQ-101", "Long-Range Target Track", "Detect 1 m^2 RCS at 150 km"),
        ("REQ-102", "Low-RCS Discrimination", "Track 0.01 m^2 RCS at 60 km"),
        ("REQ-103", "Continuous 360-Deg Scan", "24/7 duty cycle without thermal derating"),
    ]

    parts = [
        ("PN-884", "GaN T/R Amplifier Module", "Transmitter", 12000.0, None),
        ("PN-102", "HV Pulsed Capacitor Bank", "Power", 8500.0, "PN-884"),
        ("PN-105", "Gate Driver Optocoupler", "Power", 9200.0, "PN-102"),
        ("PN-409", "Beamforming FPGA Board", "Signal Proc", 25000.0, None),
        ("PN-712", "Liquid Cold-Plate Pump", "Thermal", 7800.0, None),
    ]

    req_part_map = [
        ("REQ-101", "PN-884"),
        ("REQ-101", "PN-102"),
        ("REQ-101", "PN-712"),
        ("REQ-102", "PN-884"),
        ("REQ-102", "PN-409"),
        ("REQ-103", "PN-712"),
        ("REQ-103", "PN-884"),
    ]

    failure_modes = [
        ("FM-01", "PN-884", "Channel Thermal Runaway", "Critical"),
        ("FM-02", "PN-102", "Dielectric Breakdown", "Critical"),
        ("FM-03", "PN-105", "LED Degradation Drift", "Major"),
        ("FM-04", "PN-409", "BGA Solder Joint Fatigue", "Minor"),
        ("FM-05", "PN-712", "Impeller Bearing Seizure", "Critical"),
    ]

    cur.executemany("INSERT INTO Requirements VALUES (?, ?, ?);", requirements)
    cur.executemany("INSERT INTO Parts VALUES (?, ?, ?, ?, ?);", parts)
    cur.executemany("INSERT INTO Req_Part_Map VALUES (?, ?);", req_part_map)
    cur.executemany("INSERT INTO Failure_Modes VALUES (?, ?, ?, ?);", failure_modes)
    conn.commit()

    # Query 1: Critical requirement vulnerability rollup (MTBF < 10,000 hrs)
    sql_critical_rollup = """
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

    # Query 2: Recursive BOM blast-radius trace starting from sub-component PN-105
    sql_blast_radius = """
    WITH RECURSIVE Assembly_Chain(Part_ID, Part_Name, Parent_Part_ID, Depth, Path) AS (
        SELECT Part_ID, Part_Name, Parent_Part_ID, 0, Part_ID || ' (' || Part_Name || ')'
        FROM Parts
        WHERE Part_ID = 'PN-105'
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

    tables = {
        "1_Critical_Risk_Query": pd.read_sql_query(sql_critical_rollup, conn),
        "2_Blast_Radius_CTE": pd.read_sql_query(sql_blast_radius, conn),
        "Requirements": pd.read_sql_query("SELECT * FROM Requirements;", conn),
        "Parts": pd.read_sql_query("SELECT * FROM Parts;", conn),
        "Req_Part_Map": pd.read_sql_query("SELECT * FROM Req_Part_Map;", conn),
        "Failure_Modes": pd.read_sql_query("SELECT * FROM Failure_Modes;", conn),
    }
    conn.close()
    return tables


def export_styled_excel(tables: dict[str, pd.DataFrame], xlsx_path: str = "radar_ram_demo.xlsx") -> None:
    header_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    query_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
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

            is_query_tab = sheet_name[0].isdigit()
            fill_to_use = query_header_fill if is_query_tab else header_fill

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
                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 52)

            for row_idx in range(2, len(df) + 2):
                for col_idx in range(1, len(df.columns) + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell.font = body_font
                    cell.border = thin_border
                    if row_idx % 2 == 0:
                        cell.fill = zebra_fill


if __name__ == "__main__":
    dataframes = build_and_query_db("radar_ram.sqlite")
    export_styled_excel(dataframes, "radar_ram_demo.xlsx")
    print("Generated radar_ram.sqlite and radar_ram_demo.xlsx")
