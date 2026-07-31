"""Shared parsing/normalization helpers for ticket & project spreadsheets.

Used by both the Flask app (api/index.py) and the upload endpoint / migration
script, so a CSV upload and an Excel sheet go through the exact same
standardization before hitting the database.
"""
import re

import pandas as pd

HEADER_ROW = 1

COLORS = {
    "Completed": "#34d399",
    "Closed": "#60a5fa",
    "Pending": "#fbbf24",
    "InProgress": "#f87171",
    "Inprogress": "#f87171",
    "Open": "#a78bfa",
    "Cancelled": "#94a3b8",
    "On Hold": "#22d3ee",
}
PRIORITY_COLORS = {"High": "#f87171", "Medium": "#60a5fa", "Low": "#34d399"}
AGEING_COLORS = {"1-30 Days": "#34d399", "31-60 Days": "#fbbf24", "> 60 Days": "#f87171"}

COLUMN_MAPPING = {
    "ticket no": "Ticket No",
    "ticket number": "Ticket No",
    "ticket_no": "Ticket No",
    "ticket id": "Ticket No",
    "task type": "Task Type",
    "task_type": "Task Type",
    "project": "Project",
    "company": "Company",
    "ticket title": "Ticket Title",
    "ticket_title": "Ticket Title",
    "ticket detail": "Ticket Detail",
    "ticket_detail": "Ticket Detail",
    "ticket desc": "Ticket Detail",
    "detail": "Ticket Detail",
    "description": "Ticket Detail",
    "ticket category": "Ticket Category",
    "ticket_category": "Ticket Category",
    "category": "Ticket Category",
    "priority": "Priority",
    "ticket created date": "Ticket Created Date",
    "ticket_created_date": "Ticket Created Date",
    "created date": "Ticket Created Date",
    "created": "Ticket Created Date",
    "date created": "Ticket Created Date",
    "ticket completed date": "Ticket Completed Date",
    "ticket_completed_date": "Ticket Completed Date",
    "completed date": "Ticket Completed Date",
    "completed": "Ticket Completed Date",
    "ticket closed date": "Ticket Closed Date",
    "ticket_closed_date": "Ticket Closed Date",
    "closed date": "Ticket Closed Date",
    "closed": "Ticket Closed Date",
    "ticket status": "Ticket Status",
    "ticket_status": "Ticket Status",
    "status": "Ticket Status",
    "sla dateline": "SLA Dateline",
    "sla_deadline": "SLA Dateline",
    "sla": "SLA Dateline",
    "sla late": "SLA Late",
    "sla_breach": "SLA Late",
    "days": "Days",
    "ageing": "Ageing",
    "aging": "Ageing",
    "client": "Client",
}

# Ticket dataframe columns, in the order they're stored in the DB.
TICKET_COLUMNS = [
    "Client", "Ticket No", "Task Type", "Project", "Company", "Ticket Title",
    "Ticket Detail", "Ticket Category", "Priority", "Ticket Created Date",
    "Ticket Completed Date", "Ticket Closed Date", "Ticket Status",
    "SLA Dateline", "SLA Late", "Days", "Ageing", "Days to Close",
    "SLA Breach", "Source File",
]

PROJECT_COLUMNS = [
    "Client", "Title", "Category", "Progress", "Priority", "Start date",
    "Due date", "Target Date", "Assigned to", "Status Progress",
    "Percentage", "Overall Progress Task (%)", "Source File",
]


def standardize_columns(df):
    col_map = {}
    for col in df.columns:
        cl = str(col).lower().strip()
        if cl in COLUMN_MAPPING:
            col_map[col] = COLUMN_MAPPING[cl]
    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df = df.loc[:, ~df.columns.duplicated()]
    return df


def convert_dtypes(df):
    for date_col in ["Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date", "SLA Dateline"]:
        if date_col in df.columns:
            df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)

    if "Ticket Status" in df.columns:
        ts = df["Ticket Status"]
        if isinstance(ts, pd.DataFrame):
            ts = ts.iloc[:, 0]
        df["Ticket Status"] = ts.astype(str).str.strip()
        df["Ticket Status"] = df["Ticket Status"].replace({
            "InProgress": "In Progress",
            "Inprogress": "In Progress",
            "nan": None,
        })

    if "Priority" in df.columns:
        pr = df["Priority"]
        if isinstance(pr, pd.DataFrame):
            pr = pr.iloc[:, 0]
        df["Priority"] = pr.astype(str).str.strip().str.title()
        df["Priority"] = df["Priority"].replace({"Nan": None, "None": None})

    if "Ticket Created Date" in df.columns and "Ticket Closed Date" in df.columns:
        mask = df["Ticket Created Date"].notna() & df["Ticket Closed Date"].notna()
        df.loc[mask, "Days to Close"] = (df.loc[mask, "Ticket Closed Date"] - df.loc[mask, "Ticket Created Date"]).dt.days
    elif "Ticket Created Date" in df.columns and "Ticket Completed Date" in df.columns:
        mask = df["Ticket Created Date"].notna() & df["Ticket Completed Date"].notna()
        df.loc[mask, "Days to Close"] = (df.loc[mask, "Ticket Completed Date"] - df.loc[mask, "Ticket Created Date"]).dt.days

    if "Days" in df.columns:
        df["Days"] = pd.to_numeric(df["Days"], errors="coerce")

    if "Ageing" in df.columns:
        df["Ageing"] = df["Ageing"].astype(str).str.strip().replace({
            "1-30 days": "1-30 Days",
            "1-30 Days": "1-30 Days",
            "30-60 Days": "31-60 Days",
            "30-60 days": "31-60 Days",
            "31-60 Days": "31-60 Days",
            "31-60 days": "31-60 Days",
            ">60 Days": "> 60 Days",
            ">60 days": "> 60 Days",
            "> 60 Days": "> 60 Days",
            "> 60 days": "> 60 Days",
            "nan": None,
        })
    if "Ageing" not in df.columns:
        if "Days" in df.columns:
            def classify(d):
                if pd.isna(d):
                    return None
                if d <= 30:
                    return "1-30 Days"
                elif d <= 60:
                    return "31-60 Days"
                else:
                    return "> 60 Days"
            df["Ageing"] = df["Days"].apply(classify)
        elif "Ticket Created Date" in df.columns:
            days_open = (pd.Timestamp.now() - df["Ticket Created Date"]).dt.days
            def classify2(d):
                if pd.isna(d):
                    return None
                if d <= 30:
                    return "1-30 Days"
                elif d <= 60:
                    return "31-60 Days"
                else:
                    return "> 60 Days"
            df["Ageing"] = days_open.apply(classify2)

    if "SLA Late" in df.columns:
        sla = df["SLA Late"]
        if isinstance(sla, pd.DataFrame):
            sla = sla.iloc[:, 0]
        df["SLA Late"] = sla.astype(str).str.strip()
        # SLA Late holds a numeric days-remaining value; negative means
        # the ticket blew past its deadline.
        sla_num = pd.to_numeric(df["SLA Late"].replace({"nan": None}), errors="coerce")
        df["SLA Breach"] = sla_num < 0
    else:
        df["SLA Breach"] = False

    return df


def parse_ticket_sheet(df, client, source_file):
    """Standardize a raw ticket sheet/CSV into the canonical ticket dataframe.

    Returns (parsed_df, info) where info reports what got left behind, so
    callers can surface it instead of rows/columns silently vanishing:
      - rows_dropped: rows with no Ticket No at all (blank separator/
        subtotal rows that survive the initial dropna(how="all") because
        some other cell in the row is filled in).
      - unmapped_columns: source columns that didn't match anything in
        COLUMN_MAPPING and aren't part of the fixed ticket schema, so
        their data isn't stored (e.g. a "Remarks" or "Assigned To" column
        the spreadsheet has that this dashboard has no field for).
    """
    df = df.dropna(how="all")
    df = standardize_columns(df)
    if df.empty:
        return df, {"rows_dropped": 0, "unmapped_columns": []}

    if "Client" not in df.columns or df["Client"].isna().all():
        df["Client"] = client
    else:
        df["Client"] = df["Client"].fillna(client)
    df["Source File"] = source_file

    df = convert_dtypes(df)

    unmapped_columns = [c for c in df.columns if c not in TICKET_COLUMNS]

    rows_dropped = 0
    if "Ticket No" in df.columns:
        tn = df["Ticket No"]
        if isinstance(tn, pd.DataFrame):
            tn = tn.iloc[:, 0]
        tn = tn.astype(str).str.strip()
        valid = tn.ne("") & tn.str.lower().ne("nan") & tn.str.lower().ne("none")
        rows_dropped = int((~valid).sum())
        df = df[valid]

    for col in TICKET_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[TICKET_COLUMNS], {"rows_dropped": rows_dropped, "unmapped_columns": unmapped_columns}


def parse_project_sheet(df, source_file):
    df = df.dropna(how="all").copy()
    if "Client" in df.columns:
        df["Client"] = df["Client"].ffill()
    df["Source File"] = source_file

    for c in ["Start date", "Due date", "Target Date"]:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
    if "Percentage" in df.columns:
        df["Percentage"] = pd.to_numeric(df["Percentage"].astype(str).str.replace("%", ""), errors="coerce")
    if "Overall Progress Task (%)" in df.columns:
        df["Overall Progress Task (%)"] = pd.to_numeric(df["Overall Progress Task (%)"].astype(str).str.replace("%", ""), errors="coerce")

    for col in PROJECT_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[PROJECT_COLUMNS]


def detect_ticket_sheets(filepath_or_buffer):
    """Return {sheet_name: header_row} for sheets that look like ticket sheets.

    Tries HEADER_ROW first (row 1, matching the original bundled workbook's
    layout) then falls back to rows 0 and 2, since a sheet exported from a
    different tool can put the real header one row up or down -- previously
    a sheet like that wasn't recognized as a ticket sheet at all and its
    entire client's worth of tickets went missing from the upload with no
    indication why.
    """
    xl = pd.ExcelFile(filepath_or_buffer, engine="openpyxl")
    ticket_sheets = {}
    for name in xl.sheet_names:
        for header_row in (HEADER_ROW, 0, 2):
            try:
                df_head = pd.read_excel(filepath_or_buffer, sheet_name=name, header=header_row, engine="openpyxl", nrows=3)
            except Exception:
                continue
            # pandas dedupes repeated header names as "Ticket No.1",
            # "Ticket No.2", ... -- strip that suffix before checking so a
            # printed/aggregate report with the same block of columns
            # repeated side by side several times is actually recognized
            # as repeated, not read as a single "Ticket No" column.
            base_names = [re.sub(r"\.\d+$", "", c) for c in (str(c).lower().strip() for c in df_head.columns)]
            ticket_no_matches = sum(1 for c in base_names if COLUMN_MAPPING.get(c) == "Ticket No")
            # More than one match means a repeated-block layout, which
            # isn't a one-row-per-ticket sheet and would wrongly become a
            # "client" named after the sheet -- skip it.
            if ticket_no_matches == 1:
                ticket_sheets[name] = header_row
                break
    return ticket_sheets
