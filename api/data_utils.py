"""Shared parsing/normalization helpers for ticket & project spreadsheets.

Used by both the Flask app (api/index.py) and the upload endpoint / migration
script, so a CSV upload and an Excel sheet go through the exact same
standardization before hitting the database.
"""
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

    Applies the same "Ticket No must start with T/" filter the original
    file-scanning loader used, so uploads behave identically to the old
    bundled-Excel flow.
    """
    df = df.dropna(how="all")
    df = standardize_columns(df)
    if df.empty:
        return df

    if "Client" not in df.columns or df["Client"].isna().all():
        df["Client"] = client
    else:
        df["Client"] = df["Client"].fillna(client)
    df["Source File"] = source_file

    df = convert_dtypes(df)

    if "Ticket No" in df.columns:
        tn = df["Ticket No"]
        if isinstance(tn, pd.DataFrame):
            tn = tn.iloc[:, 0]
        mask = tn.astype(str).str.match(r"^T/", na=False)
        df = df[mask]

    for col in TICKET_COLUMNS:
        if col not in df.columns:
            df[col] = None

    return df[TICKET_COLUMNS]


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
    """Return sheet names in an xlsx that look like ticket sheets."""
    xl = pd.ExcelFile(filepath_or_buffer, engine="openpyxl")
    ticket_sheets = []
    for name in xl.sheet_names:
        try:
            df_head = pd.read_excel(filepath_or_buffer, sheet_name=name, header=HEADER_ROW, engine="openpyxl", nrows=3)
            if "Ticket No" in df_head.columns or "ticket no" in str(df_head.columns).lower():
                ticket_sheets.append(name)
        except Exception:
            pass
    return ticket_sheets
