import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.io as pio
import os, sys, glob, json
from datetime import datetime, timedelta
from flask import Flask, render_template, request

app = Flask(
    __name__,
    template_folder=os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "templates"),
)

def log(msg, level="INFO"):
    print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} {level} {msg}", flush=True)

log("=" * 50)
log(f"Dashboard starting (Flask)")
log(f"Python: {sys.version}")
log(f"CWD: {os.getcwd()}")

pio.templates.default = "plotly_white"

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
HEADER_ROW = 1

COLORS = {
    "Completed": "#2ecc71",
    "Closed": "#3498db",
    "Pending": "#f39c12",
    "InProgress": "#e74c3c",
    "Inprogress": "#e74c3c",
    "Open": "#9b59b6",
    "Cancelled": "#95a5a6",
    "On Hold": "#1abc9c",
}
PRIORITY_COLORS = {"High": "#e74c3c", "Medium": "#3498db", "Low": "#2ecc71"}
AGEING_COLORS = {"1-30 Days": "#2ecc71", "31-60 Days": "#e67e22", "> 60 Days": "#e74c3c"}

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
}


def find_excel_files():
    files = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
    files = [f for f in files if not os.path.basename(f).startswith("~")]
    files.sort(key=lambda f: os.path.getmtime(f), reverse=True)
    return files


def detect_ticket_sheets(filepath):
    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
        ticket_sheets = []
        for name in xl.sheet_names:
            try:
                df_head = pd.read_excel(filepath, sheet_name=name, header=HEADER_ROW, engine="openpyxl", nrows=3)
                if "Ticket No" in df_head.columns or "ticket no" in str(df_head.columns).lower():
                    ticket_sheets.append(name)
            except Exception:
                pass
        return ticket_sheets
    except Exception as e:
        log(f"  Cannot open {filepath}: {e}", "ERROR")
        return []


def standardize_columns(df):
    col_map = {}
    for col in df.columns:
        cl = str(col).lower().strip()
        if cl in COLUMN_MAPPING:
            col_map[col] = COLUMN_MAPPING[cl]
    df = df.rename(columns=col_map)
    df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
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
        df["SLA Breach"] = df["SLA Late"].apply(lambda x: x.lower() in ("yes", "1", "true", "late", "y") if pd.notna(x) else False)
    else:
        df["SLA Breach"] = False

    return df


def load_data():
    all_dfs = []
    load_errors = []

    excel_files = find_excel_files()
    log(f"Found {len(excel_files)} Excel files: {[os.path.basename(f) for f in excel_files]}")

    for filepath in excel_files:
        fname = os.path.basename(filepath)
        log(f"Scanning: {fname}")
        sheets = detect_ticket_sheets(filepath)
        log(f"  Ticket sheets found: {sheets}")

        for sheet_name in sheets:
            try:
                log(f"  Reading sheet: {sheet_name}")
                df = pd.read_excel(filepath, sheet_name=sheet_name, header=HEADER_ROW, engine="openpyxl")
                df = df.dropna(how="all")
                df = standardize_columns(df)
                if df.empty:
                    log(f"  Empty after standardize")
                    continue

                df["Client"] = sheet_name
                df["Source File"] = fname
                all_dfs.append(df)
                log(f"  Loaded: {len(df)} rows, cols: {list(df.columns)[:10]}")
            except Exception as e:
                msg = f"{fname}/{sheet_name}: {str(e)[:200]}"
                log(msg, "ERROR")
                load_errors.append(msg)

    if not all_dfs:
        log("No data loaded from any file!", "ERROR")
        return pd.DataFrame(), load_errors

    df = pd.concat(all_dfs, ignore_index=True)
    log(f"Combined from all files: {len(df)} rows")

    df = convert_dtypes(df)

    if "Ticket No" in df.columns:
        tn = df["Ticket No"]
        if isinstance(tn, pd.DataFrame):
            tn = tn.iloc[:, 0]
        mask = tn.astype(str).str.match(r"^T/", na=False)
        df = df[mask]
        log(f"After Ticket No filter: {len(df)} rows")
    else:
        log("Ticket No column missing in all data!", "WARNING")

    log(f"Final: {len(df)} rows, {len(df.columns)} cols")
    return df, load_errors


def load_project_data():
    try:
        files = find_excel_files()
        for fp in files:
            xl = pd.ExcelFile(fp, engine="openpyxl")
            if "Client Project" in xl.sheet_names:
                df = pd.read_excel(fp, sheet_name="Client Project", header=0, engine="openpyxl")
                df = df.dropna(how="all")
                if "Client" in df.columns:
                    df["Client"] = df["Client"].ffill()
                for c in ["Start date", "Due date", "Target Date"]:
                    if c in df.columns:
                        df[c] = pd.to_datetime(df[c], errors="coerce", dayfirst=True)
                if "Percentage" in df.columns:
                    df["Percentage"] = pd.to_numeric(df["Percentage"].astype(str).str.replace("%", ""), errors="coerce")
                if "Overall Progress Task (%)" in df.columns:
                    df["Overall Progress Task (%)"] = pd.to_numeric(df["Overall Progress Task (%)"].astype(str).str.replace("%", ""), errors="coerce")
                return df
        return pd.DataFrame()
    except Exception as e:
        log(f"Error loading project data: {e}", "ERROR")
        return pd.DataFrame()


def build_warranty_charts(df):
    charts = {}
    warranty_df = df[df["Client"] == "Client Warranty"].copy() if "Client" in df.columns else pd.DataFrame()
    if warranty_df.empty:
        return charts

    total = len(warranty_df)
    completed = len(warranty_df[warranty_df["Ticket Status"].isin(["Completed", "Closed"])]) if "Ticket Status" in warranty_df.columns else 0
    pending = len(warranty_df[warranty_df["Ticket Status"] == "Pending"]) if "Ticket Status" in warranty_df.columns else 0
    in_progress = len(warranty_df[warranty_df["Ticket Status"] == "In Progress"]) if "Ticket Status" in warranty_df.columns else 0
    sla_breach = warranty_df["SLA Breach"].sum() if "SLA Breach" in warranty_df.columns else 0

    charts["metrics"] = {
        "total": total, "completed": completed, "pending": pending,
        "in_progress": in_progress, "sla_breach": int(sla_breach),
        "completed_pct": f"{completed / total * 100:.1f}%" if total > 0 else "0%",
        "pending_pct": f"{pending / total * 100:.1f}%" if total > 0 else "0%",
        "in_progress_pct": f"{in_progress / total * 100:.1f}%" if total > 0 else "0%",
    }

    if "Ticket Status" in warranty_df.columns:
        sc = warranty_df["Ticket Status"].value_counts().reset_index()
        sc.columns = ["Status", "Bilangan"]
        fig = px.pie(sc, names="Status", values="Bilangan", title="Status Tiket Warranty",
                      color="Status", color_discrete_map=COLORS, hole=0.3)
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["status_pie"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Task Type" in warranty_df.columns:
        tc = warranty_df["Task Type"].value_counts().reset_index()
        tc.columns = ["Task Type", "Bilangan"]
        fig = px.bar(tc, x="Task Type", y="Bilangan", title="Tiket Warranty mengikut Task Type",
                      color="Task Type", text="Bilangan")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["task_type_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Project" in warranty_df.columns:
        pc = warranty_df["Project"].value_counts().reset_index()
        pc.columns = ["Project", "Bilangan"]
        fig = px.bar(pc, x="Project", y="Bilangan", title="Tiket Warranty mengikut Projek",
                      color="Project", text="Bilangan")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["project_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    display_cols = ["Ticket No", "Task Type", "Project", "Company", "Ticket Title", "Priority", "Ticket Status", "Ticket Created Date", "Days"]
    avail = [c for c in display_cols if c in warranty_df.columns]
    detail = warranty_df[avail].copy()
    if "Ticket Created Date" in detail.columns:
        detail["Ticket Created Date"] = detail["Ticket Created Date"].dt.strftime("%d/%m/%Y")
    charts["detail_table"] = detail.to_html(index=False)

    return charts


def build_project_charts(df):
    charts = {}
    if df.empty:
        return charts

    total = len(df)
    completed = len(df[df["Status Progress"].str.lower().str.contains("completed", na=False)]) if "Status Progress" in df.columns else 0
    in_progress = len(df[df["Status Progress"].str.lower().str.contains("progress", na=False)]) if "Status Progress" in df.columns else 0
    not_started = len(df[df["Status Progress"].str.lower().str.contains("not started", na=False)]) if "Status Progress" in df.columns else 0

    charts["metrics"] = {
        "total": total, "completed": completed,
        "in_progress": in_progress, "not_started": not_started,
    }

    if "Client" in df.columns:
        valid_clients = df.dropna(subset=["Client"])
        if valid_clients.empty:
            return charts
        cc = valid_clients["Client"].value_counts().reset_index()
        cc.columns = ["Client", "Bilangan"]
        fig = px.bar(cc, x="Client", y="Bilangan", title="Projek mengikut Client",
                      color="Client", text="Bilangan")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False)
        charts["client_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Status Progress" in df.columns:
        valid_status = df.dropna(subset=["Status Progress"])
        if valid_status.empty:
            return charts
        sc = valid_status["Status Progress"].value_counts().reset_index()
        sc.columns = ["Status", "Bilangan"]
        fig = px.pie(sc, names="Status", values="Bilangan", title="Status Progress Projek",
                      hole=0.3)
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["status_pie"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Start date" in df.columns and "Due date" in df.columns and "Title" in df.columns:
        valid = df.dropna(subset=["Start date", "Due date", "Title"]).copy()
        valid = valid[valid["Title"].astype(str).str.strip() != ""]
        if "Description" in valid.columns:
            valid["Task Label"] = valid["Description"].astype(str)
            valid["Task Label"] = valid["Task Label"].str.replace(r"^\d+\.\s*", "", regex=True)
        else:
            valid["Task Label"] = valid["Title"].astype(str)
        if not valid.empty:
            timeline_charts_html = ""
            if "Client" in valid.columns:
                for client in sorted(valid["Client"].dropna().unique()):
                    cdf = valid[valid["Client"] == client]
                    if cdf.empty:
                        continue
                    fig = px.timeline(
                        cdf, x_start="Start date", x_end="Due date",
                        y="Task Label", color="Client",
                        title=f"{client} - PROJECT DEVELOPMENT TIMELINE",
                        color_discrete_sequence=px.colors.qualitative.Plotly,
                    )
                    fig.update_yaxes(autorange="reversed", title=None)
                    fig.update_xaxes(title="Tarikh")
                    fig.update_layout(
                        template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                        font=dict(color="#374151"), showlegend=False,
                        height=max(200, 30*len(cdf)),
                    )
                    timeline_charts_html += f'<div class="client-section"><h4>{client}</h4>{fig.to_html(full_html=False, config={"displayModeBar": False})}</div>'
            charts["timeline_chart"] = timeline_charts_html

    display_cols_p = ["Client", "Title", "Category", "Progress", "Priority", "Start date", "Due date", "Assigned to", "Status Progress", "Percentage", "Overall Progress Task (%)"]
    avail_p = [c for c in display_cols_p if c in df.columns]
    detail = df[avail_p].copy()
    for c in ["Start date", "Due date", "Target Date"]:
        if c in detail.columns:
            detail[c] = detail[c].dt.strftime("%d/%m/%Y") if not detail[c].isna().all() else detail[c]
    detail = detail.fillna("")
    charts["detail_table"] = detail.to_html(index=False)

    return charts


def apply_filters(df, args):
    if "Source File" in df.columns:
        sources = sorted(df["Source File"].unique())
        selected = args.getlist("source_file")
        if selected:
            df = df[df["Source File"].isin(selected)]

    if "Ticket Created Date" in df.columns:
        valid_dates = df["Ticket Created Date"].dropna()
        if len(valid_dates) > 0:
            start_str = args.get("date_start")
            end_str = args.get("date_end")
            if start_str:
                try:
                    start_date = datetime.strptime(start_str, "%Y-%m-%d").date()
                    df = df[df["Ticket Created Date"].dt.date >= start_date]
                except:
                    pass
            if end_str:
                try:
                    end_date = datetime.strptime(end_str, "%Y-%m-%d").date()
                    df = df[df["Ticket Created Date"].dt.date <= end_date]
                except:
                    pass

    if "Client" in df.columns:
        clients = sorted(df["Client"].unique())
        selected = args.getlist("client")
        if selected:
            df = df[df["Client"].isin(selected)]

    if "Priority" in df.columns:
        priorities = sorted(df["Priority"].dropna().unique())
        selected = args.getlist("priority")
        if selected:
            df = df[df["Priority"].isin(selected)]

    if "Ticket Status" in df.columns:
        statuses = sorted(df["Ticket Status"].dropna().unique())
        selected = args.getlist("status")
        if selected:
            df = df[df["Ticket Status"].isin(selected)]

    if "Task Type" in df.columns:
        task_types = sorted(df["Task Type"].dropna().unique())
        selected = args.getlist("task_type")
        if selected:
            df = df[df["Task Type"].isin(selected)]

    search_term = args.get("search", "")
    if search_term:
        mask = pd.Series([False] * len(df))
        for col in ["Ticket Detail", "Ticket Title", "Ticket No", "Ticket Category", "Company", "Project"]:
            if col in df.columns:
                mask = mask | df[col].astype(str).str.contains(search_term, case=False, na=False)
        df = df[mask]

    return df


def build_charts(df):
    charts = {}

    total = len(df)
    completed = len(df[df["Ticket Status"].isin(["Completed", "Closed"])]) if "Ticket Status" in df.columns else 0
    pending = len(df[df["Ticket Status"] == "Pending"]) if "Ticket Status" in df.columns else 0
    in_progress = len(df[df["Ticket Status"] == "In Progress"]) if "Ticket Status" in df.columns else 0
    sla_breach = df["SLA Breach"].sum() if "SLA Breach" in df.columns else 0
    avg_days = None
    if "Days to Close" in df.columns:
        valid_days = df["Days to Close"].dropna()
        if len(valid_days) > 0:
            avg_days = round(valid_days.mean(), 1)

    metrics = {
        "total": total,
        "completed": completed,
        "pending": pending,
        "in_progress": in_progress,
        "sla_breach": int(sla_breach),
        "avg_days": avg_days,
        "completed_pct": f"{completed / total * 100:.1f}%" if total > 0 else "0%",
        "pending_pct": f"{pending / total * 100:.1f}%" if total > 0 else "0%",
        "in_progress_pct": f"{in_progress / total * 100:.1f}%" if total > 0 else "0%",
    }

    charts["metrics"] = metrics

    if "Ticket Status" in df.columns:
        status_counts = df["Ticket Status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Bilangan"]
        fig = px.pie(
            status_counts, names="Status", values="Bilangan",
            title="Taburan Status Tiket", color="Status",
            color_discrete_map=COLORS, hole=0.3,
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["status_pie"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Priority" in df.columns:
        priority_counts = df["Priority"].value_counts().reset_index()
        priority_counts.columns = ["Keutamaan", "Bilangan"]
        fig = px.pie(
            priority_counts, names="Keutamaan", values="Bilangan",
            title="Taburan Keutamaan", color="Keutamaan",
            color_discrete_map=PRIORITY_COLORS, hole=0.3,
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), legend=dict(orientation="h", yanchor="bottom", y=-0.2))
        charts["priority_pie"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_dist = df["Client"].value_counts().reset_index()
        client_dist.columns = ["Client", "Bilangan"]
        client_colors = px.colors.qualitative.Plotly[:len(client_dist)]
        fig = px.bar(
            client_dist, x="Client", y="Bilangan",
            title="Tiket mengikut Client", color="Client",
            color_discrete_sequence=client_colors, text="Bilangan",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
        charts["client_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    return charts


def build_status_charts(df):
    charts = {}

    if "Ticket Status" not in df.columns:
        return charts

    status_by_client = df.groupby(["Client", "Ticket Status"]).size().reset_index(name="Bilangan")
    unique_statuses = status_by_client["Ticket Status"].unique()
    status_colors_seq = px.colors.qualitative.Plotly[:len(unique_statuses)]
    fig = px.bar(
        status_by_client, x="Client", y="Bilangan",
        color="Ticket Status", title="Status mengikut Client",
        color_discrete_sequence=status_colors_seq, barmode="stack",
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
    charts["status_client_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    status_counts = df["Ticket Status"].value_counts().reset_index()
    status_counts.columns = ["Status", "Bilangan"]
    statuses2 = status_counts["Status"].unique()
    colors2 = px.colors.qualitative.Plotly[:len(statuses2)]
    fig = px.bar(
        status_counts, x="Bilangan", y="Status",
        orientation="h", title="Jumlah mengikut Status",
        color="Status", color_discrete_sequence=colors2, text="Bilangan",
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
    charts["status_hbar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    pivot = df.groupby(["Client", "Ticket Status"]).size().unstack(fill_value=0)
    charts["status_pivot"] = pivot.to_html()

    return charts


def build_priority_charts(df):
    charts = {}

    if "Priority" not in df.columns:
        return charts

    priority_counts = df["Priority"].value_counts().reset_index()
    priority_counts.columns = ["Keutamaan", "Bilangan"]
    fig = px.pie(
        priority_counts, names="Keutamaan", values="Bilangan",
        title="Taburan Keutamaan", color="Keutamaan",
        color_discrete_map=PRIORITY_COLORS, hole=0.4,
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
    charts["priority_pie"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Ticket Status" in df.columns:
        cross = df.groupby(["Priority", "Ticket Status"]).size().reset_index(name="Bilangan")
        fig = px.bar(
            cross, x="Priority", y="Bilangan",
            color="Ticket Status", title="Keutamaan mengikut Status",
            color_discrete_map=COLORS, barmode="group",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["priority_status_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        pivot = df.groupby(["Client", "Priority"]).size().unstack(fill_value=0)
        charts["priority_client_pivot"] = pivot.to_html()

    return charts


def build_ageing_charts(df):
    charts = {}

    has_ageing = "Ageing" in df.columns and df["Ageing"].notna().any()
    has_days = "Days" in df.columns and df["Days"].notna().any()

    if not has_ageing and not has_days:
        return charts

    age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
    charts["age_order"] = age_order
    charts["ageing_clients"] = {}

    if has_ageing and "Client" in df.columns:
        total_all = df["Ageing"].notna().sum()
        charts["total_ageing"] = int(total_all)

        for client in sorted(df["Client"].unique()):
            dc = df[df["Client"] == client].dropna(subset=["Ageing"])
            if dc.empty:
                continue

            counts = dc["Ageing"].value_counts().reindex(age_order, fill_value=0).reset_index()
            counts.columns = ["Kumpulan Umur", "Bilangan"]

            fig = px.bar(
                counts, x="Kumpulan Umur", y="Bilangan",
                color="Kumpulan Umur", color_discrete_map=AGEING_COLORS,
                text="Bilangan", title=client,
            )
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False)
            charts["ageing_clients"][client] = {
                "count": int(len(dc)),
                "chart": fig.to_html(full_html=False, config={"displayModeBar": False}),
                "table": {k: int(counts.set_index("Kumpulan Umur").loc[k, "Bilangan"]) for k in age_order},
            }

    if has_days:
        valid_days = df["Days"].dropna()
        if len(valid_days) > 0:
            fig = px.histogram(
                df.dropna(subset=["Days"]), x="Days", nbins=30,
                title="Taburan Hari Terbuka", color_discrete_sequence=["#3498db"],
                marginal="box",
            )
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["days_hist"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "SLA Breach" in df.columns and "Client" in df.columns:
        sla_by_client = df.groupby("Client")["SLA Breach"].sum().reset_index()
        sla_by_client.columns = ["Client", "Pelanggaran SLA"]
        fig = px.bar(
            sla_by_client, x="Client", y="Pelanggaran SLA",
            title="Jumlah Pelanggaran SLA", color_discrete_sequence=["#e74c3c"],
            text="Pelanggaran SLA",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
        charts["sla_breach_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    return charts


def build_client_comparison_charts(df):
    charts = {}

    if "Client" not in df.columns:
        return charts

    client_stats = df.groupby("Client").agg(
        Jumlah=("Ticket No", "count") if "Ticket No" in df.columns else ("Ticket Status", "count"),
    ).reset_index()

    if "Ticket Status" in df.columns:
        status_counts = df.groupby(["Client", "Ticket Status"]).size().unstack(fill_value=0)
        client_stats = client_stats.merge(status_counts, on="Client", how="left")

    if "Days to Close" in df.columns:
        avg_days = df.groupby("Client")["Days to Close"].mean().reset_index()
        avg_days.columns = ["Client", "Purata Hari"]
        client_stats = client_stats.merge(avg_days, on="Client", how="left")

    if "SLA Breach" in df.columns:
        sla = df.groupby("Client")["SLA Breach"].sum().reset_index()
        sla.columns = ["Client", "Pelanggaran SLA"]
        client_stats = client_stats.merge(sla, on="Client", how="left")

    charts["client_stats_table"] = client_stats.to_html(index=False)

    fig = px.bar(
        client_stats, x="Client", y="Jumlah",
        title="Jumlah Tiket mengikut Client", color="Client", text="Jumlah",
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), showlegend=False, xaxis_tickangle=-45)
    charts["client_total_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    status_cols = [c for c in ["Completed", "Pending", "In Progress", "Closed"] if c in client_stats.columns]
    if status_cols:
        fig = go.Figure()
        palette = px.colors.qualitative.Plotly
        for i, col in enumerate(status_cols):
            fig.add_trace(go.Bar(name=col, x=client_stats["Client"], y=client_stats[col], marker_color=palette[i % len(palette)]))
        fig.update_layout(barmode="stack", title="Status mengikut Client", template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
        charts["client_status_stacked"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Priority" in df.columns:
        priority_dummies = pd.get_dummies(df[["Client", "Priority"]], columns=["Priority"])
        radar_data = priority_dummies.groupby("Client").sum().reset_index()
        categories = [c for c in radar_data.columns if c.startswith("Priority_")]
        if categories:
            fig = go.Figure()
            for _, row in radar_data.iterrows():
                values = [row[c] for c in categories]
                values.append(values[0])
                cats = [c.replace("Priority_", "") for c in categories]
                cats.append(cats[0])
                fig.add_trace(go.Scatterpolar(r=values, theta=cats, fill="toself", name=row["Client"]))
            fig.update_layout(polar=dict(radialaxis=dict(visible=True)), title="Profil Keutamaan mengikut Client", template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["client_radar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    return charts


def build_timeline_charts(df):
    charts = {}

    if "Ticket Created Date" not in df.columns:
        return charts

    df_dated = df[df["Ticket Created Date"].notna()].copy()
    if len(df_dated) == 0:
        return charts

    df_dated["Bulan"] = df_dated["Ticket Created Date"].dt.to_period("M").astype(str)
    monthly_created = df_dated.groupby("Bulan").size().reset_index(name="Dicipta")

    fig = px.line(
        monthly_created, x="Bulan", y="Dicipta",
        title="Tiket Dicipta mengikut Bulan", markers=True,
        color_discrete_sequence=["#3498db"],
    )
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
    charts["timeline_created"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Ticket Completed Date" in df.columns:
        df_completed = df[df["Ticket Completed Date"].notna()].copy()
        if len(df_completed) > 0:
            df_completed["Bulan"] = df_completed["Ticket Completed Date"].dt.to_period("M").astype(str)
            monthly_completed = df_completed.groupby("Bulan").size().reset_index(name="Selesai")

            merged = monthly_created.merge(monthly_completed, on="Bulan", how="left").fillna(0)
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=merged["Bulan"], y=merged["Dicipta"], mode="lines+markers", name="Dicipta", line=dict(color="#3498db", width=2)))
            fig.add_trace(go.Scatter(x=merged["Bulan"], y=merged["Selesai"], mode="lines+markers", name="Selesai", line=dict(color="#2ecc71", width=2)))
            fig.update_layout(title="Dicipta vs Selesai", template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), xaxis_tickangle=-45)
            charts["timeline_created_vs_completed"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_monthly = df_dated.groupby(["Bulan", "Client"]).size().reset_index(name="Bilangan")
        fig = px.area(client_monthly, x="Bulan", y="Bilangan", color="Client", title="Tiket mengikut Client dan Bulan")
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["timeline_client_area"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Ticket Category" in df.columns:
        cat_monthly = df_dated.groupby(["Bulan", "Ticket Category"]).size().reset_index(name="Bilangan")
        if len(cat_monthly) > 0:
            top_cats = df_dated["Ticket Category"].value_counts().head(8).index.tolist()
            cat_monthly = cat_monthly[cat_monthly["Ticket Category"].isin(top_cats)]
            fig = px.line(cat_monthly, x="Bulan", y="Bilangan", color="Ticket Category", title="Tiket mengikut Kategori (Top 8)", markers=True)
            fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
            charts["timeline_category"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    return charts


def build_sla_charts(df):
    charts = {}

    if "SLA Breach" not in df.columns:
        return charts

    total = len(df)
    breaches = df["SLA Breach"].sum()
    compliance_rate = round((total - breaches) / total * 100, 1) if total > 0 else 0
    charts["compliance_rate"] = compliance_rate
    charts["total_breaches"] = int(breaches)
    charts["total_compliant"] = int(total - breaches)

    fig = go.Figure(go.Indicator(
        mode="gauge+number", value=compliance_rate,
        title={"text": "Kadar Pematuhan SLA (%)"},
        gauge=dict(
            axis=dict(range=[0, 100]), bar=dict(color="#2ecc71"),
            steps=[
                dict(range=[0, 50], color="#e74c3c"),
                dict(range=[50, 75], color="#f39c12"),
                dict(range=[75, 100], color="#2ecc71"),
            ],
            threshold=dict(line=dict(color="white", width=2), thickness=0.75, value=compliance_rate),
        ),
    ))
    fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"), height=350)
    charts["sla_gauge"] = fig.to_html(full_html=False, config={"displayModeBar": False})

    if "Client" in df.columns:
        client_sla = df.groupby("Client").agg(Total=("SLA Breach", "count"), Breaches=("SLA Breach", "sum")).reset_index()
        client_sla["Kadar Pematuhan (%)"] = ((client_sla["Total"] - client_sla["Breaches"]) / client_sla["Total"] * 100).round(1)
        client_sla = client_sla.sort_values("Kadar Pematuhan (%)", ascending=True)

        fig = px.bar(
            client_sla, x="Kadar Pematuhan (%)", y="Client",
            orientation="h", title="Kadar Pematuhan SLA mengikut Client",
            color="Kadar Pematuhan (%)", color_continuous_scale=["#e74c3c", "#f39c12", "#2ecc71"],
            text="Kadar Pematuhan (%)",
        )
        fig.update_layout(template="plotly_white", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#374151"))
        charts["sla_client_bar"] = fig.to_html(full_html=False, config={"displayModeBar": False})

        if "Ticket Status" in df.columns:
            sla_pivot = df.groupby(["Client", "Ticket Status"])["SLA Breach"].agg(["sum", "count", "mean"]).reset_index()
            sla_pivot.columns = ["Client", "Status", "Pelanggaran", "Jumlah", "Kadar Pelanggaran"]
            sla_pivot["Kadar Pelanggaran"] = (sla_pivot["Kadar Pelanggaran"] * 100).round(1)
            charts["sla_pivot"] = sla_pivot.to_html(index=False)

    return charts


def get_filter_options(df):
    options = {}
    if "Source File" in df.columns:
        options["source_files"] = sorted(df["Source File"].unique())
    if "Client" in df.columns:
        options["clients"] = sorted(df["Client"].unique())
    if "Priority" in df.columns:
        options["priorities"] = sorted(df["Priority"].dropna().unique())
    if "Ticket Status" in df.columns:
        options["statuses"] = sorted(df["Ticket Status"].dropna().unique())
    if "Task Type" in df.columns:
        options["task_types"] = sorted(df["Task Type"].dropna().unique())
    if "Ticket Created Date" in df.columns:
        valid_dates = df["Ticket Created Date"].dropna()
        if len(valid_dates) > 0:
            options["min_date"] = valid_dates.min().strftime("%Y-%m-%d")
            options["max_date"] = valid_dates.max().strftime("%Y-%m-%d")
    return options


@app.route("/")
def index():
    df, load_errors = load_data()
    df_filtered = apply_filters(df, request.args)

    filter_options = get_filter_options(df)
    filter_options["search"] = request.args.get("search", "")
    filter_options["date_start"] = request.args.get("date_start", "")
    filter_options["date_end"] = request.args.get("date_end", "")
    filter_options["selected_source_files"] = request.args.getlist("source_file")
    filter_options["selected_clients"] = request.args.getlist("client")
    filter_options["selected_priorities"] = request.args.getlist("priority")
    filter_options["selected_statuses"] = request.args.getlist("status")
    filter_options["selected_task_types"] = request.args.getlist("task_type")

    has_data = not df_filtered.empty
    data_info = {
        "total_raw": len(df),
        "total_filtered": len(df_filtered),
        "excel_files": [os.path.basename(f) for f in find_excel_files()],
        "load_errors": load_errors,
        "columns": list(df.columns) if not df.empty else [],
        "source_files": list(df["Source File"].unique()) if not df.empty and "Source File" in df.columns else [],
    }

    project_df = load_project_data()
    has_project = not project_df.empty

    if has_data:
        overview_charts = build_charts(df_filtered)
        status_charts = build_status_charts(df_filtered)
        priority_charts = build_priority_charts(df_filtered)
        ageing_charts = build_ageing_charts(df_filtered)
        comparison_charts = build_client_comparison_charts(df_filtered)
        timeline_charts = build_timeline_charts(df_filtered)
        sla_charts = build_sla_charts(df_filtered)
        warranty_charts = build_warranty_charts(df_filtered)
        project_charts = build_project_charts(project_df)
    else:
        overview_charts = status_charts = priority_charts = ageing_charts = {}
        comparison_charts = timeline_charts = sla_charts = {}
        warranty_charts = project_charts = {}

    display_cols = [
        "Client", "Ticket No", "Task Type", "Project", "Company",
        "Ticket Title", "Ticket Category", "Priority", "Ticket Status",
        "Ticket Created Date", "Ticket Completed Date", "Ticket Closed Date",
        "Days to Close", "Ageing", "SLA Breach",
    ]
    avail_cols = [c for c in display_cols if c in df_filtered.columns]
    detail_df = df_filtered[avail_cols].copy() if has_data and avail_cols else pd.DataFrame()

    if "Ticket Created Date" in detail_df.columns:
        detail_df["Ticket Created Date"] = detail_df["Ticket Created Date"].dt.strftime("%d/%m/%Y")
    if "Ticket Completed Date" in detail_df.columns:
        detail_df["Ticket Completed Date"] = detail_df["Ticket Completed Date"].dt.strftime("%d/%m/%Y")
    if "Ticket Closed Date" in detail_df.columns:
        detail_df["Ticket Closed Date"] = detail_df["Ticket Closed Date"].dt.strftime("%d/%m/%Y")

    ageing_list_data = {}
    if has_data and "Ageing" in df_filtered.columns and df_filtered["Ageing"].notna().sum() > 0:
        age_order = ["1-30 Days", "31-60 Days", "> 60 Days"]
        ageing_cols = [c for c in ["Client", "Ticket No", "Ticket Title", "Ticket Status", "Priority", "Ticket Created Date", "Days"] if c in df_filtered.columns]
        ageing_df = df_filtered.dropna(subset=["Ageing"]).copy()
        if "Ticket Created Date" in ageing_df.columns:
            ageing_df["Ticket Created Date"] = ageing_df["Ticket Created Date"].dt.strftime("%d/%m/%Y")
        ageing_list_data = {
            "total": int(ageing_df["Ageing"].notna().sum()),
            "buckets": {},
        }
        for bucket in age_order:
            bucket_df = ageing_df[ageing_df["Ageing"] == bucket]
            if bucket_df.empty:
                continue
            clients_in_bucket = {}
            for client in sorted(bucket_df["Client"].unique()):
                client_df = bucket_df[bucket_df["Client"] == client]
                clients_in_bucket[client] = {
                    "count": len(client_df),
                    "rows": client_df[ageing_cols].to_dict("records"),
                }
            ageing_list_data["buckets"][bucket] = clients_in_bucket

    return render_template(
        "dashboard.html",
        has_data=has_data,
        has_project=has_project,
        data_info=data_info,
        filter_options=filter_options,
        overview_charts=overview_charts,
        status_charts=status_charts,
        priority_charts=priority_charts,
        ageing_charts=ageing_charts,
        comparison_charts=comparison_charts,
        timeline_charts=timeline_charts,
        sla_charts=sla_charts,
        warranty_charts=warranty_charts,
        project_charts=project_charts,
        project_data=project_df.to_dict("records") if has_project else [],
        detail_table=detail_df.to_html(index=False, classes="table table-striped") if has_data and not detail_df.empty else "",
        ageing_list_data=ageing_list_data,
        now=datetime.now().strftime("%d-%m-%Y %H:%M"),
    )


if __name__ == "__main__":
    app.run(debug=True, port=8501)