import pandas as pd, os, glob, json, base64, io, numpy as np, re as _re

DATA_DIR = r'D:\Work\Opencode analysis'
HEADER_ROW = 1
OUTPUT = os.path.join(DATA_DIR, 'dashboard.html')

COLUMN_MAPPING = {
    "ticket no": "Ticket No", "ticket number": "Ticket No", "ticket_no": "Ticket No",
    "ticket id": "Ticket No", "task type": "Task Type", "task_type": "Task Type",
    "project": "Project", "company": "Company", "ticket title": "Ticket Title",
    "ticket_title": "Ticket Title", "ticket detail": "Ticket Detail",
    "ticket_detail": "Ticket Detail", "ticket category": "Ticket Category",
    "ticket_category": "Ticket Category", "priority": "Priority",
    "ticket created date": "Ticket Created Date", "ticket_created_date": "Ticket Created Date",
    "created date": "Ticket Created Date", "date created": "Ticket Created Date",
    "created_date": "Ticket Created Date", "ticket completed date": "Ticket Completed Date",
    "ticket_completed_date": "Ticket Completed Date", "completed date": "Ticket Completed Date",
    "ticket closed date": "Ticket Closed Date", "ticket_closed_date": "Ticket Closed Date",
    "closed date": "Ticket Closed Date", "ticket status": "Ticket Status",
    "ticket_status": "Ticket Status", "status": "Ticket Status",
    "sla dateline": "SLA Dateline", "sla late": "SLA Late", "days": "Days",
    "ageing": "Ageing", "no": None,
}

AGEING_STANDARDIZE = {
    "1-30 days": "1-30 Days", "1-30 Days": "1-30 Days",
    "30-60 Days": "31-60 Days", "30-60 days": "31-60 Days",
    "31-60 Days": "31-60 Days", "31-60 days": "31-60 Days",
    ">60 Days": "> 60 Days", ">60 days": "> 60 Days",
    "> 60 Days": "> 60 Days", "> 60 days": "> 60 Days",
}

AGE_ORDER = ["1-30 Days", "31-60 Days", "> 60 Days"]
PRIORITY_ORDER = ["High", "Medium", "Low"]
AGE_COLORS = {"1-30 Days": "#2ecc71", "31-60 Days": "#e67e22", "> 60 Days": "#e74c3c"}
PRIORITY_COLORS = {"1-30 Days": "#2ecc71", "31-60 Days": "#e67e22", "> 60 Days": "#e74c3c"}

# ---- LOAD DATA ----
files = sorted(glob.glob(os.path.join(DATA_DIR, '*.xlsx')))
files = [f for f in files if not os.path.basename(f).startswith('~')]
all_dfs = []
for fp in files:
    xl = pd.ExcelFile(fp, engine='openpyxl')
    for name in xl.sheet_names:
        try:
            h = pd.read_excel(fp, sheet_name=name, header=HEADER_ROW, engine='openpyxl', nrows=3)
            if 'Ticket No' in h.columns:
                df = pd.read_excel(fp, sheet_name=name, header=HEADER_ROW, engine='openpyxl')
                df = df.dropna(how='all')
                df['Client'] = name
                df['Source File'] = os.path.basename(fp)
                all_dfs.append(df)
        except:
            pass

df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
df.columns = [COLUMN_MAPPING.get(c.lower().strip(), c) for c in df.columns]
none_cols = [c for c in df.columns if c is None]
if none_cols:
    df.drop(columns=none_cols, inplace=True)
df = df.loc[:, ~df.columns.duplicated()]
df = df.dropna(subset=['Ticket No'])
df = df[df['Ticket No'].astype(str).str.match(r'^T/', na=False)]

# dtypes
for c in ['Ticket Created Date','Ticket Completed Date','Ticket Closed Date','SLA Dateline']:
    if c in df.columns:
        df[c] = pd.to_datetime(df[c], errors='coerce', dayfirst=True)

if 'Ticket Status' in df.columns:
    s = df['Ticket Status']
    if isinstance(s, pd.DataFrame): s = s.iloc[:, 0]
    df['Ticket Status'] = s.astype(str).str.strip().replace({'InProgress':'In Progress','Inprogress':'In Progress','nan':None})

if 'Priority' in df.columns:
    p = df['Priority']
    if isinstance(p, pd.DataFrame): p = p.iloc[:, 0]
    df['Priority'] = p.astype(str).str.strip().str.title().replace({'Nan':None,'None':None})

if 'Ageing' in df.columns:
    df['Ageing'] = df['Ageing'].astype(str).str.strip().replace({**AGEING_STANDARDIZE, 'nan': None})

if 'Days' in df.columns:
    df['Days'] = pd.to_numeric(df['Days'], errors='coerce')

if 'Ageing' not in df.columns and 'Days' in df.columns:
    def classify(d):
        if pd.isna(d): return None
        if d <= 30: return "1-30 Days"
        elif d <= 60: return "31-60 Days"
        else: return "> 60 Days"
    df['Ageing'] = df['Days'].apply(classify)

if 'SLA Breach' not in df.columns and 'SLA Late' in df.columns:
    df['SLA Breach'] = pd.to_numeric(df['SLA Late'], errors='coerce').fillna(0).astype(int)

# ---- LOAD CLIENT PROJECT DATA ----
project_df = pd.DataFrame()
for fp in files:
    xl = pd.ExcelFile(fp, engine='openpyxl')
    if 'Client Project' in xl.sheet_names:
        project_df = pd.read_excel(fp, sheet_name='Client Project', header=0, engine='openpyxl')
        project_df = project_df.dropna(how='all')
        if 'Client' in project_df.columns:
            project_df['Client'] = project_df['Client'].ffill()
        for c in ['Start date', 'Due date', 'Target Date']:
            if c in project_df.columns:
                project_df[c] = pd.to_datetime(project_df[c], errors='coerce', dayfirst=True)
        if 'Percentage' in project_df.columns:
            project_df['Percentage'] = pd.to_numeric(project_df['Percentage'].astype(str).str.replace('%', ''), errors='coerce')
        if 'Overall Progress Task (%)' in project_df.columns:
            project_df['Overall Progress Task (%)'] = pd.to_numeric(project_df['Overall Progress Task (%)'].astype(str).str.replace('%', ''), errors='coerce')
        break

# ---- CLIENT PROJECT AGGREGATIONS ----
project_total = len(project_df)
project_status_counts = project_df['Status Progress'].value_counts().to_dict() if 'Status Progress' in project_df.columns else {}
project_completed = sum(v for k,v in project_status_counts.items() if k and 'complet' in str(k).lower())
project_in_progress = sum(v for k,v in project_status_counts.items() if k and 'progress' in str(k).lower())
project_not_started = sum(v for k,v in project_status_counts.items() if k and 'not start' in str(k).lower())
project_clients = sorted(project_df['Client'].dropna().unique()) if 'Client' in project_df.columns and not project_df['Client'].dropna().empty else []
project_client_counts = project_df['Client'].dropna().value_counts().to_dict() if 'Client' in project_df.columns else {}

project_progress_data = []
if 'Title' in project_df.columns and 'Percentage' in project_df.columns:
    pv = project_df.dropna(subset=['Percentage']).sort_values('Percentage', ascending=True)
    for _, r in pv.iterrows():
        project_progress_data.append({
            'label': str(r.get('Title', '')) + ' - ' + str(r.get('Client', '')),
            'percentage': float(r.get('Percentage', 0)),
            'client': str(r.get('Client', '')),
            'status': str(r.get('Status Progress', '')),
        })

project_tasks = []
project_timeline = []
for _, r in project_df.iterrows():
    task = {}
    for c in ['Client', 'Title', 'Category', 'Progress', 'Priority', 'Assigned to', 'Status Progress', 'Percentage', 'Overall Progress Task (%)']:
        val = r.get(c, '')
        if c in ['Start date', 'Due date', 'Target Date'] and pd.notna(r.get(c)):
            val = r[c].strftime('%d/%m/%Y') if hasattr(r[c], 'strftime') else str(r[c])
        task[c] = str(val) if not pd.isna(val) else ''
    project_tasks.append(task)
    # Collect timeline data
    sd = r.get('Start date')
    dd = r.get('Due date')
    title = str(r.get('Title', ''))
    desc = str(r.get('Description', ''))
    desc = _re.sub(r'^\d+\.\s*', '', desc)
    desc = desc.split('\n')[0].strip()
    label = desc if desc.strip() else title
    if pd.notna(sd) and pd.notna(dd) and label.strip():
        project_timeline.append({
            'title': label,
            'client': str(r.get('Client', '')),
            'start': sd.strftime('%Y-%m-%d') if hasattr(sd, 'strftime') else str(sd),
            'due': dd.strftime('%Y-%m-%d') if hasattr(dd, 'strftime') else str(dd),
        })

# ---- COMPUTE AGGREGATIONS ----
total_records = len(df)

# Overview
status_counts = df['Ticket Status'].value_counts(dropna=False).to_dict()
pending = sum(v for k,v in status_counts.items() if k and 'pend' in str(k).lower())
in_progress = sum(v for k,v in status_counts.items() if k and ('progress' in str(k).lower() or 'inprog' in str(k).lower()))
completed = sum(v for k,v in status_counts.items() if k and ('complet' in str(k).lower() or 'closed' in str(k).lower() or 'close' in str(k).lower()))
sla_breach = int(df['SLA Breach'].sum()) if 'SLA Breach' in df.columns else 0
clients_list = sorted(df['Client'].unique())
client_counts = df['Client'].value_counts().to_dict()

# Ageing per client
ageing_data = {}
for c in clients_list:
    dc = df[(df['Client']==c) & (df['Ageing'].notna())]
    if not dc.empty:
        counts = dc['Ageing'].value_counts()
        ageing_data[c] = {k: int(counts.get(k, 0)) for k in AGE_ORDER}
        ageing_data[c]['Total'] = sum(ageing_data[c].values())

# Priority per client
priority_data = {}
pr_order = df['Priority'].dropna().unique().tolist() if 'Priority' in df.columns else []
for c in clients_list:
    dc = df[(df['Client']==c) & (df['Priority'].notna())] if 'Priority' in df.columns else df[df['Client']==c]
    if not dc.empty and 'Priority' in df.columns:
        counts = dc['Priority'].value_counts()
        priority_data[c] = {k: int(counts.get(k, 0)) for k in pr_order}

# Timeline
timeline_data = {}
if 'Ticket Created Date' in df.columns:
    t = df.dropna(subset=['Ticket Created Date']).copy()
    t['Bulan'] = t['Ticket Created Date'].dt.to_period('M').astype(str)
    timeline_data = t.groupby('Bulan').size().to_dict()

# SLA
sla_data = {}
if 'SLA Breach' in df.columns:
    for c in clients_list:
        dc = df[df['Client']==c]
        total_c = len(dc)
        breach_c = int(dc['SLA Breach'].sum())
        sla_data[c] = {'Total': total_c, 'Breach': breach_c, 'Compliant': total_c - breach_c,
                        'Rate': round((total_c - breach_c)/total_c*100, 1) if total_c > 0 else 0}

# Tickets detail
tickets_detail = []
for _, r in df.iterrows():
    tickets_detail.append({
        'Ticket No': str(r.get('Ticket No','')),
        'Client': str(r.get('Client','')),
        'Task Type': str(r.get('Task Type','')),
        'Priority': str(r.get('Priority','')),
        'Ticket Status': str(r.get('Ticket Status','')),
        'Ageing': str(r.get('Ageing','')),
        'Ticket Title': str(r.get('Ticket Title','')),
    })

# ---- BUILD DATA JSON ----
data_json = json.dumps({
    'total_records': total_records,
    'total_clients': len(clients_list),
    'status_counts': {str(k):int(v) for k,v in status_counts.items()},
    'pending': pending, 'in_progress': in_progress, 'completed': completed,
    'sla_breach': sla_breach,
    'clients': clients_list,
    'client_counts': {str(k):int(v) for k,v in client_counts.items()},
    'ageing_data': ageing_data,
    'age_order': AGE_ORDER,
    'age_colors': AGE_COLORS,
    'priority_order': [str(x) for x in pr_order],
    'priority_data': priority_data,
    'timeline_data': {str(k):int(v) for k,v in timeline_data.items()},
    'sla_data': sla_data,
    'tickets': tickets_detail,
    'status_options': sorted([str(s) for s in status_counts.keys() if s]),
    'priority_options': sorted([str(p) for p in pr_order]),
    'project_total': project_total,
    'project_status_counts': {str(k):int(v) for k,v in project_status_counts.items()},
    'project_completed': project_completed,
    'project_in_progress': project_in_progress,
    'project_not_started': project_not_started,
    'project_clients': project_clients,
    'project_client_counts': {str(k):int(v) for k,v in project_client_counts.items()},
    'project_progress': project_progress_data,
    'project_timeline': project_timeline,
    'project_tasks': project_tasks,
}, indent=2)

# ---- GENERATE HTML ----
def classify_age(d):
    try:
        d = float(d)
        if d <= 30: return "1-30 Days"
        elif d <= 60: return "31-60 Days"
        else: return "> 60 Days"
    except: return d

html = r'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Client Issues Dashboard</title>
<script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',sans-serif;background:#0e1117;color:#e0e0e0;padding:20px}
h1{color:#fff;font-size:24px;margin-bottom:4px}
.caption{color:#888;font-size:13px;margin-bottom:20px}
.controls{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;padding:16px;background:#1a1d23;border-radius:8px;align-items:end}
.control-group{display:flex;flex-direction:column;gap:4px}
.control-group label{font-size:12px;color:#888}
.control-group select,.control-group input{padding:6px 10px;border:1px solid #333;border-radius:4px;background:#262a30;color:#e0e0e0;font-size:13px;min-width:140px}
.control-group input[type=text]{min-width:180px}
.tabs{display:flex;gap:2px;margin-bottom:16px;flex-wrap:wrap}
.tab-btn{padding:8px 16px;background:#1a1d23;border:none;color:#888;cursor:pointer;border-radius:4px 4px 0 0;font-size:13px}
.tab-btn.active{background:#2ecc71;color:#000;font-weight:600}
.tab-pane{display:none}
.tab-pane.active{display:block}
.metrics{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.metric-card{flex:1;min-width:140px;padding:16px;background:#1a1d23;border-radius:8px;text-align:center}
.metric-card .val{font-size:28px;font-weight:700;color:#fff}
.metric-card .lbl{font-size:12px;color:#888;margin-top:4px}
.chart-row{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:16px}
.chart-box{flex:1;min-width:300px;background:#1a1d23;border-radius:8px;padding:12px}
.chart-box h3{font-size:14px;margin-bottom:8px;color:#ccc}
.chart-full{background:#1a1d23;border-radius:8px;padding:12px;margin-bottom:16px}
.chart-full h3{font-size:14px;margin-bottom:8px;color:#ccc}
table{width:100%;border-collapse:collapse;font-size:13px}
th,td{padding:6px 10px;text-align:left;border-bottom:1px solid #2a2d33}
th{background:#262a30;color:#aaa;font-weight:600;position:sticky;top:0}
td{color:#e0e0e0}
.client-section{margin-bottom:20px;background:#1a1d23;border-radius:8px;padding:12px}
.client-section h4{font-size:15px;margin-bottom:8px;color:#2ecc71}
.data-table-wrap{max-height:400px;overflow:auto;margin-top:8px}
#searchResult{margin-bottom:12px;padding:8px 12px;background:#1a1d23;border-radius:4px;font-size:13px;color:#888}
</style>
</head>
<body>

<h1>Client Issues Dashboard</h1>
<div class="caption">Static Dashboard &mdash; Data updated: ''' + pd.Timestamp.now().strftime('%d-%m-%Y %H:%M') + r'''</div>

<div class="controls" id="filters">
  <div class="control-group">
    <label>Start Date</label>
    <input type="date" id="dateStart">
  </div>
  <div class="control-group">
    <label>End Date</label>
    <input type="date" id="dateEnd">
  </div>
  <div class="control-group">
    <label>Client</label>
    <select id="filterClient" multiple size="4"></select>
  </div>
  <div class="control-group">
    <label>Priority</label>
    <select id="filterPriority" multiple size="4"></select>
  </div>
  <div class="control-group">
    <label>Status</label>
    <select id="filterStatus" multiple size="4"></select>
  </div>
  <div class="control-group">
    <label>Search</label>
    <input type="text" id="searchText" placeholder="Search tickets...">
  </div>
</div>

<div class="tabs" id="tabs"></div>
<div id="tabContent"></div>

<script>
var DATA = ''' + data_json + r''';

var ageOrder = DATA.age_order;
var ageColors = DATA.age_colors;

// ---------- filters ----------
var filters = {clients:[], priorities:[], statuses:[], search:'', dateStart:null, dateEnd:null};

function populateFilterOptions() {
  var sel = document.getElementById('filterClient');
  DATA.clients.forEach(function(c) {
    var o = document.createElement('option'); o.value=c; o.text=c; o.selected=true;
    sel.appendChild(o);
  });
  sel = document.getElementById('filterPriority');
  (DATA.priority_options||[]).forEach(function(p) {
    var o = document.createElement('option'); o.value=p; o.text=p; o.selected=true;
    sel.appendChild(o); });
  sel = document.getElementById('filterStatus');
  (DATA.status_options||[]).forEach(function(s) {
    var o = document.createElement('option'); o.value=s; o.text=s; o.selected=true;
    sel.appendChild(o); });
}
populateFilterOptions();

function getFilteredTickets() {
  var fs = Array.from(document.getElementById('filterClient').selectedOptions).map(function(o){return o.value;});
  var ps = Array.from(document.getElementById('filterPriority').selectedOptions).map(function(o){return o.value;});
  var ss = Array.from(document.getElementById('filterStatus').selectedOptions).map(function(o){return o.value;});
  var search = document.getElementById('searchText').value.toLowerCase();
  var ds = document.getElementById('dateStart').value;
  var de = document.getElementById('dateEnd').value;
  return DATA.tickets.filter(function(t) {
    if (fs.length && fs.indexOf(t.Client)===-1) return false;
    if (ps.length && ps.indexOf(t.Priority)===-1) return false;
    if (ss.length && ss.indexOf(t['Ticket Status'])===-1) return false;
    if (search) {
      var match = (t['Ticket Title']+' '+t['Ticket No']+' '+t.Client+' '+t['Ticket Status']).toLowerCase().indexOf(search)>-1;
      if (!match) return false;
    }
    // date filter not applied on ticket level (no date in tickets array)
    return true;
  });
}

// ---------- tab definitions ----------
var tabDefs = [];

// Tab 0: Overview
tabDefs.push({
  label: 'Overview',
  render: function(tickets) {
    var html = '<div class="metrics">';
    html += '<div class="metric-card"><div class="val">'+DATA.total_records+'</div><div class="lbl">Total Tickets</div></div>';
    html += '<div class="metric-card"><div class="val">'+DATA.total_clients+'</div><div class="lbl">Client</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#f39c12">'+DATA.pending+'</div><div class="lbl">Pending</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#3498db">'+DATA.in_progress+'</div><div class="lbl">In Progress</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#2ecc71">'+DATA.completed+'</div><div class="lbl">Completed</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#e74c3c">'+DATA.sla_breach+'</div><div class="lbl">SLA Breaches</div></div>';
    html += '</div>';

    html += '<div class="chart-row"><div class="chart-box"><h3>Status Distribution</h3><div id="chartStatus"></div></div>';
    html += '<div class="chart-box"><h3>Priority Distribution</h3><div id="chartPriority"></div></div></div>';

    html += '<div class="chart-box"><h3>Tickets by Client</h3><div id="chartClientDist"></div></div>';
    return html;
  },
  afterRender: function() {
    var labels=[], values=[], colors=[];
    var cmap={'Completed':'#2ecc71','Closed':'#3498db','Pending':'#f39c12','In Progress':'#e74c3c'};
    for (var k in DATA.status_counts) {
      labels.push(k); values.push(DATA.status_counts[k]);
      colors.push(cmap[k]||'#888'); }
    Plotly.newPlot('chartStatus', [{labels:labels, values:values, type:'pie', marker:{colors:colors},
      textinfo:'label+percent', textposition:'outside', hole:0.4}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, showlegend:false, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});

    var plabels=[], pvalues=[], pcolors=[]; var pmap={'High':'#e74c3c','Medium':'#3498db','Low':'#2ecc71'};
    if (DATA.priority_order) {
      DATA.priority_order.forEach(function(p) {
        var sum=0;
        for (var c in DATA.priority_data) sum += (DATA.priority_data[c][p]||0);
        if (sum>0) { plabels.push(p); pvalues.push(sum); pcolors.push(pmap[p]||'#888'); } }); }
    Plotly.newPlot('chartPriority', [{labels:plabels, values:pvalues, type:'pie', marker:{colors:pcolors},
      textinfo:'label+percent', textposition:'outside', hole:0.4}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, showlegend:false, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});

    var clabels=[], cvalues=[];
    for (var c in DATA.client_counts) { clabels.push(c); cvalues.push(DATA.client_counts[c]); }
    var plotlyColors = ['#636efa','#EF553B','#00cc96','#ab63fa','#FFA15A','#19d3f3','#FF6692','#B6E880','#FF97FF','#FECB52'];
    var ccolors = clabels.map(function(_,i){return plotlyColors[i%plotlyColors.length];});
    Plotly.newPlot('chartClientDist', [{x:clabels, y:cvalues, type:'bar', marker:{color:ccolors},
      text:cvalues, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:40}},
      {responsive:true, displayModeBar:false});
  }
});

// Tab 1: Status Breakdown
tabDefs.push({
  label: 'Status Breakdown',
  render: function() {
    var html = '<div class="chart-row">';
    html += '<div class="chart-box"><h3>Status Pie</h3><div id="chartStatusPie"></div></div>';
    html += '<div class="chart-box"><h3>Status Bar</h3><div id="chartStatusBar"></div></div></div>';
    html += '<div class="chart-full"><h3>Status by Client</h3><div id="chartStatusClient"></div></div>';
    return html;
  },
  afterRender: function() {
    var labels=[], values=[], colors=[], cmap={'Completed':'#2ecc71','Closed':'#3498db','Pending':'#f39c12','In Progress':'#e74c3c'};
    for (var k in DATA.status_counts) { labels.push(k); values.push(DATA.status_counts[k]); colors.push(cmap[k]||'#888'); }
    Plotly.newPlot('chartStatusPie', [{labels:labels, values:values, type:'pie', marker:{colors:colors},
      textinfo:'label+percent', textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});
    Plotly.newPlot('chartStatusBar', [{x:labels, y:values, type:'bar', marker:{color:colors},
      text:values, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:40,l:40,r:20}},
      {responsive:true, displayModeBar:false});

    // Status by client
    var ct = {}, cl = DATA.clients;
    cl.forEach(function(c) { ct[c]={};
      for (var k in DATA.status_counts) ct[c][k]=0; });
    DATA.tickets.forEach(function(t) {
      if (ct[t.Client] && ct[t.Client][t['Ticket Status']]!==undefined) ct[t.Client][t['Ticket Status']]++; });
    var traces = [];
    for (var k in DATA.status_counts) {
      traces.push({x:cl, y:cl.map(function(c){return ct[c][k];}), name:k, type:'bar',
        marker:{color:cmap[k]||'#888'}}); }
    Plotly.newPlot('chartStatusClient', traces,
      {barmode:'stack', template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
  }
});

// Tab 2: Priority Analysis
tabDefs.push({
  label: 'Priority Analysis',
  render: function() {
    var html = '<div class="chart-row">';
    html += '<div class="chart-box"><h3>Priority Pie</h3><div id="chartPriPie"></div></div>';
    html += '<div class="chart-box"><h3>Priority Bar</h3><div id="chartPriBar"></div></div></div>';
    html += '<div class="chart-full"><h3>Priority by Client</h3><div id="chartPriClient"></div></div>';
    return html;
  },
  afterRender: function() {
    var plabels=[], pvalues=[], pcolors=[], pmap={'High':'#e74c3c','Medium':'#3498db','Low':'#2ecc71'};
    DATA.priority_order.forEach(function(p) {
      var sum=0;
      for (var c in DATA.priority_data) sum += (DATA.priority_data[c][p]||0);
      if (sum>0) { plabels.push(p); pvalues.push(sum); pcolors.push(pmap[p]||'#888'); } });
    Plotly.newPlot('chartPriPie', [{labels:plabels, values:pvalues, type:'pie', marker:{colors:pcolors},
      textinfo:'label+percent', textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});
    Plotly.newPlot('chartPriBar', [{x:plabels, y:pvalues, type:'bar', marker:{color:pcolors},
      text:pvalues, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:40,l:40,r:20}},
      {responsive:true, displayModeBar:false});

    var traces = [];
    DATA.priority_order.forEach(function(p) {
      traces.push({x:DATA.clients, y:DATA.clients.map(function(c){return DATA.priority_data[c]?.(p)||0;}),
        name:p, type:'bar', marker:{color:pmap[p]||'#888'}}); });
    Plotly.newPlot('chartPriClient', traces,
      {barmode:'stack', template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
  }
});

// Tab 3: Ageing Analysis
tabDefs.push({
  label: 'Ageing Analysis',
  render: function() {
    var html = '<div class="metrics"><div class="metric-card"><div class="val">'+DATA.total_records+'</div><div class="lbl">Total Records</div></div></div>';
    html += '<div id="ageingClients"></div>';
    return html;
  },
  afterRender: function() {
    var container = document.getElementById('ageingClients');
    var html = '';
    for (var ci=0; ci<DATA.clients.length; ci++) {
      var c = DATA.clients[ci];
      var cd = DATA.ageing_data[c];
      if (!cd || !cd.Total) continue;
      html += '<div class="client-section"><h4>'+c+' \u2014 '+cd.Total+' records</h4>';
      html += '<div class="chart-row"><div class="chart-box" style="min-width:200px;flex:0.4">';
      html += '<table><tr><th>Age Group</th><th>Count</th></tr>';
      ageOrder.forEach(function(a) {
        html += '<tr><td>'+a+'</td><td>'+(cd[a]||0)+'</td></tr>';
      });
      html += '<tr style="font-weight:700;border-top:2px solid #444"><td>Total</td><td>'+cd.Total+'</td></tr>';
      html += '</table></div>';
      html += '<div class="chart-box" style="flex:1"><div id="ageChart'+ci+'"></div></div></div></div>';
    }
    container.innerHTML = html;
    for (var ci=0; ci<DATA.clients.length; ci++) {
      var c = DATA.clients[ci];
      var cd = DATA.ageing_data[c];
      if (!cd || !cd.Total) continue;
      var aLabels = [], aValues = [], aColors = [];
      ageOrder.forEach(function(a) {
        if (cd[a]) { aLabels.push(a); aValues.push(cd[a]); aColors.push(ageColors[a]); }
      });
      Plotly.newPlot('ageChart'+ci, [{labels:aLabels, values:aValues, type:'pie', marker:{colors:aColors},
        textinfo:'label+percent', textposition:'outside', hole:0.4}],
        {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
         font:{color:'#e0e0e0'}, showlegend:false, margin:{t:0,b:0,l:0,r:0}},
        {responsive:true, displayModeBar:false});
    }
  }
});

// Tab 4: Client Comparison
tabDefs.push({
  label: 'Client Comparison',
  render: function() {
    var html = '<div class="chart-row">';
    html += '<div class="chart-box"><h3>Tickets by Client</h3><div id="compClientPie"></div></div>';
    html += '<div class="chart-box"><h3>Ticket Distribution</h3><div id="compClientBar"></div></div></div>';
    html += '<div class="chart-full"><h3>Client Summary</h3><div class="data-table-wrap"><table id="compClientTable"><tr><th>Client</th><th>Total Tickets</th>';
    html += '<th>Pending</th><th>In Progress</th><th>Completed</th><th>SLA Breaches</th></tr></table></div></div>';
    return html;
  },
  afterRender: function() {
    var clabels=[], cvalues=[];
    for (var c in DATA.client_counts) { clabels.push(c); cvalues.push(DATA.client_counts[c]); }
    Plotly.newPlot('compClientPie', [{labels:clabels, values:cvalues, type:'pie',
      textinfo:'label+percent', textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});
    Plotly.newPlot('compClientBar', [{x:clabels, y:cvalues, type:'bar', marker:{color:'#2ecc71'},
      text:cvalues, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});

    // Client summary table
    var tbody = document.getElementById('compClientTable');
    DATA.clients.forEach(function(c) {
      var total = DATA.client_counts[c]||0;
      var s = DATA.sla_data[c]||{};
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>'+c+'</td><td>'+total+'</td><td>'+DATA.pending+'</td><td>'+DATA.in_progress+'</td><td>'+DATA.completed+'</td><td>'+(s.Breach||0)+'</td>';
      tbody.appendChild(tr);
    });
  }
});

// Tab 5: Timeline
tabDefs.push({
  label: 'Timeline',
  render: function() {
    return '<div class="chart-full"><h3>Ticket Flow Over Time</h3><div id="chartTimeline"></div></div>';
  },
  afterRender: function() {
    var dates = Object.keys(DATA.timeline_data).sort();
    var counts = dates.map(function(d){return DATA.timeline_data[d];});
    Plotly.newPlot('chartTimeline', [{x:dates, y:counts, type:'scatter', mode:'lines+markers',
      line:{color:'#3498db', width:2}, marker:{color:'#3498db', size:6},
      fill:'tozeroy', fillcolor:'rgba(52,152,219,0.15)'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
  }
});

// Tab 6: SLA Compliance
tabDefs.push({
  label: 'SLA Compliance',
  render: function() {
    var html = '<div class="metrics">';
    var totalCompliant = 0, totalBreach = 0;
    for (var c in DATA.sla_data) { totalCompliant += DATA.sla_data[c].Compliant; totalBreach += DATA.sla_data[c].Breach; }
    var overallRate = DATA.total_records > 0 ? Math.round((totalCompliant/DATA.total_records)*100) : 0;
    html += '<div class="metric-card"><div class="val" style="color:'+(overallRate>=80?'#2ecc71':'#e74c3c')+'">'+overallRate+'%</div><div class="lbl">Compliance Rate</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#2ecc71">'+totalCompliant+'</div><div class="lbl">SLA Compliant</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#e74c3c">'+totalBreach+'</div><div class="lbl">SLA Breaches</div></div>';
    html += '</div>';
    html += '<div class="chart-row"><div class="chart-box"><h3>SLA by Client</h3><div id="chartSLA"></div></div>';
    html += '<div class="chart-box"><h3>SLA Compliance Rate</h3><div id="chartSLARate"></div></div></div>';
    return html;
  },
  afterRender: function() {
    var cl = DATA.clients.filter(function(c){return DATA.sla_data[c];});
    var breachVals = cl.map(function(c){return DATA.sla_data[c].Breach;});
    var complVals = cl.map(function(c){return DATA.sla_data[c].Compliant;});
    Plotly.newPlot('chartSLA', [
      {x:cl, y:complVals, name:'Compliant', type:'bar', marker:{color:'#2ecc71'}},
      {x:cl, y:breachVals, name:'Breach', type:'bar', marker:{color:'#e74c3c'}}
    ], {barmode:'stack', template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
        font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});

    var rates = cl.map(function(c){return DATA.sla_data[c].Rate;});
    Plotly.newPlot('chartSLARate', [{x:cl, y:rates, type:'bar', marker:{
      color:rates.map(function(v){
        if (v>=80) return '#2ecc71'; if (v>=50) return '#f39c12'; return '#e74c3c';
      })}, text:rates.map(function(v){return v+'%';}), textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, yaxis:{range:[0,100]}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
  }
});

// Tab 7: Ticket Details
tabDefs.push({
  label: 'Ticket Details',
  render: function() {
    var html = '<div id="searchResult">'+DATA.total_records+' records</div>';
    html += '<div class="data-table-wrap"><table id="ticketTable"><tr><th>Ticket</th><th>Client</th><th>Task Type</th><th>Priority</th><th>Status</th><th>Ageing</th><th>Title</th></tr></table></div>';
    return html;
  },
  afterRender: function() {
    renderTickets(DATA.tickets);
  }
});

// Tab 8: Ageing List
tabDefs.push({
  label: 'Ageing List',
  render: function() {
    var ageOrder = DATA.age_order;
    var html = '<div id="ageingList"></div>';
    return html;
  },
  afterRender: function() {
    var container = document.getElementById('ageingList');
    var ageOrder = DATA.age_order;
    var tickets = DATA.tickets.filter(function(t){return t.Ageing && t.Ageing!=='nan' && t.Ageing!=='None' && t.Ageing!=='';});
    container.innerHTML = '<div style="margin-bottom:12px;padding:8px 12px;background:#1a1d23;border-radius:4px;font-size:13px;color:#888">Total tickets with ageing: '+tickets.length+'</div>';
    ageOrder.forEach(function(bucket) {
      var bt = tickets.filter(function(t){return t.Ageing===bucket;});
      if (!bt.length) return;
      var section = document.createElement('div');
      section.style.cssText = 'margin-bottom:16px;background:#1a1d23;border-radius:8px;padding:12px';
      var header = document.createElement('h3');
      header.style.cssText = 'font-size:15px;color:#e0e0e0;margin-bottom:10px';
      header.textContent = bucket + ' \u2014 ' + bt.length + ' tickets';
      section.appendChild(header);
      var clients = {};
      bt.forEach(function(t){clients[t.Client]=(clients[t.Client]||0)+1;});
      var sortedClients = Object.keys(clients).sort();
      sortedClients.forEach(function(c) {
        var detail = document.createElement('details');
        detail.style.cssText = 'margin-bottom:6px';
        var summary = document.createElement('summary');
        summary.style.cssText = 'cursor:pointer;padding:6px 10px;background:#262a30;border-radius:4px;font-size:13px;color:#2ecc71';
        summary.textContent = bucket + ' / ' + c + ' (' + clients[c] + ' tickets)';
        detail.appendChild(summary);
        var table = document.createElement('table');
        table.style.cssText = 'width:100%;margin-top:6px;font-size:12px';
        table.innerHTML = '<tr><th>Ticket No</th><th>Priority</th><th>Status</th><th>Title</th></tr>';
        bt.filter(function(t){return t.Client===c;}).forEach(function(t) {
          var tr = document.createElement('tr');
          tr.innerHTML = '<td>'+t['Ticket No']+'</td><td>'+t.Priority+'</td><td>'+t['Ticket Status']+'</td><td>'+t['Ticket Title']+'</td>';
          table.appendChild(tr);
        });
        detail.appendChild(table);
        section.appendChild(detail);
      });
      container.appendChild(section);
    });
  }
});

// Tab 8: Client Warranty
tabDefs.push({
  label: 'Client Warranty',
  render: function() {
    var wtickets = DATA.tickets.filter(function(t){return t.Client==='Client Warranty';});
    if (!wtickets.length) return '<div style="padding:16px;color:#888">No Warranty data found</div>';
    var wTotal = wtickets.length;
    var wCompleted = wtickets.filter(function(t){return ['Completed','Closed'].indexOf(t['Ticket Status'])>=0;}).length;
    var wPending = wtickets.filter(function(t){return t['Ticket Status']==='Pending';}).length;
    var wInProg = wtickets.filter(function(t){return t['Ticket Status']==='In Progress';}).length;
    var html = '<div class="metrics">';
    html += '<div class="metric-card"><div class="val">'+wTotal+'</div><div class="lbl">Total Warranty Tickets</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#2ecc71">'+wCompleted+'</div><div class="lbl">Completed</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#f39c12">'+wPending+'</div><div class="lbl">Pending</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#e74c3c">'+wInProg+'</div><div class="lbl">In Progress</div></div>';
    html += '</div>';
    html += '<div class="chart-row"><div class="chart-box"><h3>Status Warranty</h3><div id="chartWStatus"></div></div>';
    html += '<div class="chart-box"><h3>Task Type</h3><div id="chartWTaskType"></div></div></div>';
    html += '<div class="chart-full"><h3>Warranty Ticket Details</h3><div class="data-table-wrap"><table id="warrantyTable"><tr><th>Ticket</th><th>Task Type</th><th>Project</th><th>Priority</th><th>Status</th><th>Title</th></tr></table></div></div>';
    return html;
  },
  afterRender: function() {
    var wtickets = DATA.tickets.filter(function(t){return t.Client==='Client Warranty';});
    var labels=[], values=[], colors=[], cmap={'Completed':'#2ecc71','Closed':'#3498db','Pending':'#f39c12','In Progress':'#e74c3c'};
    var scounts={};
    wtickets.forEach(function(t){scounts[t['Ticket Status']]=(scounts[t['Ticket Status']]||0)+1;});
    for (var k in scounts) { labels.push(k); values.push(scounts[k]); colors.push(cmap[k]||'#888'); }
    Plotly.newPlot('chartWStatus', [{labels:labels, values:values, type:'pie', marker:{colors:colors},
      textinfo:'label+percent', textposition:'outside', hole:0.4}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});
    var tlabels=[], tvalues=[];
    var tcounts={};
    wtickets.forEach(function(t){tcounts[t['Task Type']]=(tcounts[t['Task Type']]||0)+1;});
    for (var k in tcounts) { tlabels.push(k); tvalues.push(tcounts[k]); }
    Plotly.newPlot('chartWTaskType', [{x:tlabels, y:tvalues, type:'bar', marker:{color:'#3498db'},
      text:tvalues, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, xaxis:{tickangle:-45}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
    var tbody = document.getElementById('warrantyTable');
    wtickets.forEach(function(t) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>'+t['Ticket No']+'</td><td>'+t['Task Type']+'</td><td>'+(t.Project||'')+'</td><td>'+t.Priority+'</td><td>'+t['Ticket Status']+'</td><td>'+t['Ticket Title']+'</td>';
      tbody.appendChild(tr);
    });
  }
});

// Tab 9: Client Project
tabDefs.push({
  label: 'Client Project',
  render: function() {
    if (!DATA.project_total) return '<div style="padding:16px;color:#888">No Project data found</div>';
    var html = '<div class="metrics">';
    html += '<div class="metric-card"><div class="val">'+DATA.project_total+'</div><div class="lbl">Total Projects</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#2ecc71">'+DATA.project_completed+'</div><div class="lbl">Completed</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#3498db">'+DATA.project_in_progress+'</div><div class="lbl">In Progress</div></div>';
    html += '<div class="metric-card"><div class="val" style="color:#95a5a6">'+DATA.project_not_started+'</div><div class="lbl">Not Started</div></div>';
    html += '</div>';
    html += '<div class="chart-row"><div class="chart-box"><h3>Projects by Client</h3><div id="chartProjClient"></div></div>';
    html += '<div class="chart-box"><h3>Status Progress</h3><div id="chartProjStatus"></div></div></div>';
    if (DATA.project_timeline && DATA.project_timeline.length) {
      html += '<div class="chart-full"><h3>PROJECT DEVELOPMENT TIMELINE</h3><div id="chartProjTimelineContainer"></div></div>';
    }
    html += '<div class="chart-full"><h3>Project Details</h3><div class="data-table-wrap"><table id="projectTable"><tr><th>Client</th><th>Title</th><th>Category</th><th>Priority</th><th>Assigned To</th><th>Status</th><th>Percentage</th></tr></table></div></div>';
    return html;
  },
  afterRender: function() {
    var clabels=[], cvalues=[];
    for (var c in DATA.project_client_counts) { clabels.push(c); cvalues.push(DATA.project_client_counts[c]); }
    Plotly.newPlot('chartProjClient', [{x:clabels, y:cvalues, type:'bar', marker:{color:'#2ecc71'},
      text:cvalues, textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:20,b:60,l:40,r:20}},
      {responsive:true, displayModeBar:false});
    var plabels=[], pvalues=[];
    for (var k in DATA.project_status_counts) { plabels.push(k); pvalues.push(DATA.project_status_counts[k]); }
    Plotly.newPlot('chartProjStatus', [{labels:plabels, values:pvalues, type:'pie', hole:0.4,
      textinfo:'label+percent', textposition:'outside'}],
      {template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
       font:{color:'#e0e0e0'}, margin:{t:0,b:0,l:0,r:0}},
      {responsive:true, displayModeBar:false});
    if (DATA.project_timeline && DATA.project_timeline.length) {
      var tdata = DATA.project_timeline;
      var uniqueClients = [...new Set(tdata.map(function(d){return d.client;}))];
      var plotlyColors = ['#636efa','#EF553B','#00cc96','#ab63fa','#FFA15A','#19d3f3','#FF6692','#B6E880','#FF97FF','#FECB52'];
      var colorMap = {};
      uniqueClients.forEach(function(c, i){colorMap[c]=plotlyColors[i%plotlyColors.length];});
      var container = document.getElementById('chartProjTimelineContainer');
      container.innerHTML = '';
      uniqueClients.forEach(function(client, ci) {
        var ct = tdata.filter(function(d){return d.client===client;});
        if (!ct.length) return;
        var section = document.createElement('div');
        section.className = 'client-section';
        section.innerHTML = '<h4 style="color:#2ecc71;margin-bottom:8px">'+client+'</h4><div id="projTimeline'+ci+'"></div>';
        container.appendChild(section);
        var traces = [{
          x: ct.map(function(d){return d.due;}),
          y: ct.map(function(d){return d.title;}),
          base: ct.map(function(d){return d.start;}),
          type: 'bar', orientation: 'h',
          marker: {color: colorMap[client]},
          width: 0.6, name: client,
        }];
        var layout = {
          template:'plotly_dark', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
          font:{color:'#e0e0e0'}, yaxis:{autorange:'reversed', title:''},
          margin:{t:20,b:60,l:160,r:40}, height:Math.max(200, 30*ct.length),
          xaxis:{title:'Date'}, showlegend:false,
        };
        Plotly.newPlot('projTimeline'+ci, traces, layout,
          {responsive:true, displayModeBar:false});
      });
    }
    var tbody = document.getElementById('projectTable');
    DATA.project_tasks.forEach(function(t) {
      var tr = document.createElement('tr');
      tr.innerHTML = '<td>'+t.Client+'</td><td>'+t.Title+'</td><td>'+t.Category+'</td><td>'+t.Priority+'</td><td>'+t['Assigned to']+'</td><td>'+t['Status Progress']+'</td><td>'+(t.Percentage||'')+'</td>';
      tbody.appendChild(tr);
    });
  }
});

// ---------- render ----------
function renderTickets(tickets) {
  var tbody = document.getElementById('ticketTable');
  while (tbody.rows.length>1) tbody.deleteRow(1);
  tickets.forEach(function(t) {
    var tr = document.createElement('tr');
    tr.innerHTML = '<td>'+t['Ticket No']+'</td><td>'+t.Client+'</td><td>'+t['Task Type']+'</td><td>'+t.Priority+'</td><td>'+t['Ticket Status']+'</td><td>'+t.Ageing+'</td><td>'+t['Ticket Title']+'</td>';
    tbody.appendChild(tr);
  });
  document.getElementById('searchResult').textContent = tickets.length+' records displayed';
}

function renderTab(idx) {
  var tabs = document.querySelectorAll('.tab-btn');
  tabs.forEach(function(b,i){b.classList.toggle('active',i===idx);});
  var panes = document.querySelectorAll('.tab-pane');
  panes.forEach(function(p,i){p.classList.toggle('active',i===idx);});
  if (tabDefs[idx].afterRender) {
    setTimeout(function(){tabDefs[idx].afterRender();}, 50);
  }
}

function initTabs() {
  var tabBar = document.getElementById('tabs');
  var content = document.getElementById('tabContent');
  tabDefs.forEach(function(def, idx) {
    var btn = document.createElement('button');
    btn.className = 'tab-btn'+(idx===0?' active':'');
    btn.textContent = def.label;
    btn.onclick = function(){renderTab(idx);};
    tabBar.appendChild(btn);
    var pane = document.createElement('div');
    pane.className = 'tab-pane'+(idx===0?' active':'');
    pane.innerHTML = def.render();
    content.appendChild(pane);
  });
  renderTab(0);
}

initTabs();
</script>
</body>
</html>'''

with open(OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'Dashboard static HTML generated: {OUTPUT}')
print(f'Saiz: {os.path.getsize(OUTPUT):,} bytes')