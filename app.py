from zoneinfo import ZoneInfo
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date
import re
import requests
import io

# -------------------------------------------------------------
# 1. CẤU HÌNH TRANG & TÙY BIẾN CSS
# -------------------------------------------------------------
st.set_page_config(
    page_title="Master Plan SG — Cảnh báo Delay & Theo dõi MQL",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    .reportview-container, .main .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1550px; }
    .header-box { background-color: var(--secondary-background-color); color: var(--text-color); padding: 16px 22px; border-radius: 8px; margin-bottom: 14px; border: 1px solid rgba(128, 128, 128, 0.2); }
    .header-box h2 { margin: 0; font-size: 21px; font-weight: 700; color: var(--text-color) !important; }
    .header-box p { margin: 4px 0 0 0; font-size: 12.5px; opacity: 0.8; }
    .kpi-card { border-radius: 8px; padding: 14px 16px; box-shadow: 0 3px 6px rgba(0,0,0,0.15); margin-bottom: 12px; }
    .kpi-card .kpi-title { font-size: 11px; text-transform: uppercase; font-weight: 800; letter-spacing: 0.5px; opacity: 0.9; }
    .kpi-card .kpi-value { font-size: 27px; font-weight: 800; margin: 3px 0; }
    .kpi-card .kpi-sub { font-size: 11px; opacity: 0.95; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .kpi-red { background-color: #ef4444; } .kpi-red * { color: #ffffff !important; }
    .kpi-orange { background-color: #f97316; } .kpi-orange * { color: #ffffff !important; }
    .kpi-amber { background-color: #facc15; } .kpi-amber * { color: #111827 !important; }
    .kpi-blue { background-color: #3b82f6; } .kpi-blue * { color: #ffffff !important; }
    .kpi-green { background-color: #22c55e; } .kpi-green * { color: #ffffff !important; }
    .kpi-black { background-color: #9333EA; border: 1px solid rgba(255,255,255,0.1); } .kpi-black * { color: #ffffff !important; }
    .m-card { background-color: var(--background-color); border: 1px solid rgba(128, 128, 128, 0.2); border-radius: 7px; padding: 10px 12px; border-left: 5px solid #22c55e; margin-bottom: 10px; min-height: 100px; }
    .m-card.crit { border-left-color: #ef4444; background-color: rgba(239, 68, 68, 0.03); }
    .m-card.warn { border-left-color: #f97316; background-color: rgba(249, 115, 22, 0.03); }
    .m-card.idle { border-left-color: #94a3b8; background-color: var(--secondary-background-color); }
    .mach-badge { font-size: 14px !important; font-weight: 800 !important; color: var(--text-color) !important; background-color: var(--secondary-background-color) !important; padding: 2px 9px !important; border-radius: 5px !important; display: inline-block !important; border: 1px solid rgba(128, 128, 128, 0.2) !important; }
    .badge-red { background: rgba(239, 68, 68, 0.15); color: #ef4444 !important; padding: 2px 7px; border-radius: 10px; font-size: 11px; font-weight: 700; border: 1px solid rgba(239, 68, 68, 0.3); }
    .badge-org { background: rgba(249, 115, 22, 0.15); color: #f97316 !important; padding: 2px 7px; border-radius: 10px; font-size: 11px; font-weight: 700; border: 1px solid rgba(249, 115, 22, 0.3); }
    .badge-grn { background: rgba(34, 197, 94, 0.15); color: #22c55e !important; padding: 2px 7px; border-radius: 10px; font-size: 11px; font-weight: 700; border: 1px solid rgba(34, 197, 94, 0.3); }
    .badge-gry { background: rgba(148, 163, 184, 0.15); color: #94a3b8 !important; padding: 2px 7px; border-radius: 10px; font-size: 11px; font-weight: 700; border: 1px solid rgba(148, 163, 184, 0.3); }
</style>
""", unsafe_allow_html=True)

# -------------------------------------------------------------

@st.cache_data(show_spinner=False, ttl=600) 
def fetch_excel_from_url(gsheet_url_param):
    try:
        url = gsheet_url_param.strip()
        download_url = None
        
        if "/e/" in url:
            base = url.split("?")[0]
            if base.endswith("/pubhtml"):
                base = base.replace("/pubhtml", "/pub")
            download_url = base + "?output=xlsx"
        elif "/d/" in url:
            file_id = url.split("/d/")[1].split("/")[0]
            download_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=xlsx"
        else:
            return None, "Link không đúng định dạng Google Sheets."

        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
            'Accept-Language': 'vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7',
            'Connection': 'keep-alive'
        }
        
        session = requests.Session()
        response = session.get(download_url, headers=headers, stream=True, allow_redirects=True, timeout=(10, 30))
        
        token = next((v for k, v in response.cookies.items() if k.startswith('download_warning')), None)
        if token:
            response = session.get(download_url + f"&confirm={token}", headers=headers, stream=True, allow_redirects=True, timeout=(10, 30))
            
        if response.status_code == 200:
            content_type = response.headers.get("Content-Type", "")
            if "text/html" in content_type:
                return None, "Lỗi: Google chặn quyền (File chưa được Share hoặc Công bố công khai)."
            else:
                return response.content, None
        elif response.status_code == 404:
            return None, "Lỗi HTTP 404: Không tìm thấy file. Link có thể đã bị xóa hoặc sai mã."
        else:
            return None, f"Lỗi HTTP {response.status_code}: Không thể tải file từ máy chủ Google."
            
    except Exception as e:
        return None, f"Lỗi mạng: {e}"

FACTORY_MACHINES = [
    "MB1", "MA2", "MB3", "MA6", "MB6", "MA7", "MB7", 
    "MA1", "MB2", "MA3", "MA5", "MB5", "MA8", "MB8", 
    "MA4", "MB4", "TC2", "TD2", "TC3", "TC1", "TD1", 
    "GD1", "GD2", "WC"
]

MACHINE_DETAILS = {
    "MA1": {"a": "TIÊN", "b": "VŨ", "note": ""}, "MA2": {"a": "HẢI", "b": "TRẠNG", "note": ""},
    "MA3": {"a": "B NAM", "b": "MẪN", "note": ""}, "MA4": {"a": "B NAM", "b": "MẪN", "note": ""},
    "MA5": {"a": "NAM", "b": "LUÂN", "note": ""}, "MA6": {"a": "NAM", "b": "LUÂN", "note": ""},
    "MA7": {"a": "KỲ", "b": "VINH", "note": ""}, "MA8": {"a": "HIỀN", "b": "MINH", "note": ""},
    "MB1": {"a": "TIÊN", "b": "VŨ", "note": ""}, "MB2": {"a": "HẢI", "b": "TRẠNG", "note": ""},
    "MB3": {"a": "THÀNH", "b": "THÔNG", "note": ""}, "MB4": {"a": "THÀNH", "b": "THÔNG", "note": ""},
    "MB5": {"a": "NAM", "b": "LUÂN", "note": ""}, "MB6": {"a": "NAM", "b": "LUÂN", "note": ""},
    "MB7": {"a": "KỲ", "b": "VINH", "note": ""}, "MB8": {"a": "HIỀN", "b": "MINH", "note": ""},
    "TC1": {"a": "HOÀNG", "b": "NAM (T)", "note": ""}, "TC2": {"a": "HOÀNG", "b": "NAM (T)", "note": ""},
    "TC3": {"a": "ĐỆ", "b": "KIỆT", "note": ""}, "TD1": {"a": "ĐẠT", "b": "PHÚ", "note": ""},
    "TD2": {"a": "ĐẠT", "b": "PHÚ", "note": ""}, "GD1": {"a": "HỮU", "b": "—", "note": ""},
    "GD2": {"a": "HỮU", "b": "—", "note": ""}, "WC": {"a": "—", "b": "—", "note": ""}
}

def parse_gantt_dates(val, ref_year):
    # CHỐT CHẶN 1: Ép kiểu nếu vô tình nhận phải Series do cột trùng lặp
    if isinstance(val, pd.Series): 
        val = val.iloc[0]
        
    if pd.isna(val): return pd.NaT, pd.NaT
    s = str(val).strip()
    if not s: return pd.NaT, pd.NaT
    if isinstance(val, (datetime, date)): return pd.to_datetime(val), pd.to_datetime(val)
    try:
        if '-' in s:
            p1, p2 = s.split('-')[:2]
            p1, p2 = p1.strip(), p2.strip()
            m_p2 = re.search(r'(\d{1,2})/(\d{1,2})', p2)
            if m_p2: d2, m2 = int(m_p2.group(1)), int(m_p2.group(2))
            else: d2, m2 = int(re.sub(r'\D', '', p2)), datetime.now().month
            m_p1 = re.search(r'(\d{1,2})/(\d{1,2})', p1)
            if m_p1: d1, m1 = int(m_p1.group(1)), int(m_p1.group(2))
            else: d1, m1 = int(re.sub(r'\D', '', p1)), m2 
            start_dt = pd.Timestamp(year=ref_year, month=m1, day=d1)
            end_dt = pd.Timestamp(year=ref_year, month=m2, day=d2)
            if end_dt < start_dt: end_dt = end_dt.replace(year=ref_year + 1)
            return start_dt, end_dt
        else: 
            m_s = re.search(r'(\d{1,2})/(\d{1,2})', s)
            if m_s:
                dt = pd.Timestamp(year=ref_year, month=int(m_s.group(2)), day=int(m_s.group(1)))
                return dt, dt
            elif re.match(r'^\d{4}-\d{2}-\d{2}', s):
                dt = pd.to_datetime(s[:10], errors='coerce', dayfirst=True)
                return dt, dt
    except: pass
    return pd.NaT, pd.NaT

@st.cache_data(show_spinner=False)
def load_and_preprocess_data(file_source_bytes, ref_dt_str):
    ref_dt_val = pd.to_datetime(ref_dt_str)
    configs = [('1. MP MILLING', 3, 'Phay'), ('2. MP TURNING', 5, 'Tiện'), ('3. MP GRINDING', 3, 'Mài')]
    dfs = []
    free_dates = {}
    dynamic_notes = {}
    
    try: xls = pd.ExcelFile(io.BytesIO(file_source_bytes))
    except: return pd.DataFrame(), {}, {}
    
    if 'MACHINE&ABILITY' in xls.sheet_names:
        try:
            df_mach = pd.read_excel(xls, sheet_name='MACHINE&ABILITY', header=None)
            for idx in range(6, len(df_mach)): 
                if df_mach.shape[1] > 1: 
                    m_name = str(df_mach.iloc[idx, 1]).strip()
                    if m_name and m_name.lower() != 'nan':
                        note_str = ""
                        try:
                            val = df_mach.iloc[idx, 16]
                            if pd.notna(val):
                                note_str = str(val).strip()
                                if note_str.lower() == 'nan': note_str = ""
                        except IndexError:
                            pass 
                        dynamic_notes[m_name] = note_str
        except Exception:
            pass 
            
    for sname, h_idx, ws in configs:
        if sname in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sname, header=h_idx)
            df.columns = [str(c).replace('\n', ' ').strip() for c in df.columns]
            
            # CHỐT CHẶN 2 (QUAN TRỌNG NHẤT): Loại bỏ cột trùng lặp để xử lý tận gốc lỗi Series Truth Value
            new_cols = []
            seen = {}
            for c in df.columns:
                if c not in seen:
                    seen[c] = 1
                    new_cols.append(c)
                else:
                    new_cols.append(f"{c}_dup{seen[c]}")
                    seen[c] += 1
            df.columns = new_cols
            # -------------------------------------------------------------
            
            date_col = next((col for col in df.columns if 'start & end' in col.lower()), None)
            if date_col and 'Machine Name' in df.columns:
                df['Machine Name'] = df['Machine Name'].ffill() 
                df['End_Time_Raw'] = df.groupby('Machine Name')[date_col].shift(-1)
                
                for m in df['Machine Name'].dropna().unique():
                    m_code = str(m).strip()
                    m_df = df[df['Machine Name'] == m]
                    v_dates = m_df[m_df[date_col].notna() & (m_df[date_col].astype(str).str.strip() != '')]
                    if not v_dates.empty:
                        free_dates[m_code] = v_dates.iloc[-1][date_col]
            
            if 'Job Order' in df.columns:
                df = df[df['Job Order'].notna()].copy()
                df['Job Order'] = df['Job Order'].astype(str).str.strip()
                df = df[~df['Job Order'].str.contains(r'^\d+$|INSERT|CHỜ SẮP|MÀI|GCN|TOTAL|TỔNG', case=False, na=False)]
                df['Workshop'] = ws
                dfs.append(df)
                
    if not dfs: return pd.DataFrame(), {}, dynamic_notes
    res = pd.concat(dfs, ignore_index=True)
    
    date_col_final = next((col for col in res.columns if 'start & end' in col.lower()), None)
    if date_col_final:
        res['Real_Start'] = pd.to_datetime(res[date_col_final], errors='coerce', dayfirst=True)
    
    res['Customer'] = res['Job Order'].str.extract(r'^([A-Za-z]+)')[0].fillna('—').str.upper()
    res['Machine Name'] = res['Machine Name'].fillna('—').astype(str).str.strip()
    res['PO'] = res['PO'].fillna('—').astype(str).str.strip()
    res['Status'] = res['Status'].fillna('Waiting').astype(str).str.strip()
    
    res['D2'] = pd.to_datetime(res.get('Deadline  2nd'), errors='coerce', dayfirst=True)
    res['D1'] = pd.to_datetime(res.get('Deadline  1st'), errors='coerce', dayfirst=True)
    res['Hạn áp dụng'] = res['D2'].combine_first(res['D1'])
    
    res['Còn (ngày)'] = (res['Hạn áp dụng'] - ref_dt_val).dt.days
    
    def classify_risk(row):
        # CHỐT CHẶN 3: Xử lý an toàn khi DataFrame có nguy cơ sinh ra Series trong df.apply
        status = row.get('Status', '')
        if isinstance(status, pd.Series): status = status.iloc[0]
        if str(status).strip().lower() in ['done', 'services']: return 'Đã xong / GC ngoài'
        
        d = row.get('Còn (ngày)')
        if isinstance(d, pd.Series): d = d.iloc[0]
        
        if pd.isna(d): return 'Thiếu hạn'
        if d < 0: return 'Quá hạn'
        if d < 7: return '< 7 ngày'
        if d <= 14: return '7–14 ngày'
        if d <= 21: return '14–21 ngày'
        return 'An toàn'

    res['Mức rủi ro'] = res.apply(classify_risk, axis=1)
    
    if date_col and 'End_Time_Raw' in res.columns:
        parsed_dates = res[date_col].apply(lambda x: parse_gantt_dates(x, ref_dt_val.year))
        res['Start'] = [p[0] for p in parsed_dates]
        parsed_ends = res['End_Time_Raw'].apply(lambda x: parse_gantt_dates(x, ref_dt_val.year)[0])
        res['End'] = parsed_ends.combine_first(pd.Series([p[1] for p in parsed_dates], index=res.index))
        res.loc[res['Start'] == res['End'], 'End'] += pd.Timedelta(hours=23, minutes=59)
        res['Mã rút gọn'] = res['Job Order'].astype(str).str[:3].str.upper()

    return res, free_dates, dynamic_notes

# =============================================================
# TV FULL HD — thêm ?tv=1 vào địa chỉ ứng dụng để trình chiếu.
# Dependencies: streamlit>=1.37, pandas, plotly>=5.24,
#               requests, openpyxl, xlrd, tzdata
# =============================================================
import html
import json
import math
import time as clock
from datetime import timedelta
import streamlit.components.v1 as components
from plotly.offline import get_plotlyjs

TV_URL = ('https://docs.google.com/spreadsheets/d/e/'
          '2PACX-1vTQOMzsXaj_Ed_ooA9x8LJ8NTkikDIBYVGs87h-ajD9FYjWHktL-MrzVcGqxFqRcFaNkTHzcH-xLARR/'
          'pub?output=xlsx')
TV_SECONDS = 25
VN = ZoneInfo('Asia/Ho_Chi_Minh')
TV_COLORS = {'Quá hạn':'#ef4444','< 7 ngày':'#f97316',
             '7–14 ngày':'#eab308','14–21 ngày':'#3b82f6',
             'An toàn':'#22c55e','Thiếu hạn':'#94a3b8'}

def tv_daytime(now):
    return 450 <= now.hour * 60 + now.minute < 1170

def tv_slot(now):
    return now.replace(minute=(now.minute // 30)*30, second=0, microsecond=0)

def tv_should_fetch(now, last_slot, retry_at):
    return tv_daytime(now) and (last_slot is None or tv_slot(now) > last_slot) and now >= retry_at

def tv_text(value):
    if value is None or (not isinstance(value, (list, dict)) and pd.isna(value)):
        return '—'
    return html.escape(str(value))

def tv_snapshot(url):
    # Gọi bản gốc, bỏ cache 10 phút để nhận dữ liệu mới đúng lịch TV.
    data, error = fetch_excel_from_url.__wrapped__(url)
    if error:
        raise ValueError(error)
    now = datetime.now(VN)
    frame, free, notes = load_and_preprocess_data(data, str(now.date()))
    if frame.empty:
        raise ValueError('File rỗng hoặc không đọc được các sheet kế hoạch.')
    return dict(df=frame, free=free, notes=notes, updated=now, ref=now.date())

def tv_document(snapshot, error=''):
    df = snapshot['df']
    df = df[~df['Status'].str.lower().isin(['done','services'])].copy()
    risks = list(TV_COLORS)
    slides = []
    def add(title, body):
        slides.append('<section class="slide"><h1>'+html.escape(title)+'</h1><div class="body">'+body+'</div></section>')
    def chart(title, fig):
        fig.update_layout(autosize=True, height=None, width=None,
            font=dict(size=28), title=None,
            margin=dict(l=100,r=110,t=30,b=100),
            showlegend=False,
            paper_bgcolor='white',plot_bgcolor='white')
        fig.update_xaxes(automargin=True)
        fig.update_yaxes(automargin=True)
        # Chú giải nằm ngoài vùng vẽ, không đè nhãn trục hoặc bị cắt mép.
        legends = []
        seen = set()
        for trace in fig.data:
            if trace.type == 'pie':
                entries = [(label, TV_COLORS.get(label, '#94a3b8')) for label in trace.labels]
            else:
                color = TV_COLORS.get(trace.name)
                if color is None:
                    color = trace.line.color if trace.type == 'scatter' else trace.marker.color
                entries = [(trace.name, color or '#3b82f6')]
            for label, color in entries:
                if not label or label in seen:
                    continue
                seen.add(label)
                legends.append('<span><i style="background:'+html.escape(str(color))+'"></i>'+html.escape(str(label))+'</span>')
        legend_html = '<div class="chart-legend">'+''.join(legends)+'</div>'
        payload=fig.to_json().replace('</','<\\/')
        add(title,legend_html+'<div class="chart"></div><script type="application/json" class="figure">'+payload+'</script>')

    counts=df['Mức rủi ro'].value_counts()
    cards=[]
    labels=['Quá hạn','< 7 ngày','7–14 ngày','14–21 ngày','An toàn','Tổng tồn']
    for label in labels:
        color=('#9333EA' if label=='Tổng tồn' else '#facc15' if label=='7–14 ngày' else TV_COLORS[label])
        value=len(df) if label=='Tổng tồn' else int(counts.get(label,0))
        sub={'Quá hạn':'Cần xử lý','< 7 ngày':'Cần xử lý hôm nay','7–14 ngày':'Theo dõi sát tiến độ',
             '14–21 ngày':'Không chủ quan','An toàn':'> 21 ngày','Tổng tồn':'Chưa hoàn thành'}[label]
        cards.append(f'<div class="kpi" style="background:{color};color:{"#111827" if label=="7–14 ngày" else "white"}"><div>{html.escape(label)}</div><strong>{value}</strong><small>{sub}</small></div>')
    add('TỔNG QUAN CHỈ SỐ RỦI RO','<div class="grid">'+''.join(cards)+'</div>')
    if not df.empty:
        top=df['Customer'].value_counts().head(12).index
        grouped=df[df.Customer.isin(top)].groupby(['Customer','Mức rủi ro']).size().reset_index(name='Số đơn')
        fig=px.bar(grouped,x='Customer',y='Số đơn',color='Mức rủi ro',text='Số đơn',
            color_discrete_map=TV_COLORS,category_orders={'Mức rủi ro':risks})
        fig.update_layout(barmode='stack',xaxis=dict(categoryorder='total descending'))
        fig.update_traces(textposition='inside',textfont_size=28)
        chart('Phân bổ đơn hàng theo Khách hàng & Mức rủi ro',fig)
        rc=counts.rename_axis('Mức rủi ro').reset_index(name='Số lượng')
        fig=px.pie(rc,values='Số lượng',names='Mức rủi ro',color='Mức rủi ro',hole=.55,
            color_discrete_map=TV_COLORS,category_orders={'Mức rủi ro':risks})
        fig.update_traces(sort=False,textinfo='percent+value',textfont_size=32)
        chart('Tỷ lệ phân bổ rủi ro tổng thể',fig)
        qty=pd.to_numeric(df['Qty'],errors='coerce').fillna(0)
        for title,series,color in [
            ('Pareto — Số đơn hàng (MQL) theo khách hàng',df.groupby('Customer').size(),'#3b82f6'),
            ('Pareto — Khối lượng chi tiết (Qty) theo khách hàng',df.assign(Qty_num=qty).groupby('Customer').Qty_num.sum(),'#10b981')]:
            series=series[series>0].sort_values(ascending=False)
            if series.empty: continue
            fig=go.Figure([go.Bar(x=series.index,y=series.values,name='Số đơn' if 'MQL' in title else 'Tổng Qty',marker_color=color),
                go.Scatter(x=series.index,y=series.cumsum()/series.sum()*100,name='Lũy kế %',yaxis='y2',
                mode='lines+markers',line=dict(color='#ef4444',width=3))])
            fig.update_layout(yaxis2=dict(title='Lũy kế (%)',overlaying='y',side='right',range=[0,115]))
            chart(title,fig)

    machines=FACTORY_MACHINES
    rows='<tr><th>Máy</th>'+''.join('<th>'+html.escape(r)+'</th>' for r in risks)+'</tr>'
    for machine in machines:
        values=df[df['Machine Name']==machine]['Mức rủi ro'].value_counts()
        rows+='<tr><th>'+machine+'</th>'
        for risk in risks:
            value=int(values.get(risk,0)); color=TV_COLORS[risk] if value else '#f4f5f7'
            fg='#111827' if risk=='7–14 ngày' else 'white'
            rows+=f'<td style="background:{color};color:{fg if value else "#999"}">{value}</td>'
        rows+='</tr>'
    add('Heatmap cảnh báo theo máy — Toàn bộ 24 máy','<table class="heat">'+rows+'</table>')
    if {'Start','End'}.issubset(df.columns):
        gantt=df[df['Machine Name'].isin(machines)].dropna(subset=['Start','End']).copy()
        for machine in machines:
            if machine not in set(gantt['Machine Name']):
                gantt=pd.concat([gantt,pd.DataFrame([{'Machine Name':machine,'Start':pd.Timestamp(snapshot['ref']),
                  'End':pd.Timestamp(snapshot['ref'])+pd.Timedelta(minutes=1),'Mức rủi ro':'Thiếu hạn',
                  'Mã rút gọn':' ','Job Order':'Chưa xếp lịch'}])],ignore_index=True)
        fig=px.timeline(gantt,x_start='Start',x_end='End',y='Machine Name',color='Mức rủi ro',
            text='Mã rút gọn',hover_name='Job Order',color_discrete_map=TV_COLORS,
            category_orders={'Machine Name':machines,'Mức rủi ro':risks})
        fig.update_yaxes(autorange='reversed',title='',categoryorder='array',
            categoryarray=machines,tickmode='array',tickvals=machines,tickfont=dict(size=23))
        fig.update_traces(textfont=dict(size=20,color='white'),textposition='inside')
        chart('Lịch trình chạy máy — Toàn bộ 24 máy',fig)

    for start in range(0,len(FACTORY_MACHINES),6):
        cards=[]
        for machine in FACTORY_MACHINES[start:start+6]:
            orders=df[df['Machine Name']==machine]
            crit=int(orders['Mức rủi ro'].isin(['Quá hạn','< 7 ngày']).sum())
            waiting=int(orders.Status.str.lower().eq('waiting').sum())
            note=snapshot['notes'].get(machine,'')
            edge='#ef4444' if crit else '#94a3b8' if orders.empty else '#f97316' if note else '#22c55e'
            badge=f'{crit} rủi ro' if crit else 'Rảnh máy' if orders.empty else 'Lưu ý' if note else 'OK'
            mql='(Chưa có lệnh)'; details='Máy đang trống'
            if not orders.empty:
                valid=orders.dropna(subset=['Real_Start']) if 'Real_Start' in orders else orders.iloc[:0]
                began=valid[valid.Real_Start<=snapshot['updated'].replace(tzinfo=None)] if not valid.empty else valid
                row=(began.sort_values('Real_Start').iloc[-1] if not began.empty else
                     valid.sort_values('Real_Start').iloc[0] if not valid.empty else orders.iloc[0])
                mql=tv_text(row['Job Order']); details='BV: '+tv_text(row.get('Drawing No.'))+' · SL: '+tv_text(row.get('Qty'))
            free=snapshot['free'].get(machine,'Chưa xếp lịch')
            if isinstance(free,(date,datetime,pd.Timestamp)): free=free.strftime('%d/%m')
            elif '-' in str(free):
                if re.match(r'^\d{4}-\d{2}-\d{2}',str(free)):
                    free=pd.to_datetime(str(free)[:10]).strftime('%d/%m')
                else: free=str(free).split('-')[-1].strip()
            cards.append(f'<div class="machine" style="border-left:8px solid {edge}"><header><b>{machine}</b><span style="color:{edge}">{badge}</span></header><small>Đang chạy:</small><div class="mql">{mql}</div><div>{details}</div><footer>Đợi: <b>{waiting}</b> lệnh · <span style="color:#3b82f6">Rảnh: {tv_text(free)}</span></footer><em>{tv_text(note) if note else ""}</em></div>')
        add(f'Trạng thái Máy & Dự kiến rảnh máy — Nhóm {start//6+1}','<div class="grid">'+''.join(cards)+'</div>')

    # Khối #11 chỉ giữ trong chế độ máy tính.
    updated=snapshot['updated'].strftime('%d/%m/%Y %H:%M')
    status='Cập nhật thành công: '+updated+' · Ngày phân tích: '+snapshot['ref'].strftime('%d/%m/%Y')
    if error: status+=' · Chưa tải được bản mới; đang giữ dữ liệu cũ.'
    css='''
    *{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;color:#111827;background:white;overflow:hidden}
    .slide{display:none;height:100vh;padding:16px 24px 72px}.slide.active{display:flex;flex-direction:column}
    h1{font-size:38px;margin:0 0 18px;flex:none}.body{flex:1;min-height:0;display:flex;flex-direction:column}.chart{width:100%;flex:1;min-height:0}
    .chart-legend{display:flex;flex-wrap:wrap;justify-content:center;gap:12px 28px;font-size:24px;padding:4px 8px 16px;flex:none}.chart-legend span{display:inline-flex;align-items:center;gap:8px;white-space:nowrap}.chart-legend i{width:20px;height:20px;display:inline-block;border-radius:3px}
    .grid{height:100%;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));grid-template-rows:repeat(2,minmax(0,1fr));gap:22px}
    .kpi{border-radius:12px;padding:28px;font-size:36px;display:flex;flex-direction:column;justify-content:center}
    .kpi strong{font-size:96px;margin:12px 0}.kpi small{font-size:28px}
    table{width:100%;height:100%;border-collapse:collapse;table-layout:fixed}th,td{border:1px solid #ccd0d6;padding:7px;text-align:center}
    th{background:#f4f5f7;font-size:26px}.heat{flex:1;min-height:0}.heat th,.heat td{padding:1px 5px;font-size:24px;line-height:1.05}.heat td{font-weight:bold}.heat tr{height:4%}
    .detail td{font-size:24px;overflow-wrap:anywhere}.detail th{font-size:24px}.detail th:nth-child(4){width:16%}.detail th:nth-child(9){width:14%}
    .machine{border:1px solid #ddd;border-radius:10px;padding:22px;font-size:28px;overflow:auto}
    .machine header{display:flex;justify-content:space-between;border-bottom:1px solid #ddd;padding-bottom:14px;margin-bottom:14px}
    .machine header b{font-size:38px}.machine small{font-size:23px}.mql{font-size:32px;font-weight:bold;overflow-wrap:anywhere;margin:8px 0}
    .machine footer{border-top:1px dashed #ccc;margin-top:18px;padding-top:14px}.machine em{color:#f59e0b;font-size:25px}
    #controls{position:fixed;bottom:0;left:0;right:0;height:54px;background:#f4f5f7;display:flex;align-items:center;gap:12px;padding:8px 150px 8px 20px;font-size:20px}
    button{font-size:20px;padding:5px 12px;cursor:pointer}#stamp{flex:1}a{color:#3b82f6}
    '''
    js='''
    const slides=Array.from(document.querySelectorAll('.slide'));
    let index=Number(sessionStorage.getItem('sg-tv-page')||0)%slides.length;
    let paused=sessionStorage.getItem('sg-tv-paused')==='1';
    function show(){
      slides.forEach((s,i)=>s.classList.toggle('active',i===index));
      sessionStorage.setItem('sg-tv-page',String(index));
      const slide=slides[index], el=slide.querySelector('.chart');
      if(el){if(!el.dataset.ready){const p=JSON.parse(slide.querySelector('.figure').textContent);Plotly.newPlot(el,p.data,p.layout,{responsive:true,displayModeBar:false});el.dataset.ready='1';}else{Plotly.Plots.resize(el);}}
      document.getElementById('page').textContent=(index+1)+' / '+slides.length;
      document.getElementById('pause').textContent=paused?'▶ Tiếp tục':'Ⅱ Tạm dừng';
      const parts=new Intl.DateTimeFormat('en-GB',{timeZone:'Asia/Ho_Chi_Minh',hour:'2-digit',minute:'2-digit',hour12:false}).format(new Date()).split(':');
      const mins=Number(parts[0])*60+Number(parts[1]);
      document.getElementById('night').textContent=mins>=450&&mins<1170?'Tự cập nhật 30 phút':'Tạm dừng cập nhật ban đêm';
    }
    function move(n){index=(index+n+slides.length)%slides.length;show();}
    let timer;function reset(){clearInterval(timer);timer=setInterval(()=>{if(!paused)move(1)},SECONDS*1000);}
    document.getElementById('prev').onclick=()=>{move(-1);reset()};
    document.getElementById('next').onclick=()=>{move(1);reset()};
    document.getElementById('pause').onclick=()=>{paused=!paused;sessionStorage.setItem('sg-tv-paused',paused?'1':'0');show();reset()};
    document.addEventListener('keydown',e=>{if(e.key==='ArrowRight')move(1);if(e.key==='ArrowLeft')move(-1);if(e.code==='Space'){e.preventDefault();document.getElementById('pause').click();}});
    window.addEventListener('resize',()=>{const el=slides[index].querySelector('.chart');if(el&&el.dataset.ready)Plotly.Plots.resize(el)});
    show();reset();
    '''.replace('SECONDS',str(TV_SECONDS))
    return '<!doctype html><html><head><meta charset="utf-8"><style>'+css+'</style><script>'+get_plotlyjs()+'</script></head><body>'+''.join(slides)+'<div id="controls"><span id="stamp">'+html.escape(status)+'</span><span id="night"></span><button id="prev">◀</button><button id="pause"></button><button id="next">▶</button><span id="page"></span></div><script>'+js+'</script></body></html>'

TV_MODE = st.query_params.get('tv','0') == '1'
if TV_MODE:
    if not hasattr(st,'fragment'):
        st.error('Cần Streamlit >= 1.37. Hãy cập nhật requirements.txt.'); st.stop()
    st.markdown('''<style>
    [data-testid="stSidebar"],[data-testid="stHeader"],footer{display:none!important}
    .block-container,.stMainBlockContainer{padding:0!important;max-width:100%!important}
    [data-testid="stButton"]{position:fixed!important;right:20px;bottom:90px;z-index:10000;width:auto!important}
    [data-testid="stButton"] button{min-height:38px;font-size:20px;background:#f4f5f7;color:#3b82f6;border:1px solid #ccd0d6}
    iframe[title="st.iframe"]{height:100vh!important;width:100%!important;border:0}
    </style>''',unsafe_allow_html=True)
    now=datetime.now(VN)
    if 'tv_retry_at' not in st.session_state: st.session_state.tv_retry_at=now
    # Nguồn URL có thể được chọn ở giao diện máy tính trước khi mở TV.
    source=st.session_state.get('saved_link') or TV_URL
    if st.session_state.get('tv_source')!=source:
        for key in ['tv_snapshot','tv_slot','tv_error']: st.session_state.pop(key,None)
        st.session_state.tv_source=source
        st.session_state.tv_retry_at=now
    initial='tv_snapshot' not in st.session_state
    due=tv_should_fetch(now,st.session_state.get('tv_slot'),st.session_state.tv_retry_at)
    if (initial and now>=st.session_state.tv_retry_at) or due:
        try:
            fresh=tv_snapshot(source)
            st.session_state.tv_snapshot=fresh
            st.session_state.tv_slot=tv_slot(now)
            st.session_state.tv_error=''
        except Exception as exc:
            st.session_state.tv_error=str(exc)
            st.session_state.tv_retry_at=now+timedelta(minutes=5)
    @st.fragment(run_every=30)
    def tv_tick():
        current=datetime.now(VN)
        missing='tv_snapshot' not in st.session_state
        if (missing and current>=st.session_state.tv_retry_at) or tv_should_fetch(current,st.session_state.get('tv_slot'),st.session_state.tv_retry_at):
            st.rerun()
    tv_tick()
    if 'tv_snapshot' in st.session_state:
        components.html(tv_document(st.session_state.tv_snapshot,st.session_state.get('tv_error','')),height=1000,scrolling=False)
    else:
        st.error('Chưa tải được dữ liệu: '+st.session_state.get('tv_error',''))
        st.info('Ứng dụng tự thử lại sau 5 phút. Có thể quay lại chế độ máy tính để kiểm tra nguồn.')
        st.link_button('Về chế độ máy tính','?tv=0')
    # Nút native nằm ngoài iframe: đổi chế độ qua Streamlit, không dùng điều hướng iframe.
    if st.button('Máy tính', key='tv_exit_desktop'):
        st.query_params['tv'] = '0'
        st.rerun()
    st.stop()

if st.sidebar.button('📺 Mở chế độ TV Full HD',type='primary'):
    st.query_params['tv']='1'
    st.rerun()
st.sidebar.caption('TV: tự chuyển 25 giây; tải mới 07:30–19:00 mỗi 30 phút. F11 để toàn màn hình.')

# 1.5 HEADER VÀ Ô CHỌN NGÀY THAM CHIẾU
# -------------------------------------------------------------
h_col1, h_col2 = st.columns([4, 1])
with h_col1:
    st.markdown("""<div class="header-box">
<h2>Master Plan SG — Cảnh báo Delay & Theo dõi MQL</h2>
<p>Ưu tiên xem theo Khách hàng / MQL · Bỏ qua các đơn có cột H = Done / Services</p>
</div>""", unsafe_allow_html=True)
with h_col2:
    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
    ref_date = st.date_input("Ngày tham chiếu phân tích",
    datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).date())

# -------------------------------------------------------------
# 2. TỐI ƯU HÓA: AUTO-FIX LINK CHỐNG LỖI 404 & CHỐNG CHẶN BOT
# -------------------------------------------------------------

st.sidebar.header("📁 Cập nhật kế hoạch")

if "saved_link" not in st.session_state:
    st.session_state["saved_link"] = ""

if st.sidebar.button("🔄 Làm mới dữ liệu (Tải lại từ đầu)", type="primary"):
    st.cache_data.clear()

data_source = st.sidebar.radio(
    "Chọn phương thức tải dữ liệu:", 
    ("Dùng Link Google Sheets/Drive", "Tải file Excel từ máy")
)

raw_file_bytes = None

if data_source == "Tải file Excel từ máy":
    uploaded_file = st.sidebar.file_uploader("Tải file Excel", type=["xlsx", "xls"])
    if uploaded_file:
        raw_file_bytes = uploaded_file.getvalue()
else:
    st.sidebar.info("💡 Bạn có thể dán Link Công bố (Publish) hoặc Link Chia sẻ (Share) đều được. Hệ thống sẽ tự động xử lý.")
    gsheet_url = st.sidebar.text_input("Dán Link vào đây:", key="saved_link")
    
    if not gsheet_url and not st.session_state.get('user_changed_link', False):
        gsheet_url = ("https://docs.google.com/spreadsheets/d/e/"
                      "2PACX-1vTQOMzsXaj_Ed_ooA9x8LJ8NTkikDIBYVGs87h-ajD9FYjWHktL-MrzVcGqxFqRcFaNkTHzcH-xLARR/"
                      "pub?output=xlsx")
                      
    st.sidebar.caption("Nguồn dữ liệu mặc định: Master Plan SG")
    
    if gsheet_url:
        with st.sidebar.status("Đang tải dữ liệu...", expanded=False) as status:
            content, err = fetch_excel_from_url(gsheet_url)
            if err:
                status.update(label=err, state="error")
            else:
                raw_file_bytes = content
                status.update(label="Dữ liệu đã sẵn sàng!", state="complete")



# -------------------------------------------------------------
# 3. HÀM ĐỌC DỮ LIỆU & LẤY NGÀY NỐI ĐUÔI
# -------------------------------------------------------------


if raw_file_bytes is not None:
    df_raw, machine_free_dates, parsed_notes = load_and_preprocess_data(raw_file_bytes, str(ref_date))
    
    for m_code, text_note in parsed_notes.items():
        if m_code in MACHINE_DETAILS:
            MACHINE_DETAILS[m_code]["note"] = text_note
        else:
            MACHINE_DETAILS[m_code] = {"a": "—", "b": "—", "note": text_note}
else:
    df_raw, machine_free_dates, parsed_notes = pd.DataFrame(), {}, {}

if df_raw.empty:
    st.info("👋 Vui lòng Tải file Excel hoặc Dán link chia sẻ/công bố ở thanh menu bên trái để bắt đầu.")
    st.stop()

df_open = df_raw[~df_raw['Status'].astype(str).str.strip().str.lower().isin(['done', 'services'])].copy()

# -------------------------------------------------------------
# 5. BỘ LỌC TƯƠNG TÁC
# -------------------------------------------------------------
f1, f2, f3, f4 = st.columns([1.5, 1.5, 1.5, 2.5])
with f1: sel_cust = st.selectbox("Khách hàng", ['Tất cả'] + sorted(list(df_open['Customer'].unique())))
with f2: sel_mach = st.selectbox("Máy", ['Tất cả'] + sorted(list(df_open['Machine Name'].unique())))
with f3: sel_risk = st.selectbox("Mức rủi ro", ['Tất cả', 'Quá hạn', '< 7 ngày', '7–14 ngày', '14–21 ngày', 'An toàn'])
with f4: search_q = st.text_input("Tìm MQL / PO / BV", placeholder="VD: NLA2608, HC_64...").strip().lower()

df_filtered = df_open.copy()
if sel_cust != 'Tất cả': df_filtered = df_filtered[df_filtered['Customer'] == sel_cust]
if sel_mach != 'Tất cả': df_filtered = df_filtered[df_filtered['Machine Name'] == sel_mach]
if sel_risk != 'Tất cả': df_filtered = df_filtered[df_filtered['Mức rủi ro'] == sel_risk]
if search_q:
    df_filtered = df_filtered[
        df_filtered['Job Order'].str.lower().str.contains(search_q, na=False) |
        df_filtered['PO'].str.lower().str.contains(search_q, na=False) |
        df_filtered['Drawing No.'].astype(str).str.lower().str.contains(search_q, na=False)
    ]

# -------------------------------------------------------------
# 6. KHỐI THẺ KPI & NÚT XUẤT EXCEL
# -------------------------------------------------------------
n_over = len(df_filtered[df_filtered['Mức rủi ro'] == 'Quá hạn'])
n_cam  = len(df_filtered[df_filtered['Mức rủi ro'] == '< 7 ngày'])
n_vang = len(df_filtered[df_filtered['Mức rủi ro'] == '7–14 ngày'])
n_blue = len(df_filtered[df_filtered['Mức rủi ro'] == '14–21 ngày'])
n_safe = len(df_filtered[df_filtered['Mức rủi ro'] == 'An toàn'])
n_open = len(df_filtered)

worst_over = df_filtered[df_filtered['Mức rủi ro'] == 'Quá hạn'].sort_values('Còn (ngày)')
worst_txt = f"Trễ nhất: {worst_over.iloc[0]['Job Order']} ({-worst_over.iloc[0]['Còn (ngày)']} ngày)" if not worst_over.empty else "Không có"

df_summary = pd.DataFrame({
    "Ngày xuất": [datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).strftime("%d/%m/%Y %H:%M")],
    "Số đơn quá hạn": [n_over], "< 7 ngày (Cần xử lý)": [n_cam], "7-14 ngày (Theo dõi)": [n_vang],
    "14-21 ngày (Không chủ quan)": [n_blue], "> 21 ngày (An toàn)": [n_safe], "Tổng đơn chưa hoàn thành": [n_open]
})

@st.cache_data(show_spinner=False)
def convert_summary_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer: df.to_excel(writer, index=False, sheet_name='ThongKe_KPI')
    return output.getvalue()

btn_col1, btn_col2 = st.columns([7, 2])
with btn_col1: st.markdown("##### 📊 TỔNG QUAN CHỈ SỐ RỦI RO")
with btn_col2:
    st.download_button(label="📥 Tải thống kê KPI (Excel)", data=convert_summary_to_excel(df_summary), file_name=f"ThongKe_KPI_{datetime.now(ZoneInfo('Asia/Ho_Chi_Minh')).strftime('%Y%m%d_%H%M')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1: st.markdown(f'<div class="kpi-card kpi-red"><div class="kpi-title">Quá hạn</div><div class="kpi-value">{n_over}</div><div class="kpi-sub">{worst_txt}</div></div>', unsafe_allow_html=True)
with k2: st.markdown(f'<div class="kpi-card kpi-orange"><div class="kpi-title">&lt; 7 ngày</div><div class="kpi-value">{n_cam}</div><div class="kpi-sub">Cần xử lý hôm nay</div></div>', unsafe_allow_html=True)
with k3: st.markdown(f'<div class="kpi-card kpi-amber"><div class="kpi-title">7–14 ngày</div><div class="kpi-value">{n_vang}</div><div class="kpi-sub">Theo dõi sát tiến độ</div></div>', unsafe_allow_html=True)
with k4: st.markdown(f'<div class="kpi-card kpi-blue"><div class="kpi-title">14–21 ngày</div><div class="kpi-value">{n_blue}</div><div class="kpi-sub">Không chủ quan</div></div>', unsafe_allow_html=True)
with k5: st.markdown(f'<div class="kpi-card kpi-green"><div class="kpi-title">An toàn</div><div class="kpi-value">{n_safe}</div><div class="kpi-sub">&gt; 21 ngày</div></div>', unsafe_allow_html=True)
with k6: st.markdown(f'<div class="kpi-card kpi-black"><div class="kpi-title">Tổng tồn</div><div class="kpi-value">{n_open}</div><div class="kpi-sub">Chưa hoàn thành</div></div>', unsafe_allow_html=True)

color_map = {'Quá hạn': '#ef4444', '< 7 ngày': '#f97316', '7–14 ngày': '#eab308', '14–21 ngày': '#3b82f6', 'An toàn': '#22c55e', 'Thiếu hạn': '#94a3b8'}
risk_order = ['Quá hạn', '< 7 ngày', '7–14 ngày', '14–21 ngày', 'An toàn', 'Thiếu hạn']

# -------------------------------------------------------------
# 7. KHỐI BIỂU ĐỒ TRÒN & CỘT TỔNG QUAN
# -------------------------------------------------------------
st.markdown("---")
c_left, c_right = st.columns([1.2, 0.8])

with c_left:
    st.markdown("##### Phân bổ đơn hàng theo Khách hàng & Mức rủi ro")
    if not df_filtered.empty:
        top_cust = df_filtered['Customer'].value_counts().head(12).index
        c_counts = df_filtered[df_filtered['Customer'].isin(top_cust)].groupby(['Customer', 'Mức rủi ro']).size().reset_index(name='Số đơn')
        c_counts['Mức rủi ro'] = pd.Categorical(c_counts['Mức rủi ro'], categories=risk_order, ordered=True)
        c_counts = c_counts.sort_values(['Customer', 'Mức rủi ro'])

        fig_bar = px.bar(c_counts, x='Customer', y='Số đơn', color='Mức rủi ro', color_discrete_map=color_map, text='Số đơn', category_orders={"Mức rủi ro": risk_order})
        fig_bar.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=260, yaxis=dict(title='Số đơn'), xaxis=dict(title='Khách hàng', categoryorder='total descending'), barmode='stack', showlegend=False)
        fig_bar.update_traces(textposition='inside', insidetextanchor='middle')
        st.plotly_chart(fig_bar, use_container_width=True)
    else: st.info("Không có dữ liệu để hiển thị biểu đồ.")

with c_right:
    st.markdown("##### Tỷ lệ phân bổ rủi ro tổng thể")
    if not df_filtered.empty:
        risk_counts = df_filtered['Mức rủi ro'].value_counts().reset_index()
        risk_counts.columns = ['Mức rủi ro', 'Số lượng']
        risk_counts['Mức rủi ro'] = pd.Categorical(risk_counts['Mức rủi ro'], categories=risk_order, ordered=True)
        risk_counts = risk_counts.sort_values('Mức rủi ro')

        fig_donut = px.pie(risk_counts, values='Số lượng', names='Mức rủi ro', color='Mức rủi ro', color_discrete_map=color_map, hole=0.55, category_orders={"Mức rủi ro": risk_order})
        fig_donut.update_traces(sort=False)
        fig_donut.update_layout(margin=dict(l=10, r=10, t=10, b=10), height=260, legend=dict(traceorder='normal'))
        st.plotly_chart(fig_donut, use_container_width=True)
    else: st.info("Không có dữ liệu để hiển thị.")


# -------------------------------------------------------------
# 7.5 KHỐI BIỂU ĐỒ PARETO
# -------------------------------------------------------------
st.markdown("---")
st.markdown("##### 📈 Phân tích Pareto Khách hàng (Quy tắc 80/20)")
p_col1, p_col2 = st.columns(2)

if not df_filtered.empty:
    df_pareto_order = df_filtered.groupby('Customer').size().reset_index(name='Số đơn').sort_values(by='Số đơn', ascending=False)
    df_pareto_order['Lũy kế (%)'] = df_pareto_order['Số đơn'].cumsum() / df_pareto_order['Số đơn'].sum() * 100

    fig_p1 = go.Figure()
    fig_p1.add_trace(go.Bar(x=df_pareto_order['Customer'], y=df_pareto_order['Số đơn'], name='Số đơn', marker_color='#3b82f6'))
    fig_p1.add_trace(go.Scatter(x=df_pareto_order['Customer'], y=df_pareto_order['Lũy kế (%)'], name='Lũy kế %', yaxis='y2', mode='lines+markers', line=dict(color='#ef4444', width=2)))
    fig_p1.update_layout(
        title="Khách hàng có nhiều đơn hàng (MQL) nhất", yaxis=dict(title='Số đơn'), yaxis2=dict(title='Lũy kế (%)', overlaying='y', side='right', range=[0, 115]),
        legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, bgcolor='rgba(0,0,0,0)'), margin=dict(l=10, r=10, t=50, b=10), height=320
    )

    df_qty = df_filtered.copy()
    df_qty['Qty_num'] = pd.to_numeric(df_qty['Qty'], errors='coerce').fillna(0)
    df_pareto_qty = df_qty.groupby('Customer')['Qty_num'].sum().reset_index(name='Tổng Qty').sort_values(by='Tổng Qty', ascending=False)
    df_pareto_qty = df_pareto_qty[df_pareto_qty['Tổng Qty'] > 0]
    df_pareto_qty['Lũy kế (%)'] = df_pareto_qty['Tổng Qty'].cumsum() / df_pareto_qty['Tổng Qty'].sum() * 100

    fig_p2 = go.Figure()
    if not df_pareto_qty.empty:
        fig_p2.add_trace(go.Bar(x=df_pareto_qty['Customer'], y=df_pareto_qty['Tổng Qty'], name='Tổng Qty', marker_color='#10b981'))
        fig_p2.add_trace(go.Scatter(x=df_pareto_qty['Customer'], y=df_pareto_qty['Lũy kế (%)'], name='Lũy kế %', yaxis='y2', mode='lines+markers', line=dict(color='#ef4444', width=2)))
        fig_p2.update_layout(
            title="Khách hàng có khối lượng chi tiết (Qty) cao nhất", yaxis=dict(title='Tổng Qty'), yaxis2=dict(title='Lũy kế (%)', overlaying='y', side='right', range=[0, 115]),
            legend=dict(orientation="h", yanchor="bottom", y=1.05, xanchor="right", x=1, bgcolor='rgba(0,0,0,0)'), margin=dict(l=10, r=10, t=50, b=10), height=320
        )

    with p_col1: st.plotly_chart(fig_p1, use_container_width=True)
    with p_col2:
        if not df_pareto_qty.empty: st.plotly_chart(fig_p2, use_container_width=True)
        else: st.info("Không có dữ liệu số lượng Qty hợp lệ.")
else: st.info("Không có dữ liệu để vẽ biểu đồ Pareto.")

# -------------------------------------------------------------
# 8. BẢNG TỔNG HỢP RỦI RO THEO MÁY (MA TRẬN KHUNG VUÔNG & TÔ FULL MÀU Ô)
# -------------------------------------------------------------
st.markdown("---")
st.markdown("##### 🔥 Heatmap cảnh báo theo máy")
st.markdown("<p style='font-size:12.5px; opacity:0.8; margin-top:-5px;'>Theo dõi giám sát các máy có nguy cơ cao.</p>", unsafe_allow_html=True)

heatmap_data = pd.DataFrame(index=FACTORY_MACHINES, columns=risk_order).fillna(0)

if not df_filtered.empty:
    grouped_risk = df_filtered.groupby(['Machine Name', 'Mức rủi ro']).size().reset_index(name='Số lượng')
    for _, row in grouped_risk.iterrows():
        m_name = row['Machine Name']
        m_risk = row['Mức rủi ro']
        if m_name in FACTORY_MACHINES and m_risk in risk_order:
            heatmap_data.at[m_name, m_risk] = row['Số lượng']

heatmap_data = heatmap_data.astype(int)

risk_colors_hex = {
    'Quá hạn': '#ef4444',      # Đỏ
    '< 7 ngày': '#f97316',    # Cam
    '7–14 ngày': '#eab308',   # Vàng sáng
    '14–21 ngày': '#3b82f6',  # Xanh lam
    'An toàn': '#22c55e',      # Xanh lục
    'Thiếu hạn': '#94a3b8'     # Xám
}

# Xây dựng cấu trúc CSS đảm bảo khung tổng thể cân xứng hình vuông và các ô tô full màu
heatmap_html = '''
<style>
.hm-square-wrapper {
    display: flex;
    justify-content: flex-start;
    margin-bottom: 10px;
}
.hm-square-container {
    width: 650px;
    height: 650px;
    max-width: 100%;
    aspect-ratio: 1 / 1;
    overflow: auto;
    border: 2px solid rgba(128, 128, 128, 0.3);
    border-radius: 8px;
    background-color: var(--secondary-background-color);
    padding: 5px;
}
.hm-square-table {
    width: 100%;
    height: 100%;
    border-collapse: collapse;
    text-align: center;
    font-size: 13px;
    font-family: sans-serif;
    table-layout: fixed;
}
.hm-square-table th {
    padding: 8px 4px;
    font-weight: 700;
    color: var(--text-color);
    background-color: var(--background-color);
    position: sticky;
    top: 0;
    z-index: 2;
    border: 1px solid rgba(128,128,128,0.3);
    font-size: 12px;
}
.hm-square-table td {
    border: 1px solid rgba(128, 128, 128, 0.2);
    padding: 0;
    height: 100%;
}
.hm-full-cell {
    width: 100%;
    height: 100%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-size: 13px;
    min-height: 24px;
}
.hm-row-mach {
    font-weight: bold;
    text-align: center;
    background-color: var(--background-color);
    position: sticky;
    left: 0;
    z-index: 1;
    color: var(--text-color);
    border: 1px solid rgba(128,128,128,0.3) !important;
}
</style>
<div class="hm-square-wrapper">
<div class="hm-square-container">
<table class="hm-square-table">
<thead>
    <tr><th class="hm-row-mach" style="width: 15%;">Máy</th>
'''
for r in risk_order:
    heatmap_html += f"<th>{r}</th>"
heatmap_html += "</tr></thead><tbody>"

for m in FACTORY_MACHINES:
    heatmap_html += f"<tr><td class='hm-row-mach'>{m}</td>"
    for r in risk_order:
        val = heatmap_data.at[m, r]
        tooltip = f"Máy {m} | {r}: {val} đơn"
        
        if val == 0:
            heatmap_html += f"<td><div class='hm-full-cell' title='{tooltip}' style='color: rgba(128,128,128,0.4); background-color: transparent;'>0</div></td>"
        else:
            bg_color = risk_colors_hex[r]
            font_color = "#111827" if r == '7–14 ngày' else "#ffffff"
            heatmap_html += f"<td><div class='hm-full-cell' title='{tooltip}' style='background-color: {bg_color}; color: {font_color};'>{val}</div></td>"
    heatmap_html += "</tr>"
heatmap_html += "</tbody></table></div></div>"

st.markdown(heatmap_html, unsafe_allow_html=True)

# -------------------------------------------------------------
# 9. BIỂU ĐỒ GANTT LỊCH TRÌNH CHẠY MÁY NỐI TIẾP
# -------------------------------------------------------------
st.markdown("---")

if 'Start' in df_filtered.columns and not df_filtered.empty:
    df_gantt = df_filtered[df_filtered['Machine Name'].isin(FACTORY_MACHINES)].dropna(subset=['Start', 'End']).copy()
    
    existing_machines = df_gantt['Machine Name'].unique()
    missing_machines = [m for m in FACTORY_MACHINES if m not in existing_machines]
    
    if missing_machines:
        dummy_records = []
        for m in missing_machines:
            dummy_records.append({
                'Machine Name': m,
                'Start': pd.to_datetime(ref_date),
                'End': pd.to_datetime(ref_date) + pd.Timedelta(minutes=1), 
                'Mức rủi ro': 'Thiếu hạn', 
                'Job Order': 'Chưa xếp lịch',
                'Customer': '—',
                'Status': '—',
                'Qty': 0,
                'Mã rút gọn': ' ' 
            })
        df_gantt = pd.concat([df_gantt, pd.DataFrame(dummy_records)], ignore_index=True)

    if not df_gantt.empty:
        df_gantt['Machine Name'] = pd.Categorical(df_gantt['Machine Name'], categories=FACTORY_MACHINES, ordered=True)
        df_gantt = df_gantt.sort_values(by=['Machine Name', 'Start'])
        
        fig_gantt = px.timeline(
            df_gantt, x_start='Start', x_end='End', y='Machine Name', color='Mức rủi ro', color_discrete_map=color_map,
            hover_name='Job Order', hover_data={'Customer': True, 'Status': True, 'Qty': True, 'Mã rút gọn': False, 'Mức rủi ro': False, 'Machine Name': False},
            text='Mã rút gọn', category_orders={"Machine Name": FACTORY_MACHINES, "Mức rủi ro": risk_order}
        )
        
        fig_gantt.update_yaxes(autorange="reversed", categoryorder="array", categoryarray=FACTORY_MACHINES, title="", tickfont=dict(size=14, weight="bold"))
        fig_gantt.update_xaxes(title="Trục thời gian")
        
        # Thêm tiêu đề vào trực tiếp trong khung biểu đồ Plotly để hiển thị khi Fullscreen
        fig_gantt.update_layout(
            title=dict(
                text="📅 Lịch trình chạy máy ",
                font=dict(size=20, weight="bold"),
                x=0.0,
                y=1.
            ),
            height=max(450, len(FACTORY_MACHINES) * 38), 
            margin=dict(l=10, r=10, t=50, b=10), 
            showlegend=True, 
            legend_title_text='Mức rủi ro'
        )
        
        fig_gantt.update_traces(textfont=dict(size=14, color='white', weight='bold'), textposition='inside', insidetextanchor='middle')
        st.plotly_chart(fig_gantt, use_container_width=True)
    else: st.info("⚠️️ Không có dữ liệu lịch chạy hợp lệ để vẽ biểu đồ.")
else: st.error("❌ Không tìm thấy cột chứa dữ liệu ngày tháng trong file Excel của bạn.")

st.markdown("---")

# -------------------------------------------------------------
# 10. KHỐI: THẺ TRẠNG THÁI MÁY CHI TIẾT
# -------------------------------------------------------------
st.markdown("##### ⚙ Trạng thái Máy & Dự kiến rảnh máy *(Load dữ liệu từ cột Start & End Date)*")

current_time_vn = datetime.now(ZoneInfo("Asia/Ho_Chi_Minh")).replace(tzinfo=None)

m_cols = st.columns(4)
for idx, m_code in enumerate(FACTORY_MACHINES):
    col = m_cols[idx % 4]
    
    mach_orders = df_open[df_open['Machine Name'] == m_code].copy()
    m_wait = mach_orders[mach_orders['Status'].astype(str).str.strip().str.lower() == 'waiting']
    crit_count = len(mach_orders[mach_orders['Mức rủi ro'].isin(['Quá hạn', '< 7 ngày'])])
    m_info = MACHINE_DETAILS.get(m_code, {"a": "—", "b": "—", "note": ""})
    
    curr_row = None
    if not mach_orders.empty: 
        if 'Real_Start' in mach_orders.columns:
            valid_runs = mach_orders.dropna(subset=['Real_Start'])
            if not valid_runs.empty:
                started_runs = valid_runs[valid_runs['Real_Start'] <= current_time_vn]
                if not started_runs.empty:
                    curr_row = started_runs.sort_values(by='Real_Start').iloc[-1]
                else:
                    curr_row = valid_runs.sort_values(by='Real_Start').iloc[0]
            else:
                curr_row = mach_orders.iloc[0]
        else:
            curr_row = mach_orders.iloc[0]
            
        mql_name = curr_row['Job Order']
        sub_info = f"BV: {curr_row['Drawing No.']} · SL: {curr_row['Qty']}"
    else: 
        mql_name = "(Chưa có lệnh)"
        sub_info = "Máy đang trống"
    
    free_date_text = "Chưa xếp lịch"
    raw_val = machine_free_dates.get(m_code, None)
    
    if isinstance(raw_val, pd.Series): 
        raw_val = raw_val.iloc[0]
        
    if raw_val is not None and str(raw_val).strip() != '':
        if isinstance(raw_val, (datetime, date, pd.Timestamp)): free_date_text = raw_val.strftime('%d/%m')
        else:
            val_str = str(raw_val).strip()
            if re.match(r'^\d{4}-\d{2}-\d{2}', val_str):
                try: free_date_text = pd.to_datetime(val_str[:10], errors='coerce', dayfirst=True).strftime('%d/%m')
                except: free_date_text = val_str[:10]
            elif '-' in val_str: free_date_text = val_str.split('-')[-1].strip()
            else: free_date_text = val_str

    if crit_count > 0: cls_name = "crit"; badge = f'<span class="badge-red">{crit_count} rủi ro</span>'
    elif mach_orders.empty: cls_name = "idle"; badge = '<span class="badge-gry">Rảnh máy</span>'
    elif m_info["note"]: cls_name = "warn"; badge = '<span class="badge-org">Lưu ý</span>'
    else: cls_name = ""; badge = '<span class="badge-grn">OK</span>'
        
    note_html = f'<div style="font-size:10.5px; color:#f59e0b !important; margin-top:3px; font-style:italic;">{m_info["note"]}</div>' if m_info["note"] else ''
    
    with col:
        st.markdown(f"""<div class="m-card {cls_name}">
<div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid rgba(128,128,128,0.2); padding-bottom:5px; margin-bottom:6px;">
<span class="mach-badge">{m_code}</span>{badge}
</div>
<div style="font-size:10px; text-transform:uppercase; opacity:0.7; margin-top:5px; letter-spacing:0.5px;">Đang chạy:</div>
<div style="font-family:ui-monospace, Consolas, monospace; font-size:12.5px; color:var(--primary-color) !important; font-weight:700; word-break:break-all; margin-top:1px;">{mql_name}</div>
<div style="font-size:11px; color:var(--text-color) !important; opacity:0.85; margin-top:1px;">{sub_info}</div>
<div style="margin-top:8px; padding-top:6px; border-top:1px dashed rgba(128,128,128,0.2); display:flex; justify-content:space-between; align-items:center;">
<div style="font-size:11px; color:var(--text-color) !important; opacity:0.9;">Đợi: <b>{len(m_wait)} lệnh</b></div>
<div style="font-size:11px; color:#3b82f6 !important; font-weight:700; background:rgba(59, 130, 246, 0.1); padding:2px 6px; border-radius:4px;">Rảnh: {free_date_text}</div>
</div>{note_html}</div>""", unsafe_allow_html=True)

st.markdown("---")

# -------------------------------------------------------------
# 11. BẢNG DANH SÁCH CHI TIẾT
# -------------------------------------------------------------
st.markdown(f"##### Danh sách đơn có nguy cơ delay ({n_open} đơn chưa hoàn thành)")

display_df = df_filtered.sort_values(by='Còn (ngày)', ascending=True).copy()
table_cols = ['Còn (ngày)', 'Mức rủi ro', 'Customer', 'Job Order', 'Machine Name', 'PO', 'Hạn áp dụng', 'Status', 'Drawing No.', 'Qty']
avail_cols = [c for c in table_cols if c in display_df.columns]
tbl_out = display_df[avail_cols].rename(columns={'Customer': 'KH', 'Job Order': 'MQL', 'Machine Name': 'Máy', 'Drawing No.': 'Mã BV', 'Qty': 'SL', 'Status': 'Trạng thái'})

if 'Hạn áp dụng' in tbl_out.columns:
    tbl_out['Hạn áp dụng'] = pd.to_datetime(tbl_out['Hạn áp dụng']).dt.strftime('%d/%m/%Y').fillna('—')

styler = tbl_out.style
style_func = getattr(styler, 'map', getattr(styler, 'applymap', None))

def color_risk(v):
    if not isinstance(v, (int, float)): return ''
    if v < 0: return 'color: #ef4444; font-weight: bold;'
    if v < 7: return 'color: #f97316; font-weight: bold;'
    if v <= 14: return 'color: #eab308; font-weight: bold;'
    if v <= 21: return 'color: #3b82f6; font-weight: bold;'
    return 'color: #22c55e; font-weight: bold;'

st.dataframe(style_func(color_risk, subset=['Còn (ngày)']), use_container_width=True, height=450)
