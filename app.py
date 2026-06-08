"""
Capital Gains Classifier - Streamlit Dashboard
================================================
Upload a CapitalGain CSV file and get:
  - Classified Excel report (Short Term / Long Term / Squaring)
  - Interactive charts with download options
  - Summary metrics and detailed tables
"""

import io
import csv
import os
from collections import defaultdict
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


# ─────────────────────────── Page Config ───────────────────────────

st.set_page_config(
    page_title="Capital Gains Classifier",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────── Custom CSS ───────────────────────────

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

    /* Global font */
    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    /* Main header styling */
    .main-header {
        background: linear-gradient(135deg, #0f0c29 0%, #302b63 50%, #24243e 100%);
        padding: 2.5rem 2rem;
        border-radius: 16px;
        margin-bottom: 2rem;
        text-align: center;
        box-shadow: 0 8px 32px rgba(48, 43, 99, 0.3);
    }
    .main-header h1 {
        color: #ffffff;
        font-size: 2.2rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .main-header p {
        color: #b8b5ff;
        font-size: 1.05rem;
        margin: 0.5rem 0 0 0;
        font-weight: 300;
    }

    /* Metric cards */
    .metric-card {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 14px;
        padding: 1.5rem;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0,0,0,0.15);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .metric-card:hover {
        transform: translateY(-3px);
        box-shadow: 0 8px 30px rgba(0,0,0,0.25);
    }
    .metric-card .label {
        color: #8892b0;
        font-size: 0.8rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1.2px;
        margin-bottom: 0.5rem;
    }
    .metric-card .value {
        font-size: 1.6rem;
        font-weight: 700;
        margin: 0;
    }
    .metric-card .sub {
        color: #8892b0;
        font-size: 0.75rem;
        margin-top: 0.3rem;
    }
    .profit { color: #64ffda; }
    .loss { color: #ff6b6b; }
    .neutral { color: #ccd6f6; }

    /* Section headers */
    .section-header {
        font-size: 1.3rem;
        font-weight: 700;
        color: #ccd6f6;
        margin: 2rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid rgba(100, 255, 218, 0.3);
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    /* Upload area */
    .upload-area {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border: 2px dashed rgba(100, 255, 218, 0.3);
        border-radius: 16px;
        padding: 2rem;
        text-align: center;
        margin: 1rem 0;
    }

    /* Chart container */
    .chart-container {
        background: rgba(26, 26, 46, 0.5);
        border: 1px solid rgba(255,255,255,0.06);
        border-radius: 14px;
        padding: 1rem;
        margin-bottom: 1rem;
    }

    /* Table styling */
    .dataframe {
        border-radius: 10px;
        overflow: hidden;
    }

    /* Sidebar styling */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f0c29 0%, #1a1a2e 100%);
    }

    /* Download button styling */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.5rem !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }
    .stDownloadButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4) !important;
    }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 20px;
        font-weight: 600;
    }

    /* Hide default streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────── Helper Functions ───────────────────────────

def detect_fy_from_csv(csv_content):
    import re
    from datetime import datetime
    
    fy_counts = defaultdict(int)
    
    def parse_date(date_str):
        if not date_str:
            return None
        date_str = str(date_str).strip()
        if not date_str or date_str.lower() in ('nan', 'none', ''):
            return None
        # Try explicit formats
        for fmt in (
            '%d-%b-%y', '%d-%b-%Y', '%d %b %Y', '%d %b %y', 
            '%Y-%m-%d', '%d-%m-%Y', '%d-%m-%y', '%d/%m/%Y', '%d/%m/%y', 
            '%b %Y', '%d %B %Y',
            '%Y-%m-%d %H:%M:%S', '%d-%b-%y %H:%M:%S', '%d-%b-%Y %H:%M:%S',
            '%d/%m/%Y %H:%M:%S', '%d/%m/%y %H:%M:%S', '%d-%m-%Y %H:%M:%S',
            '%d-%m-%y %H:%M:%S'
        ):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                pass
        # Fallback parsing logic
        parts = re.split(r'[-/\s]+', date_str)
        year = None
        for p in parts:
            if p.isdigit() and len(p) == 4:
                year = int(p)
                break
        if year:
            months = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
                      'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12}
            month = 6
            for p in parts:
                p_low = p.lower()[:3]
                if p_low in months:
                    month = months[p_low]
                    break
                elif p.isdigit() and 1 <= int(p) <= 12:
                    month = int(p)
            try:
                return datetime(year, month, 1)
            except Exception:
                pass
        return None

    reader = csv.DictReader(io.StringIO(csv_content))
    for row in reader:
        # Clean keys to remove potential quotes (e.g. 'scrip_name' vs "'scrip_name'") and whitespace
        row = {k.strip().replace("'", "").replace('"', ''): v for k, v in row.items() if k is not None}
        if not row or 'scrip_name' not in row or not row['scrip_name']:
            continue
        dt = parse_date(row.get('SellDate')) or parse_date(row.get('BuyDate'))
        if dt:
            fy = f"{dt.year}-{str(dt.year+1)[-2:]}" if dt.month >= 4 else f"{dt.year-1}-{str(dt.year)[-2:]}"
            fy_counts[fy] += 1
            
    if fy_counts:
        return max(fy_counts, key=fy_counts.get)
    return "2024-25"


def read_and_classify(csv_content):
    """Read CSV content and aggregate into Short Term, Long Term, and Squaring buckets."""
    short_term = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})
    long_term = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})
    squaring = defaultdict(lambda: {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0})

    cl_code = None
    cl_name = None

    # Also collect raw rows for detailed analysis
    all_rows = []

    reader = csv.DictReader(io.StringIO(csv_content))
    for row in reader:
        # Clean keys to remove potential quotes (e.g. 'scrip_name' vs "'scrip_name'") and whitespace
        row = {k.strip().replace("'", "").replace('"', ''): v for k, v in row.items() if k is not None}
        if not row or 'scrip_name' not in row or not row['scrip_name']:
            continue
        
        if cl_code is None:
            cl_code = row.get('ClCode', '').strip()
            cl_name = row.get('ClName', '').strip()

        name = row['scrip_name'].strip()
        stp = float(row['ShortTermProfit'])
        ltp = float(row['LongTermProfit'])
        sqp = float(row['SquringProfit'])

        buy_qty = float(row['BuyQty'])
        buy_val = float(row['BuyValue'])
        sell_qty = float(row['sellQty'])
        sell_val = float(row['SellValue'])

        all_rows.append(row)

        if stp != 0:
            short_term[name]['BuyQty'] += buy_qty
            short_term[name]['BuyValue'] += buy_val
            short_term[name]['SellQty'] += sell_qty
            short_term[name]['SellValue'] += sell_val
            short_term[name]['Profit'] += stp

        if ltp != 0:
            long_term[name]['BuyQty'] += buy_qty
            long_term[name]['BuyValue'] += buy_val
            long_term[name]['SellQty'] += sell_qty
            long_term[name]['SellValue'] += sell_val
            long_term[name]['Profit'] += ltp

        if sqp != 0:
            squaring[name]['BuyQty'] += buy_qty
            squaring[name]['BuyValue'] += buy_val
            squaring[name]['SellQty'] += sell_qty
            squaring[name]['SellValue'] += sell_val
            squaring[name]['Profit'] += sqp

    return short_term, long_term, squaring, cl_code, cl_name, all_rows


def compute_grand_total(data):
    """Compute grand total across all scrips."""
    total = {'BuyQty': 0, 'BuyValue': 0, 'SellQty': 0, 'SellValue': 0, 'Profit': 0}
    for vals in data.values():
        total['BuyQty'] += vals['BuyQty']
        total['BuyValue'] += vals['BuyValue']
        total['SellQty'] += vals['SellQty']
        total['SellValue'] += vals['SellValue']
        total['Profit'] += vals['Profit']
    return total


def dict_to_df(data, profit_col_name):
    """Convert classified dict to a pandas DataFrame."""
    rows = []
    for name in sorted(data.keys()):
        vals = data[name]
        rows.append({
            'Scrip Name': name,
            'Buy Qty': int(vals['BuyQty']),
            'Buy Value': round(vals['BuyValue'], 2),
            'Sell Qty': int(vals['SellQty']),
            'Sell Value': round(vals['SellValue'], 2),
            profit_col_name: round(vals['Profit'], 2),
        })
    return pd.DataFrame(rows)


def generate_excel(short_term, long_term, squaring, cl_code, cl_name, fy_year):
    """Generate Excel report and return as bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Capital Gains Report"

    header_font = Font(name='Calibri', bold=True, size=11)
    section_font = Font(name='Calibri', bold=True, size=12, color='1F4E79')
    title_font = Font(name='Calibri', bold=True, size=11, color='1F4E79')
    data_font = Font(name='Calibri', size=11)
    total_font = Font(name='Calibri', bold=True, size=11)
    number_format = '#,##0.00'
    qty_format = '#,##0'

    header_fill = PatternFill(start_color='D6E4F0', end_color='D6E4F0', fill_type='solid')
    total_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 20
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 22

    current_row = 2
    ws.cell(row=current_row, column=2, value=cl_code).font = title_font
    current_row += 1
    ws.cell(row=current_row, column=2, value=cl_name).font = title_font
    ws.cell(row=current_row, column=7, value=fy_year).font = title_font
    current_row += 1

    def write_section(section_name, data, profit_label, start_row):
        row = start_row
        ws.cell(row=row, column=2, value=section_name).font = section_font
        row += 1

        headers = ['Scrip Name', 'Sum of BuyQty', 'Sum of BuyValue',
                    'Sum of sellQty', 'Sum of SellValue', profit_label]
        for col_idx, header in enumerate(headers, start=2):
            cell = ws.cell(row=row, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        row += 1

        for name in sorted(data.keys()):
            vals = data[name]
            ws.cell(row=row, column=2, value=name).font = data_font
            ws.cell(row=row, column=2).border = thin_border

            for col, val, fmt in [
                (3, round(vals['BuyQty']), qty_format),
                (4, round(vals['BuyValue'], 2), number_format),
                (5, round(vals['SellQty']), qty_format),
                (6, round(vals['SellValue'], 2), number_format),
                (7, round(vals['Profit'], 2), number_format),
            ]:
                cell = ws.cell(row=row, column=col, value=val)
                cell.font = data_font
                cell.number_format = fmt
                cell.border = thin_border

            profit_cell = ws.cell(row=row, column=7)
            if vals['Profit'] >= 0:
                profit_cell.font = Font(name='Calibri', size=11, color='006100')
            else:
                profit_cell.font = Font(name='Calibri', size=11, color='9C0006')
            row += 1

        totals = compute_grand_total(data)
        ws.cell(row=row, column=2, value='Grand Total').font = total_font
        ws.cell(row=row, column=2).fill = total_fill
        ws.cell(row=row, column=2).border = thin_border

        for col, val, fmt in [
            (3, round(totals['BuyQty']), qty_format),
            (4, round(totals['BuyValue'], 2), number_format),
            (5, round(totals['SellQty']), qty_format),
            (6, round(totals['SellValue'], 2), number_format),
            (7, round(totals['Profit'], 2), number_format),
        ]:
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = total_font
            cell.number_format = fmt
            cell.fill = total_fill
            cell.border = thin_border

        return row + 2

    current_row = write_section('Short Term', short_term, 'Sum of ShortTermProfit', current_row)
    current_row = write_section('Long Term', long_term, 'Sum of LongTermProfit', current_row)
    current_row = write_section('Squaring', squaring, 'Sum of SquringProfit', current_row)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def format_inr(value):
    """Format a number in Indian currency style."""
    if value < 0:
        return f"-{format_inr(-value)}"
    s = f"{value:,.2f}"
    return f"Rs. {s}"


def get_chart_download(fig, filename):
    """Generate a PNG download from a plotly figure."""
    try:
        img_bytes = fig.to_image(format="png", width=1200, height=600, scale=2)
        return img_bytes
    except Exception:
        return None


# ─────────────────────────── Plotly Chart Theme ───────────────────────────

CHART_BASE = dict(
    paper_bgcolor='rgba(0,0,0,0)',
    plot_bgcolor='rgba(0,0,0,0)',
    font=dict(family='Inter, sans-serif', color='#ccd6f6', size=13),
    title_font=dict(size=18, color='#ccd6f6', family='Inter, sans-serif'),
    xaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)'),
    yaxis=dict(gridcolor='rgba(255,255,255,0.06)', zerolinecolor='rgba(255,255,255,0.1)'),
    legend=dict(bgcolor='rgba(0,0,0,0)', font=dict(color='#8892b0')),
    margin=dict(l=60, r=30, t=60, b=60),
)


def apply_theme(fig, **overrides):
    """Apply the dark chart theme to a figure with optional overrides.
    
    First applies CHART_BASE, then applies overrides separately
    so Plotly merges them correctly without duplicate keyword issues.
    """
    fig.update_layout(**CHART_BASE)
    if overrides:
        fig.update_layout(**overrides)
    return fig


COLORS = {
    'profit': '#64ffda',
    'loss': '#ff6b6b',
    'short_term': '#667eea',
    'long_term': '#64ffda',
    'squaring': '#f093fb',
    'gradient': ['#667eea', '#764ba2', '#f093fb', '#64ffda', '#4facfe',
                 '#43e97b', '#fa709a', '#fee140', '#a18cd1', '#fbc2eb'],
}


# ─────────────────────────── Main App ───────────────────────────

def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <h1>📊 Capital Gains Classifier</h1>
        <p>Upload your trading data CSV and get instant classified reports with interactive analytics</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### 📁 Upload Data")
        st.markdown("Upload your **CapitalGain CSV** file exported from your broker.")

        uploaded_file = st.file_uploader(
            "Choose a CSV or Excel file",
            type=['csv', 'xlsx'],
            help="Upload the CapitalGain CSV or Excel file",
            label_visibility="collapsed",
        )

        if uploaded_file:
            st.success(f"Loaded: **{uploaded_file.name}**")
            
            # Detect FY from content (supports both CSV and Excel)
            try:
                if uploaded_file.name.lower().endswith('.xlsx'):
                    df = pd.read_excel(uploaded_file)
                    content_str = df.to_csv(index=False)
                else:
                    file_bytes = uploaded_file.getvalue()
                    content_str = file_bytes.decode('utf-8-sig')
                detected_fy = detect_fy_from_csv(content_str)
            except Exception:
                detected_fy = "2024-25"
                
            st.markdown("---")
            st.markdown("### Settings")
            fy_year = st.text_input("Financial Year", value=detected_fy, help="e.g., 2024-25")
        else:
            fy_year = "2024-25"

        pass

    # ─── No file uploaded ───
    if not uploaded_file:
        st.markdown("""
        <div class="upload-area">
            <h3 style="color: #ccd6f6; margin-bottom: 0.5rem;">👈 Upload your CSV to get started</h3>
            <p style="color: #8892b0;">
                Use the sidebar to upload your CapitalGain CSV file.<br>
                The app will classify your trades into <b>Short Term</b>, <b>Long Term</b>, and <b>Squaring (Intraday)</b>.
            </p>
        </div>
        """, unsafe_allow_html=True)

        # Show sample format
        with st.expander("📋 Expected CSV Format"):
            st.code(
                "scrip_code,isin,scrip_name,BuyDate,BuyQty,BuyRate,BuyValue,"
                "Descer,SellDate,sellQty,SellRate,SellValue,"
                "ShortTermProfit,LongTermProfit,SquringProfit,Fair_rate,ClCode,ClName",
                language="text",
            )
            st.info("The CSV should contain the columns above. The classification is based on "
                     "the ShortTermProfit, LongTermProfit, and SquringProfit columns.")
        return

    # ─── Process uploaded file (supports both CSV and Excel) ───
    if uploaded_file.name.lower().endswith('.xlsx'):
        df = pd.read_excel(uploaded_file)
        csv_content = df.to_csv(index=False)
    else:
        # Use getvalue() to avoid file pointer issues if read in sidebar
        csv_content = uploaded_file.getvalue().decode('utf-8-sig')
        
    short_term, long_term, squaring, cl_code, cl_name, all_rows = read_and_classify(csv_content)

    if not all_rows:
        st.warning("⚠️ No valid transactions found in the uploaded file. Please make sure you are uploading the raw Capital Gains statement from your broker, not an already classified report.")
        st.info("Expected columns: `scrip_code`, `isin`, `scrip_name`, `BuyDate`, `BuyQty`, `BuyRate`, `BuyValue`, `Descer`, `SellDate`, `sellQty`, `SellRate`, `SellValue`, `ShortTermProfit`, `LongTermProfit`, `SquringProfit`, `Fair_rate`, `ClCode`, `ClName`")
        return

    st_total = compute_grand_total(short_term)
    lt_total = compute_grand_total(long_term)
    sq_total = compute_grand_total(squaring)
    overall_pnl = st_total['Profit'] + lt_total['Profit'] + sq_total['Profit']

    # ─── Client Info Bar ───
    st.markdown(
        f"<div style='background: linear-gradient(90deg, #1a1a2e, #16213e); padding: 1rem 1.5rem; "
        f"border-radius: 12px; margin-bottom: 1.5rem; display: flex; justify-content: space-between; "
        f"align-items: center; border: 1px solid rgba(255,255,255,0.06);'>"
        f"<div><span style='color:#8892b0; font-size:0.8rem;'>CLIENT</span><br>"
        f"<span style='color:#ccd6f6; font-size:1.1rem; font-weight:600;'>{cl_name}</span>"
        f"<span style='color:#8892b0; font-size:0.85rem;'> ({cl_code})</span></div>"
        f"<div><span style='color:#8892b0; font-size:0.8rem;'>FINANCIAL YEAR</span><br>"
        f"<span style='color:#64ffda; font-size:1.1rem; font-weight:600;'>{fy_year}</span></div>"
        f"<div><span style='color:#8892b0; font-size:0.8rem;'>TOTAL TRANSACTIONS</span><br>"
        f"<span style='color:#ccd6f6; font-size:1.1rem; font-weight:600;'>{len(all_rows)}</span></div>"
        f"</div>",
        unsafe_allow_html=True,
    )

    # ─── KPI Metrics ───
    m1, m2, m3, m4 = st.columns(4)

    with m1:
        css_class = 'profit' if overall_pnl >= 0 else 'loss'
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">Total P&L</div>
            <div class="value {css_class}">{format_inr(overall_pnl)}</div>
            <div class="sub">{len(short_term) + len(long_term) + len(squaring)} unique scrips</div>
        </div>""", unsafe_allow_html=True)

    with m2:
        css_class = 'profit' if st_total['Profit'] >= 0 else 'loss'
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">Short Term P&L</div>
            <div class="value {css_class}">{format_inr(st_total['Profit'])}</div>
            <div class="sub">{len(short_term)} scrips</div>
        </div>""", unsafe_allow_html=True)

    with m3:
        css_class = 'profit' if lt_total['Profit'] >= 0 else 'loss'
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">Long Term P&L</div>
            <div class="value {css_class}">{format_inr(lt_total['Profit'])}</div>
            <div class="sub">{len(long_term)} scrips</div>
        </div>""", unsafe_allow_html=True)

    with m4:
        css_class = 'profit' if sq_total['Profit'] >= 0 else 'loss'
        st.markdown(f"""
        <div class="metric-card">
            <div class="label">Squaring P&L</div>
            <div class="value {css_class}">{format_inr(sq_total['Profit'])}</div>
            <div class="sub">{len(squaring)} scrips</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── Download Excel Report ───
    excel_bytes = generate_excel(short_term, long_term, squaring, cl_code, cl_name, fy_year)

    # Generate dynamic filename matching user format: e.g. BSWS024_CapitalGain_2025_26_.xlsx
    if cl_code:
        clean_cl_code = cl_code.split('-')[0].strip()
        formatted_fy = fy_year.replace('-', '_')
        download_filename = f"{clean_cl_code}_CapitalGain_{formatted_fy}_.xlsx"
    else:
        formatted_fy = fy_year.replace('-', '_')
        download_filename = f"CapitalGain_{formatted_fy}_Classified.xlsx"

    dl1, dl2, _ = st.columns([1, 1, 2])
    with dl1:
        st.download_button(
            label="📥 Download Classified Excel Report",
            data=excel_bytes,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ─── Charts Section ───
    st.markdown('<div class="section-header">📈 Interactive Analytics</div>', unsafe_allow_html=True)

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Overview", "📊 Short Term", "📊 Long Term", "📊 Squaring", "🔍 Detailed Data"
    ])

    # ────────── TAB 1: Overview ──────────
    with tab1:
        col_left, col_right = st.columns(2)

        with col_left:
            # P&L by Category - Donut Chart
            fig_donut = go.Figure()

            categories = []
            values = []
            colors_list = []

            if st_total['Profit'] != 0:
                categories.append('Short Term')
                values.append(abs(st_total['Profit']))
                colors_list.append(COLORS['short_term'])
            if lt_total['Profit'] != 0:
                categories.append('Long Term')
                values.append(abs(lt_total['Profit']))
                colors_list.append(COLORS['long_term'])
            if sq_total['Profit'] != 0:
                categories.append('Squaring')
                values.append(abs(sq_total['Profit']))
                colors_list.append(COLORS['squaring'])

            fig_donut = go.Figure(data=[go.Pie(
                labels=categories,
                values=values,
                hole=0.55,
                marker_colors=colors_list,
                textinfo='label+percent',
                textfont=dict(size=13, color='white'),
                hovertemplate='<b>%{label}</b><br>Amount: Rs. %{value:,.2f}<br>Share: %{percent}<extra></extra>',
            )])
            apply_theme(fig_donut,
                title='P&L Distribution by Category',
                showlegend=True,
                height=420,
                annotations=[dict(
                    text=f'<b>Total</b><br>Rs. {overall_pnl:,.0f}',
                    x=0.5, y=0.5, font_size=14, showarrow=False,
                    font_color='#ccd6f6',
                )],
            )
            st.plotly_chart(fig_donut, width='stretch', key="donut_overview")

            img = get_chart_download(fig_donut, "pnl_distribution")
            if img:
                st.download_button("📥 Download Chart", img, "pnl_distribution.png", "image/png", key="dl_donut")

        with col_right:
            # P&L by Category - Bar Chart
            pnl_data = pd.DataFrame({
                'Category': ['Short Term', 'Long Term', 'Squaring'],
                'Profit/Loss': [st_total['Profit'], lt_total['Profit'], sq_total['Profit']],
                'Color': [
                    COLORS['profit'] if st_total['Profit'] >= 0 else COLORS['loss'],
                    COLORS['profit'] if lt_total['Profit'] >= 0 else COLORS['loss'],
                    COLORS['profit'] if sq_total['Profit'] >= 0 else COLORS['loss'],
                ]
            })

            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                x=pnl_data['Category'],
                y=pnl_data['Profit/Loss'],
                marker_color=pnl_data['Color'],
                text=[f"Rs. {v:,.0f}" for v in pnl_data['Profit/Loss']],
                textposition='outside',
                textfont=dict(size=12, color='#ccd6f6'),
                hovertemplate='<b>%{x}</b><br>P&L: Rs. %{y:,.2f}<extra></extra>',
            ))
            apply_theme(fig_bar,
                title='Profit/Loss by Category',
                yaxis_title='Amount (Rs.)',
                height=420,
                showlegend=False,
            )
            st.plotly_chart(fig_bar, width='stretch', key="bar_overview")

            img = get_chart_download(fig_bar, "pnl_by_category")
            if img:
                st.download_button("📥 Download Chart", img, "pnl_by_category.png", "image/png", key="dl_bar_cat")

        # ─── Turnover Chart ───
        st.markdown("<br>", unsafe_allow_html=True)

        # Buy/Sell volume comparison
        turnover_data = pd.DataFrame({
            'Category': ['Short Term', 'Short Term', 'Long Term', 'Long Term', 'Squaring', 'Squaring'],
            'Type': ['Buy Value', 'Sell Value', 'Buy Value', 'Sell Value', 'Buy Value', 'Sell Value'],
            'Amount': [
                st_total['BuyValue'], st_total['SellValue'],
                lt_total['BuyValue'], lt_total['SellValue'],
                sq_total['BuyValue'], sq_total['SellValue'],
            ]
        })

        fig_turnover = px.bar(
            turnover_data, x='Category', y='Amount', color='Type',
            barmode='group',
            color_discrete_map={'Buy Value': '#667eea', 'Sell Value': '#64ffda'},
            title='Buy vs Sell Turnover by Category',
        )
        apply_theme(fig_turnover,
            yaxis_title='Amount (Rs.)',
            height=400,
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1, bgcolor='rgba(0,0,0,0)', font=dict(color='#8892b0')),
        )
        fig_turnover.update_traces(
            hovertemplate='<b>%{x}</b><br>%{data.name}: Rs. %{y:,.2f}<extra></extra>'
        )
        st.plotly_chart(fig_turnover, width='stretch', key="turnover")

        img = get_chart_download(fig_turnover, "turnover")
        if img:
            st.download_button("📥 Download Chart", img, "turnover_comparison.png", "image/png", key="dl_turnover")

    # ────────── TAB 2: Short Term ──────────
    with tab2:
        if not short_term:
            st.info("No Short Term transactions found.")
        else:
            df_st = dict_to_df(short_term, 'Short Term Profit')

            # Horizontal bar chart - Top gainers/losers
            df_st_sorted = df_st.sort_values('Short Term Profit', ascending=True)
            colors = [COLORS['profit'] if v >= 0 else COLORS['loss'] for v in df_st_sorted['Short Term Profit']]

            fig_st = go.Figure()
            fig_st.add_trace(go.Bar(
                y=df_st_sorted['Scrip Name'],
                x=df_st_sorted['Short Term Profit'],
                orientation='h',
                marker_color=colors,
                text=[f"Rs. {v:,.0f}" for v in df_st_sorted['Short Term Profit']],
                textposition='outside',
                textfont=dict(size=11, color='#ccd6f6'),
                hovertemplate='<b>%{y}</b><br>P&L: Rs. %{x:,.2f}<extra></extra>',
            ))
            apply_theme(fig_st,
                title=f'Short Term Profit/Loss by Scrip ({len(short_term)} scrips)',
                xaxis_title='Profit/Loss (Rs.)',
                height=max(400, len(short_term) * 35),
                yaxis=dict(
                    gridcolor='rgba(255,255,255,0.06)',
                    tickfont=dict(size=11),
                ),
            )
            st.plotly_chart(fig_st, width='stretch', key="st_bar")

            img = get_chart_download(fig_st, "short_term_pnl")
            if img:
                st.download_button("📥 Download Chart", img, "short_term_pnl.png", "image/png", key="dl_st")

            # Winners vs Losers Pie
            winners = sum(1 for v in short_term.values() if v['Profit'] >= 0)
            losers = len(short_term) - winners

            col_a, col_b = st.columns(2)
            with col_a:
                fig_wl = go.Figure(data=[go.Pie(
                    labels=['Winners', 'Losers'],
                    values=[winners, losers],
                    marker_colors=[COLORS['profit'], COLORS['loss']],
                    hole=0.5,
                    textinfo='label+value',
                    textfont=dict(size=14, color='white'),
                )])
                apply_theme(fig_wl,
                    title='Win/Loss Ratio (Short Term)',
                    height=350,
                )
                st.plotly_chart(fig_wl, width='stretch', key="st_pie")

            with col_b:
                # Volume treemap
                df_vol = df_st.copy()
                df_vol['Abs Buy Value'] = df_vol['Buy Value'].abs()
                fig_tree = px.treemap(
                    df_vol, path=['Scrip Name'], values='Abs Buy Value',
                    color='Short Term Profit',
                    color_continuous_scale=['#ff6b6b', '#1a1a2e', '#64ffda'],
                    color_continuous_midpoint=0,
                    title='Investment Volume (size) & P&L (color)',
                )
                apply_theme(fig_tree,
                    height=350,
                )
                st.plotly_chart(fig_tree, width='stretch', key="st_tree")

            # Data table
            st.markdown("**Detailed Short Term Data**")
            st.dataframe(
                df_st.style.format({
                    'Buy Value': 'Rs. {:,.2f}',
                    'Sell Value': 'Rs. {:,.2f}',
                    'Short Term Profit': 'Rs. {:,.2f}',
                }).map(
                    lambda v: 'color: #64ffda' if isinstance(v, (int, float)) and v > 0
                    else ('color: #ff6b6b' if isinstance(v, (int, float)) and v < 0 else ''),
                    subset=['Short Term Profit']
                ),
                width='stretch',
                height=min(400, len(df_st) * 40 + 40),
            )

    # ────────── TAB 3: Long Term ──────────
    with tab3:
        if not long_term:
            st.info("No Long Term transactions found.")
        else:
            df_lt = dict_to_df(long_term, 'Long Term Profit')

            df_lt_sorted = df_lt.sort_values('Long Term Profit', ascending=True)
            colors = [COLORS['profit'] if v >= 0 else COLORS['loss'] for v in df_lt_sorted['Long Term Profit']]

            fig_lt = go.Figure()
            fig_lt.add_trace(go.Bar(
                y=df_lt_sorted['Scrip Name'],
                x=df_lt_sorted['Long Term Profit'],
                orientation='h',
                marker_color=colors,
                text=[f"Rs. {v:,.0f}" for v in df_lt_sorted['Long Term Profit']],
                textposition='outside',
                textfont=dict(size=12, color='#ccd6f6'),
                hovertemplate='<b>%{y}</b><br>P&L: Rs. %{x:,.2f}<extra></extra>',
            ))
            apply_theme(fig_lt,
                title=f'Long Term Profit/Loss by Scrip ({len(long_term)} scrips)',
                xaxis_title='Profit/Loss (Rs.)',
                height=max(350, len(long_term) * 50),
            )
            st.plotly_chart(fig_lt, width='stretch', key="lt_bar")

            img = get_chart_download(fig_lt, "long_term_pnl")
            if img:
                st.download_button("📥 Download Chart", img, "long_term_pnl.png", "image/png", key="dl_lt")

            # Buy vs Sell comparison
            if len(df_lt) > 1:
                col_a, col_b = st.columns(2)
                with col_a:
                    winners = sum(1 for v in long_term.values() if v['Profit'] >= 0)
                    losers = len(long_term) - winners
                    fig_wl = go.Figure(data=[go.Pie(
                        labels=['Winners', 'Losers'],
                        values=[winners, losers],
                        marker_colors=[COLORS['profit'], COLORS['loss']],
                        hole=0.5,
                        textinfo='label+value',
                        textfont=dict(size=14, color='white'),
                    )])
                    apply_theme(fig_wl,
                        title='Win/Loss Ratio (Long Term)',
                        height=350,
                    )
                    st.plotly_chart(fig_wl, width='stretch', key="lt_pie")

            st.markdown("**Detailed Long Term Data**")
            st.dataframe(
                df_lt.style.format({
                    'Buy Value': 'Rs. {:,.2f}',
                    'Sell Value': 'Rs. {:,.2f}',
                    'Long Term Profit': 'Rs. {:,.2f}',
                }).map(
                    lambda v: 'color: #64ffda' if isinstance(v, (int, float)) and v > 0
                    else ('color: #ff6b6b' if isinstance(v, (int, float)) and v < 0 else ''),
                    subset=['Long Term Profit']
                ),
                width='stretch',
                height=min(400, len(df_lt) * 40 + 40),
            )

    # ────────── TAB 4: Squaring ──────────
    with tab4:
        if not squaring:
            st.info("No Squaring/Intraday transactions found.")
        else:
            df_sq = dict_to_df(squaring, 'Squaring Profit')

            df_sq_sorted = df_sq.sort_values('Squaring Profit', ascending=True)
            colors = [COLORS['profit'] if v >= 0 else COLORS['loss'] for v in df_sq_sorted['Squaring Profit']]

            fig_sq = go.Figure()
            fig_sq.add_trace(go.Bar(
                y=df_sq_sorted['Scrip Name'],
                x=df_sq_sorted['Squaring Profit'],
                orientation='h',
                marker_color=colors,
                text=[f"Rs. {v:,.0f}" for v in df_sq_sorted['Squaring Profit']],
                textposition='outside',
                textfont=dict(size=11, color='#ccd6f6'),
                hovertemplate='<b>%{y}</b><br>P&L: Rs. %{x:,.2f}<extra></extra>',
            ))
            apply_theme(fig_sq,
                title=f'Squaring Profit/Loss by Scrip ({len(squaring)} scrips)',
                xaxis_title='Profit/Loss (Rs.)',
                height=max(400, len(squaring) * 35),
                yaxis=dict(
                    gridcolor='rgba(255,255,255,0.06)',
                    tickfont=dict(size=11),
                ),
            )
            st.plotly_chart(fig_sq, width='stretch', key="sq_bar")

            img = get_chart_download(fig_sq, "squaring_pnl")
            if img:
                st.download_button("📥 Download Chart", img, "squaring_pnl.png", "image/png", key="dl_sq")

            col_a, col_b = st.columns(2)
            with col_a:
                winners = sum(1 for v in squaring.values() if v['Profit'] >= 0)
                losers = len(squaring) - winners
                fig_wl = go.Figure(data=[go.Pie(
                    labels=['Winners', 'Losers'],
                    values=[winners, losers],
                    marker_colors=[COLORS['profit'], COLORS['loss']],
                    hole=0.5,
                    textinfo='label+value',
                    textfont=dict(size=14, color='white'),
                )])
                apply_theme(fig_wl,
                    title='Win/Loss Ratio (Squaring)',
                    height=350,
                )
                st.plotly_chart(fig_wl, width='stretch', key="sq_pie")

            with col_b:
                df_vol = df_sq.copy()
                df_vol['Abs Buy Value'] = df_vol['Buy Value'].abs()
                fig_tree = px.treemap(
                    df_vol, path=['Scrip Name'], values='Abs Buy Value',
                    color='Squaring Profit',
                    color_continuous_scale=['#ff6b6b', '#1a1a2e', '#64ffda'],
                    color_continuous_midpoint=0,
                    title='Trading Volume (size) & P&L (color)',
                )
                apply_theme(fig_tree,
                    height=350,
                )
                st.plotly_chart(fig_tree, width='stretch', key="sq_tree")

            st.markdown("**Detailed Squaring Data**")
            st.dataframe(
                df_sq.style.format({
                    'Buy Value': 'Rs. {:,.2f}',
                    'Sell Value': 'Rs. {:,.2f}',
                    'Squaring Profit': 'Rs. {:,.2f}',
                }).map(
                    lambda v: 'color: #64ffda' if isinstance(v, (int, float)) and v > 0
                    else ('color: #ff6b6b' if isinstance(v, (int, float)) and v < 0 else ''),
                    subset=['Squaring Profit']
                ),
                width='stretch',
                height=min(400, len(df_sq) * 40 + 40),
            )

    # ────────── TAB 5: Detailed Raw Data ──────────
    with tab5:
        st.markdown("**Complete Transaction Data**")

        # Convert all_rows to DataFrame
        df_raw = pd.DataFrame(all_rows)

        # Filters
        col_f1, col_f2, col_f3 = st.columns(3)
        with col_f1:
            scrip_filter = st.multiselect(
                "Filter by Scrip",
                options=sorted(df_raw['scrip_name'].unique()),
                default=None,
                key="scrip_filter",
            )
        with col_f2:
            type_filter = st.multiselect(
                "Filter by Type (Descer)",
                options=[d for d in sorted(df_raw['Descer'].unique()) if d],
                default=None,
                key="type_filter",
            )
        with col_f3:
            pnl_filter = st.radio(
                "P&L Filter",
                options=["All", "Profitable Only", "Loss Only"],
                horizontal=True,
                key="pnl_filter",
            )

        df_display = df_raw.copy()
        if scrip_filter:
            df_display = df_display[df_display['scrip_name'].isin(scrip_filter)]
        if type_filter:
            df_display = df_display[df_display['Descer'].isin(type_filter)]

        # Convert numeric columns
        numeric_cols = ['BuyQty', 'BuyValue', 'sellQty', 'SellValue',
                        'ShortTermProfit', 'LongTermProfit', 'SquringProfit']
        for col in numeric_cols:
            df_display[col] = pd.to_numeric(df_display[col], errors='coerce')

        if pnl_filter == "Profitable Only":
            df_display = df_display[
                (df_display['ShortTermProfit'] > 0) |
                (df_display['LongTermProfit'] > 0) |
                (df_display['SquringProfit'] > 0)
            ]
        elif pnl_filter == "Loss Only":
            df_display = df_display[
                (df_display['ShortTermProfit'] < 0) |
                (df_display['LongTermProfit'] < 0) |
                (df_display['SquringProfit'] < 0)
            ]

        st.markdown(f"Showing **{len(df_display)}** of {len(df_raw)} transactions")
        st.dataframe(df_display, use_container_width=True, height=500)

        # Download filtered data as CSV
        csv_download = df_display.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Download Filtered Data as CSV",
            csv_download,
            f"filtered_transactions_{fy_year}.csv",
            "text/csv",
            key="dl_csv_raw",
        )


if __name__ == '__main__':
    main()
