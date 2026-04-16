"""
SRKK Document Intelligence
Streamlit UI for OCR, Extraction & Bank Matching Demo
"""

import streamlit as st
import json
import sys
import base64
import tempfile
import pandas as pd
import random
import hashlib
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime, date, timedelta

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).resolve().parent))

from core.pdf_to_images import pdf_to_images
from core.ocr_agent import ocr_images_with_chat_model, ocr_image_with_chat_model, _maybe_parse_json
from core.orchestrator import run as orchestrator_run, AGENT_REGISTRY
from agents.classifier import classify_document

# ─── Reconciliation Summary — colour palette & mock data ────────────────────
_RC_PURPLE      = "#BFB2F9"
_RC_PINK        = "#FD8FD5"
_RC_RED         = "#DB3C36"
_RC_TEAL        = "#00A0AF"
_RC_PURPLE_DARK = "#8B7AD6"
_RC_PINK_DARK   = "#D96FB0"
_RC_RED_DARK    = "#A82D28"
_RC_TEAL_DARK   = "#007A85"
_RC_BG          = "#F8F9FC"
_RC_CARD        = "#FFFFFF"
_RC_TEXT        = "#1A1A2E"
_RC_MUTED       = "#6B7280"
_RC_BORDER      = "#E5E7EB"
_RC_AMBER       = "#F59E0B"
_RC_GREEN       = "#10B981"
_RC_PLOTLY_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter, system-ui, sans-serif", color="#1A1A2E"),
)


@st.cache_data
def _generate_ocr_summary_mock() -> pd.DataFrame:
    """Mock: monthly OCR document processing summary."""
    rng = random.Random(7)
    doc_types = [
        "Commercial Invoice", "Credit Note", "Travel", "Rental",
        "Hotel", "Utility", "SOA", "Bank Statement",
        "SRKK - Vendor Invoice", "SRKK - Purchase Order", "SRKK - Microsoft Billing",
    ]
    rows = []
    base = date(2025, 9, 1)
    for m in range(7):
        month_date = base + timedelta(days=m * 30)
        month_label = month_date.strftime("%b %Y")
        for dt in doc_types:
            docs = rng.randint(5, 80)
            pages = docs * rng.randint(1, 6)
            success = rng.randint(int(docs * 0.85), docs)
            failed = docs - success
            rows.append({
                "Month": month_label,
                "Document Type": dt,
                "Documents Processed": docs,
                "Pages Processed": pages,
                "Successful": success,
                "Failed": failed,
            })
    return pd.DataFrame(rows)


@st.cache_data
def _generate_ocr_doc_type_mock() -> pd.DataFrame:
    """Mock: average pages uploaded and average fields extracted per document type."""
    return pd.DataFrame([
        {"Document Type": "Commercial Invoice", "Avg Pages Uploaded": 2, "Avg Fields Extracted": 18},
        {"Document Type": "Credit Note",        "Avg Pages Uploaded": 2, "Avg Fields Extracted": 14},
        {"Document Type": "Travel",             "Avg Pages Uploaded": 3, "Avg Fields Extracted": 22},
        {"Document Type": "Rental",             "Avg Pages Uploaded": 2, "Avg Fields Extracted": 16},
        {"Document Type": "Hotel",              "Avg Pages Uploaded": 3, "Avg Fields Extracted": 20},
        {"Document Type": "Utility",            "Avg Pages Uploaded": 2, "Avg Fields Extracted": 12},
        {"Document Type": "SOA",                        "Avg Pages Uploaded": 5, "Avg Fields Extracted": 30},
        {"Document Type": "Bank Statement",              "Avg Pages Uploaded": 5, "Avg Fields Extracted": 28},
        {"Document Type": "SRKK - Vendor Invoice",       "Avg Pages Uploaded": 2, "Avg Fields Extracted": 20},
        {"Document Type": "SRKK - Purchase Order",       "Avg Pages Uploaded": 2, "Avg Fields Extracted": 18},
        {"Document Type": "SRKK - Microsoft Billing",    "Avg Pages Uploaded": 3, "Avg Fields Extracted": 25},
    ])


@st.cache_data
def _generate_recon_mock_runs() -> pd.DataFrame:
    """Generate realistic mock reconciliation run history for dashboard display."""
    rng = random.Random(42)
    runs = []
    base_dt = date(2025, 10, 1)
    for i in range(25):
        run_date   = base_dt + timedelta(days=i * 7 + rng.randint(0, 3))
        period_from = run_date - timedelta(days=7)
        period_to   = run_date - timedelta(days=1)
        income_rows  = rng.randint(800, 5000)
        balance_rows = rng.randint(800, 5000)
        sales_rows   = rng.randint(200, 1500)
        recon_rows   = rng.randint(700, min(income_rows, balance_rows))
        if i % 5 == 0:
            outstanding_rows  = 0
            refund_rows       = rng.randint(0, int(recon_rows * 0.08))
            match_rate        = 100.0
            total_sales       = round(rng.uniform(50_000, 500_000), 2)
            total_fees        = round(total_sales * rng.uniform(0.03, 0.12), 2)
            total_payment     = round(total_sales - total_fees, 2)
            total_outstanding = 0.0
        else:
            outstanding_rows  = rng.randint(1, max(2, int(recon_rows * 0.15)))
            refund_rows       = rng.randint(0, int(recon_rows * 0.08))
            match_rate        = round((recon_rows - outstanding_rows) / recon_rows * 100, 2) if recon_rows else 0
            total_sales       = round(rng.uniform(50_000, 500_000), 2)
            total_fees        = round(total_sales * rng.uniform(0.03, 0.12), 2)
            max_payment       = total_sales - total_fees
            total_payment     = round(max_payment * rng.uniform(0.85, 0.97), 2)
            total_outstanding = round(total_sales - total_payment - total_fees, 2)
        income_not_balance = rng.randint(0, 30)
        balance_not_income = rng.randint(0, 25)
        duration           = round(rng.uniform(3.5, 45.0), 2)
        fees_pct           = round(total_fees / total_sales * 100, 2) if total_sales else 0
        needs_review       = outstanding_rows != 0 or total_outstanding != 0
        review_reasons     = []
        if outstanding_rows != 0:
            review_reasons.append(f"Outstanding orders ({outstanding_rows})")
        if total_outstanding != 0:
            review_reasons.append(f"Outstanding amount (RM {total_outstanding:,.2f})")
        run_id = hashlib.sha256(f"run-{i}-{run_date}".encode()).hexdigest()[:8].upper()
        runs.append({
            "Run ID": f"RUN-{run_id}", "Run Date": run_date,
            "Period From": period_from, "Period To": period_to,
            "Income Rows": income_rows, "Balance Rows": balance_rows,
            "Sales Rows": sales_rows, "Recon Rows": recon_rows,
            "Outstanding Orders": outstanding_rows, "Refund Orders": refund_rows,
            "Match Rate (%)": match_rate,
            "Total Sales (RM)": total_sales, "Total Payment (RM)": total_payment,
            "Total Fees (RM)": total_fees, "Total Outstanding (RM)": total_outstanding,
            "Income Not In Balance": income_not_balance, "Balance Not In Income": balance_not_income,
            "Duration (s)": duration, "Fees % of Sales": fees_pct,
            "Needs Review": needs_review,
            "Review Reasons": "; ".join(review_reasons) if review_reasons else "—",
            "Status": "⚠️ Needs Review" if needs_review else "✅ OK",
        })
    return pd.DataFrame(runs)


def _rc_metric_card(label: str, value: str, accent: str = "#00A0AF") -> str:
    return f"""
    <div style="background:#fff;border-radius:12px;padding:1.25rem;border:1px solid #E5E7EB;
                box-shadow:0 1px 3px rgba(0,0,0,0.04);">
        <div style="display:flex;align-items:center;">
            <span style="width:4px;height:32px;border-radius:2px;background:{accent};
                         display:inline-block;margin-right:0.6rem;vertical-align:middle;"></span>
            <div>
                <div style="font-size:0.78rem;font-weight:600;text-transform:uppercase;
                            letter-spacing:0.5px;color:#6B7280;margin-bottom:0.35rem;">{label}</div>
                <div style="font-size:1.75rem;font-weight:700;color:#1A1A2E;line-height:1.2;">{value}</div>
            </div>
        </div>
    </div>
    """

# ─── Page Config ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="SRKK Document Intelligence",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── Custom CSS — Brand Theme ────────────────────────────────────────
# Palette:  Emerald #065f46 · #059669 · #047857 · #6ee7b7 · #34d399
#           Red  #EE2D25 · White #FFFFFF
st.markdown("""
<style>
    /* ═══ GLOBAL FOUNDATION ═══ */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { padding-top: 1rem; }

    /* ═══ MAIN HEADER BAR ═══ */
    .main-header {
        background: linear-gradient(135deg, #022c22 0%, #065f46 30%, #059669 100%);
        padding: 1.4rem 2rem; border-radius: 12px; margin-bottom: 1.5rem; color: white;
        border-bottom: 3px solid #34d399;
        box-shadow: 0 4px 16px rgba(6,95,70,0.18);
    }
    .main-header h1 { color: white; font-size: 1.8rem; margin: 0; font-weight: 800; letter-spacing: -0.01em; }
    .main-header p { color: #a7f3d0; font-size: 0.95rem; margin: 0.3rem 0 0 0; }

    .pipeline-step {
        background: #ecfdf5; border-left: 4px solid #059669;
        padding: 1rem 1.2rem; border-radius: 0 8px 8px 0; margin-bottom: 0.8rem;
    }
    .confidence-high { color: #059669; font-weight: 600; }
    .confidence-medium { color: #d68000; font-weight: 600; }
    .confidence-low { color: #EE2D25; font-weight: 600; }

    /* ═══ SIDEBAR ═══ */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #022c22 0%, #065f46 40%, #047857 100%) !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] li,
    [data-testid="stSidebar"] .stCaption p { color: #a7f3d0 !important; }
    [data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12) !important; }
    /* Sidebar selectbox / radio / widget labels */
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stRadio label {
        color: #d1fae5 !important;
    }
    [data-testid="stSidebar"] .stRadio [data-testid="stMarkdownContainer"] p {
        color: #d1fae5 !important;
    }
    /* Radio button active state */
    [data-testid="stSidebar"] .stRadio [role="radiogroup"] label[data-checked="true"] {
        background: rgba(52,211,153,0.15) !important;
        border-radius: 6px;
    }

    .sidebar-logo {
        text-align: center; padding: 1.2rem 0 0.6rem 0;
    }
    .sidebar-logo-icon {
        font-size: 2.4rem; display: block; margin-bottom: 0.3rem;
    }
    .sidebar-logo-text {
        font-size: 1.2rem; font-weight: 800; color: #ffffff;
        letter-spacing: 0.02em;
    }
    .sidebar-logo-sub {
        font-size: 0.72rem; color: #6ee7b7; letter-spacing: 0.06em;
        text-transform: uppercase; margin-top: 2px;
    }
    .sidebar-section {
        font-size: 0.68rem; font-weight: 700; color: #34d399;
        text-transform: uppercase; letter-spacing: 0.1em;
        padding: 0.8rem 0 0.4rem 0.5rem;
    }
    .sidebar-about-card {
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1);
        border-radius: 10px; padding: 0.8rem 1rem; margin-top: 0.3rem;
    }
    .sidebar-about-title {
        font-size: 0.7rem; font-weight: 600; color: #6ee7b7;
        text-transform: uppercase; letter-spacing: 0.08em; margin-bottom: 0.6rem;
    }
    .sidebar-about-item {
        display: flex; align-items: flex-start; gap: 0.5rem;
        padding: 0.35rem 0; font-size: 0.8rem; color: #a7f3d0;
    }
    .sidebar-about-item strong { color: #ffffff; }
    .sidebar-badge {
        display: inline-block; font-size: 0.62rem; font-weight: 700;
        padding: 2px 10px; border-radius: 20px;
        background: rgba(52,211,153,0.2); color: #34d399;
        letter-spacing: 0.03em;
    }
    .sidebar-footer {
        text-align: center; padding: 0.5rem 0; font-size: 0.72rem; color: #6ee7b7;
    }

    /* ═══ STREAMLIT WIDGET OVERRIDES (Emerald) ═══ */
    /* Primary buttons */
    .stButton > button[kind="primary"],
    .stButton > button[data-testid="baseButton-primary"] {
        background: linear-gradient(135deg, #065f46, #059669) !important;
        border: none !important; font-weight: 600;
        border-radius: 8px !important;
        justify-content: flex-start !important;
        box-shadow: 0 2px 8px rgba(6,95,70,0.18) !important;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="primary"] p,
    .stButton > button[data-testid="baseButton-primary"] p {
        color: white !important;
        text-align: left !important;
        width: 100% !important;
    }
    .stButton > button[kind="primary"]:hover,
    .stButton > button[data-testid="baseButton-primary"]:hover {
        background: linear-gradient(135deg, #022c22, #047857) !important;
        box-shadow: 0 4px 12px rgba(6,95,70,0.3) !important;
    }
    /* Secondary buttons */
    .stButton > button[kind="secondary"],
    .stButton > button[data-testid="baseButton-secondary"] {
        border: 1.5px solid #059669 !important;
        font-weight: 600; border-radius: 8px !important;
        justify-content: flex-start !important;
        transition: all 0.2s ease;
    }
    .stButton > button[kind="secondary"] p,
    .stButton > button[data-testid="baseButton-secondary"] p {
        color: #000000 !important;
        text-align: left !important;
        width: 100% !important;
    }
    .stButton > button[kind="secondary"]:hover,
    .stButton > button[data-testid="baseButton-secondary"]:hover {
        background: #d1fae5 !important; border-color: #065f46 !important;
    }
    /* Default buttons */
    .stButton > button {
        border: 1.5px solid #a7f3d0 !important;
        font-weight: 500; border-radius: 8px !important;
        justify-content: flex-start !important;
        text-align: left !important;
        transition: all 0.2s ease;
    }
    .stButton > button > div {
        display: flex !important;
        justify-content: flex-start !important;
        width: 100% !important;
        text-align: left !important;
    }
    .stButton > button p {
        color: #000000 !important;
        text-align: left !important;
        width: 100% !important;
        margin: 0 !important;
    }
    .stButton > button:hover {
        background: #d1fae5 !important; border-color: #059669 !important;
    }
    /* Download buttons */
    .stDownloadButton > button {
        border: 1.5px solid #059669 !important; color: #065f46 !important;
        font-weight: 600; border-radius: 8px !important;
    }
    .stDownloadButton > button:hover {
        background: #d1fae5 !important;
    }
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0 !important; font-weight: 600;
        color: #065f46 !important;
    }
    .stTabs [aria-selected="true"] {
        background: #d1fae5 !important; color: #065f46 !important;
        border-bottom: 3px solid #059669 !important;
    }
    /* Metrics */
    [data-testid="stMetricValue"] { color: #065f46 !important; font-weight: 800; }
    [data-testid="stMetricLabel"] { color: #047857 !important; }
    /* Expanders */
    details[data-testid="stExpander"] summary {
        color: #065f46 !important; font-weight: 600;
    }
    /* Selectbox / multiselect borders */
    .stSelectbox [data-baseweb="select"] > div,
    .stMultiSelect [data-baseweb="select"] > div {
        border-color: #a7f3d0 !important; border-radius: 8px !important;
    }
    .stSelectbox [data-baseweb="select"] > div:focus-within,
    .stMultiSelect [data-baseweb="select"] > div:focus-within {
        border-color: #059669 !important;
        box-shadow: 0 0 0 1px #059669 !important;
    }
    /* Progress bar */
    .stProgress > div > div > div { background: #059669 !important; }
    /* Dividers */
    hr { border-color: #a7f3d0 !important; }
    /* File uploader */
    [data-testid="stFileUploader"] section {
        border: 2px dashed #6ee7b7 !important; border-radius: 10px !important;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: #059669 !important; background: #ecfdf5 !important;
    }
    /* Toggle */
    .stToggle label span[data-checked="true"] {
        background-color: #059669 !important;
    }
    /* Text input */
    .stTextInput input { border-color: #a7f3d0 !important; border-radius: 8px !important; }
    .stTextInput input:focus { border-color: #059669 !important; box-shadow: 0 0 0 1px #059669 !important; }
    /* Status widget */
    [data-testid="stStatusWidget"] { border-left: 4px solid #059669 !important; }

    /* ═══ INFO / SUCCESS / WARNING / ERROR BANNERS ═══ */
    .stAlert [data-testid="stAlertContentInfo"] {
        background: #d1fae5 !important; border-left-color: #059669 !important;
    }

    /* ═══ GENERAL CARDS ═══ */
    .info-card {
        background: #fff; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 1rem; margin-bottom: 0.5rem; box-shadow: 0 1px 4px rgba(6,95,70,0.06);
    }
    .info-card .card-title { font-weight: 700; color: #065f46; margin-bottom: 0.3rem; }
    .info-card .card-subtitle { font-size: 0.8rem; color: #059669; }

    /* ═══ REPORT DOCUMENT CARDS ═══ */
    .doc-card {
        background: #fff; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 0; margin-bottom: 0.75rem; overflow: hidden;
        box-shadow: 0 1px 4px rgba(6,95,70,0.06); transition: all 0.2s;
    }
    .doc-card:hover { box-shadow: 0 4px 14px rgba(6,95,70,0.12); border-color: #6ee7b7; }
    .doc-card-header {
        display: flex; align-items: center; padding: 0.8rem 1.2rem;
        gap: 1rem; cursor: pointer;
    }
    .doc-card-num {
        font-weight: 700; font-size: 0.85rem; color: #059669;
        min-width: 32px; text-align: center;
    }
    .doc-card-type {
        font-size: 0.7rem; font-weight: 600; padding: 3px 10px;
        border-radius: 20px; text-transform: uppercase; letter-spacing: 0.03em;
    }
    .doc-card-company { font-weight: 600; color: #065f46; font-size: 0.95rem; flex: 1; }
    .doc-card-detail { color: #047857; font-size: 0.85rem; }
    .doc-card-amount { font-weight: 700; color: #065f46; font-size: 1rem; white-space: nowrap; }
    .doc-card-match {
        font-size: 0.72rem; padding: 2px 8px; border-radius: 10px;
        font-weight: 600; white-space: nowrap;
    }
    .doc-field-grid {
        display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
        gap: 0.6rem; padding: 0.8rem 1.2rem;
    }
    .doc-field { padding: 0.4rem 0; }
    .doc-field-label { font-size: 0.7rem; color: #34d399; text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 2px; }
    .doc-field-value { font-size: 0.88rem; color: #022c22; font-weight: 500; word-break: break-word; }
    .doc-status-verified { border-left: 4px solid #059669; }
    .doc-status-rejected { border-left: 4px solid #EE2D25; opacity: 0.6; }
    .doc-status-pending { border-left: 4px solid #a7f3d0; }
    .doc-card-status { font-size: 0.8rem; white-space: nowrap; color: #059669; }

    /* ═══ LINE-ITEM MATCHING PAGE — AP Reconciliation ═══ */
    .lim-header {
        background: linear-gradient(135deg, #022c22 0%, #065f46 50%, #059669 100%);
        padding: 1.6rem 2.2rem; border-radius: 12px; margin-bottom: 1.2rem;
        display: flex; align-items: center; gap: 1.2rem;
        border-bottom: 3px solid #34d399;
        box-shadow: 0 4px 16px rgba(6,95,70,0.18);
    }
    .lim-header-icon { font-size: 2.2rem; }
    .lim-header h2 { color: #fff; margin: 0; font-size: 1.5rem; font-weight: 800; letter-spacing: -0.01em; }
    .lim-header p { color: #a7f3d0; margin: 0.2rem 0 0 0; font-size: 0.85rem; }

    /* KPI cards */
    .lim-kpi-card {
        background: #fff; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 0.9rem 1rem; text-align: center;
        box-shadow: 0 1px 4px rgba(6,95,70,0.06); position: relative;
    }
    .lim-kpi-card .kpi-value { font-size: 1.45rem; font-weight: 800; margin-bottom: 0.1rem; line-height: 1.2; }
    .lim-kpi-card .kpi-label {
        font-size: 0.68rem; color: #059669; text-transform: uppercase;
        letter-spacing: 0.06em; font-weight: 600;
    }
    .lim-kpi-card .kpi-tooltip {
        font-size: 0.65rem; color: #6ee7b7; margin-top: 0.2rem;
        font-style: italic;
    }
    .lim-kpi-card .kpi-border-top { position: absolute; top: 0; left: 10%; right: 10%; height: 3px; border-radius: 0 0 3px 3px; }
    .lim-kpi-teal .kpi-value { color: #065f46; }
    .lim-kpi-teal .kpi-border-top { background: #065f46; }
    .lim-kpi-blue .kpi-value { color: #059669; }
    .lim-kpi-blue .kpi-border-top { background: #059669; }
    .lim-kpi-green .kpi-value { color: #047857; }
    .lim-kpi-green .kpi-border-top { background: #047857; }
    .lim-kpi-red .kpi-value { color: #EE2D25; }
    .lim-kpi-red .kpi-border-top { background: #EE2D25; }
    .lim-kpi-amber .kpi-value { color: #d68000; }
    .lim-kpi-amber .kpi-border-top { background: #d68000; }
    .lim-kpi-navy .kpi-value { color: #022c22; }
    .lim-kpi-navy .kpi-border-top { background: #022c22; }

    .lim-section-title {
        font-size: 0.95rem; font-weight: 700; color: #065f46;
        padding: 0.5rem 0 0.25rem 0; display: flex; align-items: center; gap: 0.5rem;
    }

    /* Tables */
    .lim-table {
        width: 100%; border-collapse: separate; border-spacing: 0;
        font-size: 0.84rem; border-radius: 10px; overflow: hidden;
        border: 1px solid #6ee7b7;
    }
    .lim-table thead th {
        background: #065f46; color: #fff; padding: 0.55rem 0.8rem;
        font-weight: 700; font-size: 0.75rem; text-transform: uppercase;
        letter-spacing: 0.04em; text-align: left;
    }
    .lim-table tbody td {
        padding: 0.5rem 0.8rem; border-bottom: 1px solid #d1fae5;
        color: #022c22; font-size: 0.84rem;
    }
    .lim-table tbody tr:last-child td { border-bottom: none; }
    .lim-table tbody tr:hover { background: #ecfdf5; }
    .lim-table .amount-col { text-align: right; font-weight: 600; font-variant-numeric: tabular-nums; }
    .lim-table .neg-amount { color: #EE2D25; }
    .lim-table-ledger thead th { background: #047857; }
    .lim-table-result thead th { background: #022c22; }
    .lim-table-result tbody td { font-size: 0.83rem; }

    /* Status badges */
    .lim-badge {
        display: inline-flex; align-items: center; gap: 0.3rem;
        padding: 2px 10px; border-radius: 20px; font-weight: 700;
        font-size: 0.73rem; white-space: nowrap;
    }
    .lim-badge-match { background: #d1fae5; color: #047857; }
    .lim-badge-mismatch { background: #fde8e7; color: #EE2D25; }
    .lim-badge-missing-ledger { background: #fff3e0; color: #d68000; }
    .lim-badge-missing-soa { background: #d1fae5; color: #059669; }
    .lim-badge-investigating { background: #d1fae5; color: #059669; }
    .lim-badge-resolved { background: #d1fae5; color: #065f46; }
    .lim-badge-approved { background: #d1fae5; color: #022c22; border: 1px solid #6ee7b7; }

    /* Summary / reconciliation bar */
    .lim-recon-bar {
        background: #fff; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 1rem 1.5rem; margin-top: 0.8rem;
        box-shadow: 0 1px 4px rgba(6,95,70,0.06);
    }
    .lim-recon-row {
        display: flex; justify-content: space-between; align-items: center;
        padding: 0.3rem 0; font-size: 0.88rem;
    }
    .lim-recon-row .recon-label { color: #047857; font-weight: 500; }
    .lim-recon-row .recon-value { font-weight: 700; color: #065f46; font-variant-numeric: tabular-nums; }
    .lim-recon-divider { border-bottom: 1px dashed #6ee7b7; margin: 0.3rem 0; }
    .lim-progress-track {
        background: #d1fae5; border-radius: 6px; height: 10px; width: 100%;
        overflow: hidden; margin-top: 0.4rem;
    }
    .lim-progress-fill { height: 100%; border-radius: 6px; transition: width 0.6s ease; }

    /* Variance rows */
    .lim-variance-row {
        background: #fff8e6; border-left: 4px solid #d68000;
        padding: 0.6rem 1rem; border-radius: 0 8px 8px 0;
        margin-bottom: 0.4rem; font-size: 0.84rem;
    }
    /* Filter bar */
    .lim-filter-bar {
        background: #ecfdf5; border: 1px solid #a7f3d0; border-radius: 12px;
        padding: 0.8rem 1.2rem; margin-bottom: 1rem;
    }
    /* Drill-down panel */
    .lim-drill-panel {
        background: #fff; border: 1px solid #6ee7b7; border-radius: 12px;
        padding: 1.2rem 1.5rem; margin-top: 0.5rem;
        box-shadow: 0 2px 8px rgba(6,95,70,0.08);
    }
    .lim-drill-field { margin-bottom: 0.5rem; }
    .lim-drill-field .dl-label {
        font-size: 0.68rem; color: #059669; text-transform: uppercase;
        letter-spacing: 0.05em; font-weight: 600; margin-bottom: 1px;
    }
    .lim-drill-field .dl-value { font-size: 0.92rem; color: #065f46; font-weight: 600; }
    .lim-drill-field .dl-value.dl-highlight { color: #EE2D25; background: #fde8e7; padding: 1px 6px; border-radius: 4px; }
    .lim-drill-field .dl-value.dl-ok { color: #047857; }

    .lim-comment-box {
        width: 100%; border: 1px solid #6ee7b7; border-radius: 8px;
        padding: 0.5rem 0.7rem; font-size: 0.84rem; resize: vertical;
        min-height: 60px; font-family: inherit;
    }
    .lim-comment-box:focus { border-color: #059669; outline: none; box-shadow: 0 0 0 2px rgba(5,150,105,0.15); }

    /* ═══ DATAFRAME THEMING ═══ */
    .stDataFrame { border-radius: 10px !important; overflow: hidden; }
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════
# DISPLAY HELPER FUNCTIONS (must be defined before page logic calls them)
# ═══════════════════════════════════════════════════════════════════════════

def display_confidence_bar(score: float) -> str:
    pct = int(score * 100)
    if score >= 0.95:
        return f"🟢 {pct}%"
    elif score >= 0.85:
        return f"🟡 {pct}%"
    else:
        return f"🔴 {pct}%"


def load_json_file(path: Path) -> dict | list | str:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as e:
        return f"Error: {e}"


def display_processing_file_preview(file_path: Path):
    """Preview a source document file from docs/uploads."""
    if not file_path.exists() or not file_path.is_file():
        st.error("Selected document file was not found.")
        return

    suffix = file_path.suffix.lower()
    st.caption(f"File: {file_path.name}")

    if suffix == ".pdf":
        pdf_bytes = file_path.read_bytes()
        # Render PDF pages as images (avoids iframe/CSP blocking on Streamlit Cloud)
        try:
            import tempfile, fitz  # noqa: E401
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            n_pages = len(doc)
            st.markdown(
                f'<div style="color:#5a8a8f;font-size:0.82rem;margin-bottom:0.4rem;">'
                f'📄 {n_pages} page{"s" if n_pages != 1 else ""}</div>',
                unsafe_allow_html=True,
            )
            for page_idx in range(n_pages):
                pix = doc[page_idx].get_pixmap(dpi=150)
                img_bytes = pix.tobytes("png")
                st.image(img_bytes, caption=f"Page {page_idx + 1} of {n_pages}", use_container_width=True)
            doc.close()
        except Exception as img_err:
            st.warning(f"Could not render PDF pages as images: {img_err}")
            st.info("Use the download button below to view the PDF.")
        st.download_button(
            "⬇️ Download PDF",
            data=pdf_bytes,
            file_name=file_path.name,
            mime="application/pdf",
        )
    elif suffix in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        st.image(str(file_path), use_container_width=True)
        st.download_button(
            "⬇️ Download Image",
            data=file_path.read_bytes(),
            file_name=file_path.name,
            mime="application/octet-stream",
        )
    else:
        st.info("Preview is not available for this file type. You can download the file below.")
        st.download_button(
            "⬇️ Download File",
            data=file_path.read_bytes(),
            file_name=file_path.name,
            mime="application/octet-stream",
        )


def display_ocr_result(data: dict):
    """Display OCR output with confidence scoring and section breakdown."""
    pages = []
    if "model_output" in data and isinstance(data["model_output"], dict):
        pages = data["model_output"].get("pages", [])
    elif "results" in data:
        for idx, result in enumerate(data["results"], start=1):
            if isinstance(result, dict):
                mo = result.get("model_output", {})
                if isinstance(mo, dict) and "pages" in mo:
                    for p in mo["pages"]:
                        p["page_number"] = idx  # normalise to sequential order
                        if not p.get("file_name"):
                            p["file_name"] = result.get("file", "")
                    pages.extend(mo["pages"])
                elif isinstance(mo, dict) and "pages" not in mo:
                    pass  # skip non-page outputs
    elif "pages" in data:
        pages = data["pages"]

    if not pages:
        st.warning("No structured OCR pages found. Showing raw JSON.")
        st.json(data)
        return

    # Summary metrics
    total_sections = sum(len(p.get("sections", [])) for p in pages)
    avg_confidence = 0.0
    conf_count = 0
    for p in pages:
        for s in p.get("sections", []):
            c = s.get("confidence", 0)
            if c > 0:
                avg_confidence += c
                conf_count += 1
    avg_confidence = avg_confidence / conf_count if conf_count > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("📄 Pages", len(pages))
    c2.metric("📝 Sections", total_sections)
    c3.metric("🎯 Avg Confidence", f"{avg_confidence:.0%}")
    st.divider()

    for page_data in pages:
        page_num = page_data.get("page_number", "?")
        file_name = page_data.get("file_name", "")
        sections = page_data.get("sections", [])

        with st.expander(f"📄 Page {page_num} — {file_name} ({len(sections)} sections)", expanded=(page_num == 1 or page_num == "1")):
            if not sections:
                st.caption("No sections.")
                continue

            type_counts: dict[str, int] = {}
            for s in sections:
                t = s.get("type", "unknown")
                type_counts[t] = type_counts.get(t, 0) + 1
            st.markdown("**Section types:** " + " ".join(f"`{t}` ×{c}" for t, c in sorted(type_counts.items(), key=lambda x: -x[1])))

            icon_map = {
                "header": "📌", "address": "📍", "key_value": "🔑",
                "table_header": "📊", "table_row": "📋", "subtotal": "💰",
                "paragraph": "📝", "footer": "📎", "signature": "✍️", "empty": "⬜",
            }

            for i, s in enumerate(sections):
                sec_type = s.get("type", "unknown")
                content = s.get("content", "")
                confidence = s.get("confidence", 0)
                icon = icon_map.get(sec_type, "📄")

                with st.container():
                    col_c, col_conf = st.columns([5, 1])
                    with col_c:
                        st.markdown(f"**{icon} {sec_type.upper()}**")
                        if sec_type in ("table_header", "table_row"):
                            st.code(content, language=None)
                        else:
                            st.text(content[:500])
                    with col_conf:
                        st.markdown(f"**{display_confidence_bar(confidence)}**")
                    if i < len(sections) - 1:
                        st.markdown("---")


def _display_ms_billing(data: dict):
    """Dedicated renderer for Microsoft / Cloud Billing Statement schema."""
    inv = data.get("invoice", {})
    credit_notes = data.get("credit_notes", []) or []
    currency_code = str(inv.get("currency") or "").strip().upper()

    st.markdown(f"**Document Type:** `{inv.get('document_type', 'SRKK - Microsoft Billing')}`")

    # ── Header ──────────────────────────────────────────────────────────
    st.markdown("#### 📋 Invoice Header")
    _header_fields = [
        ("vendor_name",    "🏢 Vendor"),
        ("billing_number", "🔢 Billing #"),
        ("document_date",  "📅 Document Date"),
        ("due_date",       "📅 Due Date"),
        ("payment_terms",  "📝 Payment Terms"),
        ("billing_profile","👤 Billing Profile"),
        ("currency",       "💱 Currency"),
        ("currency_note",  "ℹ️ Currency Note"),
    ]
    _hdr_vals = [(lbl, str(inv[k])) for k, lbl in _header_fields if inv.get(k)]

    # Billing period
    _bp = inv.get("billing_period") or {}
    if _bp.get("from") or _bp.get("to"):
        _hdr_vals.append(("📅 Billing Period", f"{_bp.get('from', '')} → {_bp.get('to', '')}"))

    # Vendor registration
    _vr = inv.get("vendor_registration") or {}
    if _vr.get("uen"):
        _hdr_vals.append(("🆔 UEN", str(_vr["uen"])))
    if _vr.get("service_tax_reg_no"):
        _hdr_vals.append(("🆔 Service Tax Reg", str(_vr["service_tax_reg_no"])))

    if _hdr_vals:
        cols = st.columns(min(len(_hdr_vals), 3))
        for i, (lbl, val) in enumerate(_hdr_vals):
            with cols[i % 3]:
                st.markdown(f"**{lbl}**")
                st.code(val[:150] + "..." if len(val) > 150 else val, language=None)

    st.divider()

    # ── Sold To / Bill To ────────────────────────────────────────────────
    _sold_to = inv.get("sold_to") or {}
    _bill_to = inv.get("bill_to") or {}
    if _sold_to or _bill_to:
        st.markdown("#### 📍 Parties")
        _party_cols = st.columns(2)
        with _party_cols[0]:
            if _sold_to:
                st.markdown("**Sold To**")
                st.info(
                    f"{_sold_to.get('name', '')}\n\n"
                    f"{_sold_to.get('address', '')}\n\n"
                    + (f"Reg: {_sold_to['registration_no']}" if _sold_to.get('registration_no') else "")
                )
        with _party_cols[1]:
            if _bill_to:
                st.markdown("**Bill To**")
                st.info(
                    f"{_bill_to.get('name', '')}\n\n"
                    f"{_bill_to.get('address', '')}"
                )
        st.divider()

    # ── Billing Summary ──────────────────────────────────────────────────
    _bs = inv.get("billing_summary") or {}
    if any(_bs.values()):
        st.markdown("#### 💰 Billing Summary")
        _bs_cols = st.columns(5)
        for i, (k, lbl) in enumerate([
            ("charges", "Charges"), ("credits", "Credits"),
            ("subtotal", "Subtotal"), ("tax", "Tax"), ("total", "Total"),
        ]):
            val = _bs.get(k)
            if val:
                _bs_cols[i].metric(lbl, f"{currency_code} {val}" if currency_code else str(val))
        st.divider()

    # ── Tax Invoice Line Items ───────────────────────────────────────────
    _ti = inv.get("tax_invoice") or {}
    _li = _ti.get("line_items") or []
    if _li:
        st.markdown(f"#### 📦 Tax Invoice Line Items ({len(_li)} rows)")
        if _ti.get("tax_invoice_number"):
            st.caption(f"Tax Invoice #: {_ti['tax_invoice_number']}  |  Date: {_ti.get('tax_invoice_date', '')}")
        _li_df = pd.DataFrame(_li)
        _li_df.columns = [c.replace("_", " ").title() for c in _li_df.columns]
        _li_df = _li_df.dropna(axis=1, how="all")
        st.dataframe(_li_df, use_container_width=True, hide_index=True, height=min(600, 40 + len(_li_df) * 35))
        if _ti.get("total_including_vat"):
            st.markdown(f"**Total (incl. VAT):** `{currency_code} {_ti['total_including_vat']}`")
        st.divider()

    # ── Payment Instructions ─────────────────────────────────────────────
    _pi = inv.get("payment_instructions") or {}
    if any(v for v in _pi.values() if v):
        st.markdown("#### 💳 Payment Instructions")
        _pi_cols = st.columns(3)
        for i, (k, lbl) in enumerate([
            ("method", "Method"), ("bank", "Bank"), ("branch", "Branch"),
            ("swift_code", "SWIFT"), ("account_number", "Account #"), ("account_name", "Account Name"),
        ]):
            if _pi.get(k):
                with _pi_cols[i % 3]:
                    st.markdown(f"**{lbl}**")
                    st.code(str(_pi[k]), language=None)
        st.divider()

    # ── Credit Notes ─────────────────────────────────────────────────────
    if credit_notes:
        st.markdown(f"#### 📌 Credit Notes ({len(credit_notes)})")
        for cn in credit_notes:
            _cn_title = f"CN #{cn.get('credit_note_number', '—')}  |  {cn.get('credit_note_date', '')}  |  Orig: {cn.get('original_billing_number', '')}"
            with st.expander(_cn_title, expanded=False):
                _cn_bp = cn.get("billing_period") or {}
                if _cn_bp.get("from") or _cn_bp.get("to"):
                    st.caption(f"Period: {_cn_bp.get('from', '')} → {_cn_bp.get('to', '')}")
                if cn.get("reason"):
                    st.caption(f"Reason: {cn['reason']}")
                _cn_li = cn.get("line_items") or []
                if _cn_li:
                    _cn_df = pd.DataFrame(_cn_li)
                    _cn_df.columns = [c.replace("_", " ").title() for c in _cn_df.columns]
                    st.dataframe(_cn_df, use_container_width=True, hide_index=True)
                if cn.get("total_including_vat"):
                    st.markdown(f"**Total (incl. VAT):** `{currency_code} {cn['total_including_vat']}`")
                if cn.get("credit_applicable_note"):
                    st.info(cn["credit_applicable_note"])
        st.divider()

    # ── Publisher Information ─────────────────────────────────────────────
    _pub = inv.get("publisher_information") or []
    if _pub:
        st.markdown(f"#### 🏭 Publisher Information ({len(_pub)} publishers)")
        with st.expander("View publishers", expanded=False):
            _pub_df = pd.DataFrame(_pub)
            _pub_df.columns = [c.replace("_", " ").title() for c in _pub_df.columns]
            st.dataframe(_pub_df, use_container_width=True, hide_index=True)


def display_extraction_result(data: dict, doc_type: str = "Unknown"):
    """Display extraction result in structured, professional format."""
    if not isinstance(data, dict):
        st.json(data)
        return

    # Delegate to dedicated renderer for Microsoft / Cloud Billing Statement
    if isinstance(data.get("invoice"), dict) and "billing_summary" in data.get("invoice", {}):
        _display_ms_billing(data)
        return

    actual_type = data.get("document_type", doc_type)
    st.markdown(f"**Document Type:** `{actual_type}`")
    currency_code = str(data.get("currency") or "").strip().upper()

    # Core fields
    st.markdown("#### 📋 Document Summary")
    core_fields = [
        ("vendor_name", "🏢 Vendor"), ("document_number", "🔢 Document #"),
        ("invoice_number", "🔢 Invoice #"), ("po_number", "🔢 PO #"),
        ("document_date", "📅 Date"), ("invoice_date", "📅 Invoice Date"),
        ("po_date", "📅 PO Date"), ("due_date", "📅 Due Date"),
        ("currency", "💱 Currency"), ("total_amount", "💰 Total Amount"),
        ("grand_total", "💰 Grand Total"), ("total_incl_tax", "💰 Total (incl. Tax)"),
        ("bill_to", "📍 Bill To"), ("account_number", "🔑 Account #"),
        ("customer_name", "👤 Customer"), ("account_holder", "👤 Account Holder"),
        ("bank_name", "🏦 Bank"), ("statement_date", "📅 Statement Date"),
        ("billing_period_from", "📅 Period From"), ("billing_period_to", "📅 Period To"),
        ("statement_period_from", "📅 Period From"), ("statement_period_to", "📅 Period To"),
    ]
    core_money_keys = {"total_amount", "grand_total", "total_incl_tax"}
    displayed = []
    for key, label in core_fields:
        if not data.get(key) or str(data.get(key)) == "null":
            continue
        value = data[key]
        if key in core_money_keys:
            value = _format_money_with_currency(value, currency_code)
        displayed.append((label, str(value)))

    if displayed:
        cols = st.columns(min(len(displayed), 3))
        for i, (label, val) in enumerate(displayed):
            with cols[i % 3]:
                display_val = val[:120] + "..." if len(val) > 120 else val
                st.markdown(f"**{label}**")
                st.code(display_val, language=None)

    st.divider()

    # PO-specific: Supplier / Buyer / Delivery Recipient parties
    _po_supplier = data.get("supplier") or {}
    _po_buyer = data.get("buyer") or {}
    _po_delivery = data.get("delivery_recipient") or {}
    if any([_po_supplier, _po_buyer, _po_delivery]):
        st.markdown("#### 📍 Parties")
        _party_cols = st.columns(3)
        with _party_cols[0]:
            if _po_buyer.get("name"):
                st.markdown("**Buyer (Issuer)**")
                st.info(
                    f"{_po_buyer.get('name', '')}\n\n"
                    f"{_po_buyer.get('address', '')}"
                )
        with _party_cols[1]:
            if _po_supplier.get("name"):
                st.markdown("**Supplier (Order From)**")
                st.info(
                    f"{_po_supplier.get('name', '')}\n\n"
                    f"{_po_supplier.get('address', '')}"
                )
        with _party_cols[2]:
            if _po_delivery.get("name"):
                st.markdown("**Deliver To**")
                st.info(
                    f"{_po_delivery.get('name', '')}\n\n"
                    f"{_po_delivery.get('address', '')}"
                )
        st.divider()

    # Line items / Transactions
    line_items = data.get("line_items", [])
    transactions = data.get("transactions", [])
    items = line_items or transactions

    if items:
        label = "Transactions" if transactions else "Line Items"
        st.markdown(f"#### 📦 {label} ({len(items)} rows)")
        df = pd.DataFrame(items)
        df.columns = [c.replace("_", " ").title() for c in df.columns]
        for money_col in ("Unit Price", "Tax", "Amount"):
            if money_col in df.columns:
                df[money_col] = df[money_col].apply(lambda v: _format_money_with_currency(v, currency_code))
        df = df.dropna(axis=1, how="all")
        if "Low Confidence" in df.columns and not df["Low Confidence"].any():
            df = df.drop(columns=["Low Confidence"])
        st.dataframe(df, use_container_width=True, hide_index=True, height=min(500, 40 + len(df) * 35))

    # Totals
    total_fields = [
        ("subtotal", "Subtotal"), ("tax_total", "Tax Total"), ("discount", "Discount"),
        ("freight_charges", "Freight"), ("grand_total", "Grand Total"),
        ("amount_in_words", "In Words"), ("opening_balance", "Opening Bal"),
        ("closing_balance", "Closing Bal"), ("total_debits", "Total Debits"),
        ("total_credits", "Total Credits"), ("previous_balance", "Prev Balance"),
        ("current_charges", "Current Charges"), ("payment_received", "Payment Recv"),
    ]
    total_money_keys = {
        "subtotal", "tax_total", "discount", "freight_charges", "grand_total",
        "opening_balance", "closing_balance", "total_debits", "total_credits",
        "previous_balance", "current_charges", "payment_received",
    }
    totals = []
    for key, label in total_fields:
        if not data.get(key):
            continue
        value = data[key]
        if key in total_money_keys:
            value = _format_money_with_currency(value, currency_code)
        totals.append((label, value))
    if totals:
        st.markdown("#### 💰 Totals & Summary")
        cols = st.columns(min(len(totals), 4))
        for i, (lbl, val) in enumerate(totals):
            with cols[i % 4]:
                st.metric(lbl, str(val))

    # Surcharges
    surcharges = data.get("surcharges", [])
    if surcharges:
        st.markdown("#### 📎 Surcharges & Levies")
        st.dataframe(pd.DataFrame(surcharges), use_container_width=True, hide_index=True)

    # Additional fields
    addl = data.get("additional_fields", {})
    if addl:
        st.markdown("#### 📎 Additional Fields")
        st.dataframe(pd.DataFrame([{"Field": k, "Value": str(v)} for k, v in addl.items()]), use_container_width=True, hide_index=True)

    # Payment info
    payment = data.get("payment_info", {})
    if payment and any(v for v in payment.values() if v):
        st.markdown("#### 💳 Payment Information")
        for k, v in payment.items():
            if v:
                st.markdown(f"**{k.replace('_', ' ').title()}:** `{v}`")


def display_bank_matching(data: dict, report_path: Path = None):
    """Display bank matching results with visual indicators."""
    bank = data.get("bank_statement_summary", {})

    st.markdown("#### 🏦 Bank Statement Summary")
    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Bank", bank.get("bank", "N/A"))
    b2.metric("Account #", str(bank.get("account_no", "N/A")))
    b3.metric("Total Credits", f"MYR {bank.get('total_credits', 0):,.2f}")
    b4.metric("Total Debits", f"MYR {bank.get('total_debits', 0):,.2f}")
    st.markdown(f"**Period:** {bank.get('period', 'N/A')} | **Entries:** {bank.get('total_entries', 0)}")
    st.divider()

    # Documents summary
    docs = data.get("documents_summary", [])
    if docs:
        st.markdown(f"#### 📄 Extracted Documents ({len(docs)} files)")
        df_docs = pd.DataFrame(docs)
        df_docs.columns = [c.replace("_", " ").title() for c in df_docs.columns]
        st.dataframe(df_docs, use_container_width=True, hide_index=True)
        st.divider()

    # Match results
    exact = data.get("exact_matches", [])
    near = data.get("near_matches", [])
    unmatched_bank = data.get("unmatched_bank_entries", [])
    unmatched_docs = data.get("unmatched_documents", [])

    st.markdown("#### 🎯 Matching Results")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("✅ Exact Matches", len(exact))
    m2.metric("🟡 Near Matches", len(near))
    m3.metric("❌ Unmatched Bank", len(unmatched_bank))
    m4.metric("❌ Unmatched Docs", len(unmatched_docs))

    if exact:
        st.markdown("##### ✅ Exact Matches")
        for m in exact:
            st.success(
                f"**Bank:** {m.get('bank_date','')} | {m.get('bank_description','')} | "
                f"{m.get('bank_type','')} MYR {m.get('bank_amount','')}\n\n"
                f"**Doc:** {m.get('doc_file','')} | #{m.get('doc_number','')} | "
                f"{m.get('doc_vendor','')} | MYR {m.get('doc_amount','')} ({m.get('match_field','')})"
            )

    if near:
        st.markdown("##### 🟡 Near Matches")
        for m in near:
            st.warning(
                f"**Bank:** {m.get('bank_date','')} | {m.get('bank_description','')} | "
                f"MYR {m.get('bank_amount','')}\n\n"
                f"**Doc:** {m.get('doc_file','')} | MYR {m.get('doc_amount','')} | "
                f"Diff: {m.get('difference_pct','?')}%"
            )

    if unmatched_bank:
        st.markdown("##### ❌ Unmatched Bank Entries")
        df_ub = pd.DataFrame(unmatched_bank)
        if not df_ub.empty:
            df_ub.columns = [c.replace("_", " ").title() for c in df_ub.columns]
            st.dataframe(df_ub, use_container_width=True, hide_index=True)

    # Full text report
    if report_path and report_path.exists():
        st.divider()
        with st.expander("📄 View Full Matching Report", expanded=False):
            st.code(report_path.read_text(encoding="utf-8"), language=None)
        st.download_button(
            "⬇️ Download Full Report",
            data=report_path.read_text(encoding="utf-8"),
            file_name="bank_matching_report.txt", mime="text/plain",
        )


# ═══════════════════════════════════════════════════════════════════════════
# REPORT FORMAT HELPERS
# ═══════════════════════════════════════════════════════════════════════════

REPORT_COLUMNS = [
    "No", "Upload Date", "Company Name", "TIN No", "Types (Inv/CN)", "Invoice No",
    "Tax Invoice No", "Invoice Date",
    "Lot No", "Location", "Account No", "Lease ID", "Unit No", "Project",
    "Premise Address", "Contract No / Batch No", "Contract Account No",
    "Description", "Total Amount (incl. Tax)", "Electricity Amount",
    "LHDN UUID", "Validate On", "Kwh Reading Before", "Kwh Reading After",
    "Current Reading / Total Units",
]


def _safe(d: dict, *keys, default=""):
    """Dig through nested dicts safely."""
    cur = d
    for k in keys:
        if isinstance(cur, dict):
            cur = cur.get(k, default)
        else:
            return default
    return cur if cur not in (None, "") else default


def _format_money_with_currency(value: object, currency: str) -> str:
    if value in (None, ""):
        return ""

    text = str(value).strip()
    if not text:
        return ""

    currency_code = (currency or "").strip().upper()
    if not currency_code:
        return text

    if text.upper().startswith(f"{currency_code} ") or text.upper() == currency_code:
        return text

    return f"{currency_code} {text}"


def _parse_lot_no(data: dict) -> str:
    """Extract Lot No from bill_to or additional_fields."""
    bill_to = _safe(data, "bill_to")
    if isinstance(bill_to, str):
        import re
        m = re.search(r"Lot\s*No\.?\s*[:：]?\s*(\S+)", bill_to, re.IGNORECASE)
        if m:
            return m.group(1)
    return _safe(data, "additional_fields", "Lot No")


def _parse_unit_no(data: dict) -> str:
    """Extract Unit No from line items or additional_fields."""
    af = _safe(data, "additional_fields", "Unit No")
    if af:
        return af
    for item in data.get("line_items", []):
        desc = item.get("description", "") or item.get("product_description", "") or ""
        import re
        m = re.search(r"Unit\s*No\.?\s*[:：]?\s*(\S+)", desc, re.IGNORECASE)
        if m:
            return m.group(1)
    return ""


def _parse_location(data: dict) -> str:
    """Extract location from line items description or bill_to."""
    for item in data.get("line_items", []):
        desc = item.get("description", "") or ""
        if "Mall" in desc or "Hotel" in desc or "Plaza" in desc:
            import re
            m = re.search(r"(?:The\s+\w+\s+Mall|[\w\s]+Mall|[\w\s]+Hotel|[\w\s]+Plaza)", desc)
            if m:
                return m.group(0).strip()
    return ""


def _parse_kwh_readings(data: dict) -> tuple:
    """Parse kWh meter readings from utility bill line items."""
    readings_before, readings_after, total_units = [], [], []
    for item in data.get("line_items", []):
        desc = item.get("description", "") or ""
        import re
        m = re.search(r"Meter Readings?\s*([\d,]+(?:\.\d+)?)\s*[-–]\s*([\d,]+(?:\.\d+)?)", desc)
        if m:
            before = m.group(1).replace(",", "")
            after = m.group(2).replace(",", "")
            if before not in readings_before:
                readings_before.append(before)
            if after not in readings_after:
                readings_after.append(after)
        qty = item.get("quantity")
        if qty and desc and "Meter" in desc:
            total_units.append(str(qty))
    return (
        " / ".join(readings_before) if readings_before else "",
        " / ".join(readings_after) if readings_after else "",
        " / ".join(total_units) if total_units else "",
    )


def _electricity_amount(data: dict) -> str:
    """Sum electricity charge amounts for utility bills."""
    for item in data.get("line_items", []):
        desc = (item.get("description", "") or "").lower()
        if "electricity" in desc and item.get("amount"):
            return str(item["amount"])
    return ""


def _build_description(data: dict) -> str:
    """Build a short description from line items."""
    descs = []
    for item in data.get("line_items", []):
        d = item.get("description") or item.get("product_description") or ""
        first_line = d.split("\n")[0].strip()
        if first_line and first_line not in descs:
            descs.append(first_line)
    return "; ".join(descs[:4]) + ("..." if len(descs) > 4 else "")


def _doc_type_label(data: dict) -> str:
    """Return the document type label (Inv / CN / Utility / etc.)."""
    dt = (_safe(data, "document_type") or "").lower()
    if "credit" in dt:
        return "CN"
    if "utility" in dt:
        return "Utility"
    if "rental" in dt or "lease" in dt:
        return "Rental"
    if "statement" in dt:
        return "SOA"
    if "hotel" in dt:
        return "Hotel"
    if "travel" in dt:
        return "Travel"
    if ("srkk" in dt and "vendor" in dt) or dt == "vendor invoice":
        return "SRKK-Vendor"
    if ("srkk" in dt and "purchase" in dt) or dt == "purchase order":
        return "SRKK-PO"
    if ("srkk" in dt and "microsoft" in dt) or "srkk_microsoft_billing" in dt:
        return "SRKK-MS Billing"
    return "Inv"


def map_extraction_to_report_row(data: dict, index: int) -> dict:
    """Map a single extraction JSON to a report format row."""
    af = data.get("additional_fields", {}) or {}
    currency_code = str(_safe(data, "currency") or "").strip().upper()
    kwh_before, kwh_after, total_units = _parse_kwh_readings(data)
    dt_lower = (_safe(data, "document_type") or "").lower()

    # ── SRKK-specific field resolution ───────────────────────────────────
    if "microsoft" in dt_lower or "billing" in dt_lower:
        # MS Billing data is nested under data["invoice"]
        inv_data     = data.get("invoice", data)
        invoice_no   = _safe(inv_data, "billing_number") or _safe(inv_data, "invoice_number") or ""
        invoice_date = _safe(inv_data, "document_date") or _safe(inv_data, "billing_date") or _safe(inv_data, "invoice_date") or ""
        company_name = _safe(inv_data, "vendor_name") or _safe(inv_data, "bill_to", "name") or ""
        total_amt    = _safe(inv_data, "grand_total") or _safe(inv_data, "total_amount") or ""
    elif "purchase order" in dt_lower or "srkk" in dt_lower and "purchase" in dt_lower:
        invoice_no   = _safe(data, "po_number") or ""
        invoice_date = _safe(data, "po_date") or ""
        company_name = _safe(data, "supplier", "name") or _safe(data, "buyer", "name") or ""
        total_amt    = _safe(data, "total_incl_tax") or _safe(data, "total_amount") or _safe(data, "grand_total") or ""
    elif "vendor" in dt_lower:
        invoice_no   = _safe(data, "invoice_number") or _safe(data, "document_number") or ""
        invoice_date = _safe(data, "invoice_date") or _safe(data, "document_date") or ""
        company_name = _safe(data, "vendor_name") or ""
        total_amt    = _safe(data, "total_amount_payable") or _safe(data, "grand_total") or _safe(data, "total_amount") or ""
    else:
        invoice_no = (
            _safe(data, "invoice_number")
            or _safe(data, "document_number")
            or _safe(data, "statement_number")
            or ""
        )
        invoice_date = (
            _safe(data, "invoice_date")
            or _safe(data, "document_date")
            or _safe(data, "statement_date")
            or _safe(af, "Invoice Date")
            or ""
        )
        company_name = _safe(data, "vendor_name")
        total_amt    = _safe(data, "grand_total") or _safe(data, "total_amount")

    account_no = (
        _safe(data, "account_number")
        or _safe(data, "customer_account")
        or _safe(af, "Account No")
        or _safe(data, "payment_info", "account_number")
        or ""
    )

    uploaded_at_raw = data.get("_uploaded_at", "")
    try:
        uploaded_at = datetime.strptime(uploaded_at_raw, "%Y-%m-%d %H:%M:%S").date() if uploaded_at_raw else None
    except ValueError:
        uploaded_at = None

    return {
        "No": index,
        "Upload Date": uploaded_at,
        "Company Name": company_name,
        "TIN No": _safe(af, "TIN No.") or _safe(af, "TIN No"),
        "Types (Inv/CN)": _doc_type_label(data),
        "Invoice No": invoice_no,
        "Tax Invoice No": _safe(af, "No. Invois Cukai") or _safe(af, "Tax Invoice No"),
        "Invoice Date": invoice_date,
        "Lot No": _parse_lot_no(data),
        "Location": _parse_location(data),
        "Account No": account_no,
        "Lease ID": _safe(af, "Lease ID"),
        "Unit No": _parse_unit_no(data),
        "Project": _safe(af, "Project") or _safe(af, "Project Name"),
        "Premise Address": _safe(data, "service_address") or _safe(data, "bill_to"),
        "Contract No / Batch No": _safe(af, "Contract No") or _safe(af, "Batch No") or _safe(af, "Contract No / Batch No"),
        "Contract Account No": _safe(af, "Contract Account No"),
        "Description": _build_description(data),
        "Total Amount (incl. Tax)": _format_money_with_currency(total_amt, currency_code),
        "Electricity Amount": _format_money_with_currency(_electricity_amount(data), currency_code),
        "LHDN UUID": _safe(af, "LHDN UUID") or _safe(af, "e-Invoice UUID"),
        "Validate On": _safe(af, "Validate On") or _safe(af, "Validated On"),
        "Kwh Reading Before": kwh_before,
        "Kwh Reading After": kwh_after,
        "Current Reading / Total Units": total_units,
    }


def load_all_extraction_rows() -> pd.DataFrame:
    """Load all extraction JSON files and map to report DataFrame."""
    extraction_dir = Path(__file__).resolve().parent / "output" / "extraction"
    rows = []
    if extraction_dir.exists():
        files = sorted(
            [
                f for f in extraction_dir.glob("*.json")
                if f.name not in {"bank_matching_results.json"}
            ]
        )
        for i, f in enumerate(files, start=1):
            try:
                data = json.loads(f.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    # Skip bank statement extractions – they lack invoice-level fields
                    dt_lower = (data.get("document_type") or "").lower()
                    if "bank statement" in dt_lower:
                        continue
                    row = map_extraction_to_report_row(data, i)
                    row["_source_file"] = f.name
                    row["_uploaded_at_raw"] = data.get("_uploaded_at", "")
                    rows.append(row)
            except Exception:
                pass
    # Re-number rows sequentially after skipping
    for idx, row in enumerate(rows, start=1):
        row["No"] = idx
    if not rows:
        return pd.DataFrame(columns=REPORT_COLUMNS)
    df = pd.DataFrame(rows)
    # Ensure all report columns exist
    for col in REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df


def _team_from_doc_type(doc_type: str) -> str:
    text = (doc_type or "").strip().lower()
    if "rental" in text or "lease" in text or "utility" in text:
        return "rental"
    return "sales"


def _doc_team_map_path() -> Path:
    return Path(__file__).resolve().parent / "docs" / "database" / "doc_teams.json"


def load_doc_team_map() -> dict[str, str]:
    path = _doc_team_map_path()
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(data, dict):
            return {str(k): str(v).lower() for k, v in data.items()}
    except Exception:
        pass
    return {}


def save_doc_team_map(doc_team_map: dict[str, str]) -> None:
    path = _doc_team_map_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(doc_team_map, ensure_ascii=False, indent=2), encoding="utf-8")


def infer_document_team(file_path: Path, doc_team_map: dict[str, str]) -> str:
    mapped = doc_team_map.get(file_path.name)
    if mapped in {"sales", "rental"}:
        return mapped

    extraction_dir = Path(__file__).resolve().parent / "output" / "extraction"
    if extraction_dir.exists():
        candidates = sorted(
            [
                p for p in extraction_dir.glob("*.json")
                if p.name not in {"bank_matching_results.json"}
                and (p.stem == file_path.stem or p.stem.startswith(file_path.stem + "_extracted"))
            ],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        for candidate in candidates:
            try:
                data = json.loads(candidate.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    return _team_from_doc_type(str(data.get("document_type") or ""))
            except Exception:
                continue

    name_l = file_path.name.lower()
    if any(token in name_l for token in ["rental", "lease", "ll_"]):
        return "rental"
    return "sales"


def find_source_pdf_for_extraction(source_file: str) -> Path | None:
    """Match extraction JSON source filename to original PDF in docs/uploads."""
    if not source_file:
        return None

    database_dir = Path(__file__).resolve().parent / "docs" / "uploads"
    if not database_dir.exists():
        return None

    source_stem = Path(source_file).stem
    base = source_stem.replace("_extracted", "")

    exact_name = f"{base}.pdf"
    exact_path = database_dir / exact_name
    if exact_path.exists():
        return exact_path

    pdf_files = [p for p in database_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    for pdf in pdf_files:
        if pdf.name.lower() == exact_name.lower():
            return pdf

    for pdf in pdf_files:
        if pdf.stem.lower().startswith(base.lower()):
            return pdf

    return None


def load_extraction_repository_items() -> list[dict]:
    """Build row data for Extraction Viewer repository layout."""
    extraction_dir = Path(__file__).resolve().parent / "output" / "extraction"
    if not extraction_dir.exists():
        return []

    items: list[dict] = []
    files = sorted(
        [
            f for f in extraction_dir.glob("*.json")
            if f.name not in {"bank_matching_results.json"}
        ],
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    for f in files:
        try:
            data = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if not isinstance(data, dict):
            continue

        additional = data.get("additional_fields", {}) or {}

        # Handle nested Microsoft/Cloud Billing schema (top-level keys: "invoice", "credit_notes")
        _inv_node = data.get("invoice", {}) if isinstance(data.get("invoice"), dict) else {}
        _is_ms_billing = bool(_inv_node)

        invoice_id = (
            data.get("invoice_number")
            or data.get("document_number")
            or data.get("statement_number")
            or (_inv_node.get("billing_number") if _is_ms_billing else None)
            or f.stem.replace("_extracted", "")
        )
        _supplier_node = data.get("supplier") or {}
        _buyer_node = data.get("buyer") or {}
        vendor = (
            data.get("vendor_name")
            or data.get("customer_name")
            or data.get("account_holder")
            or (_inv_node.get("vendor_name") if _is_ms_billing else None)
            or (_supplier_node.get("name") if _supplier_node else None)
            or (_buyer_node.get("name") if _buyer_node else None)
            or "-"
        )
        date_text = (
            data.get("invoice_date")
            or data.get("document_date")
            or data.get("statement_date")
            or data.get("po_date")
            or (_inv_node.get("document_date") if _is_ms_billing else None)
            or "-"
        )
        currency_code = str(
            data.get("currency")
            or (_inv_node.get("currency") if _is_ms_billing else None)
            or ""
        ).strip().upper()
        _ms_total = (
            (_inv_node.get("billing_summary") or {}).get("total")
            if _is_ms_billing else None
        )
        total_raw = data.get("grand_total") or data.get("total_amount") or data.get("subtotal") or data.get("total_incl_tax") or _ms_total or ""
        total_text = _format_money_with_currency(total_raw, currency_code) if total_raw else "-"
        team = _team_from_doc_type(str(data.get("document_type") or ""))

        status_text = st.session_state["processing_doc_status"].get(f.name, "Ready for Review")

        items.append(
            {
                "invoice_id": str(invoice_id),
                "vendor": str(vendor),
                "date": str(date_text),
                "total": str(total_text),
                "status": str(status_text),
                "team": team,
                "last_updated": datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d"),
                "source_file": f.name,
                "data": data,
            }
        )

    return items


def _normalize_lease_id(lid: str) -> str:
    """Strip non-digit prefixes from Lease ID for fuzzy comparison."""
    import re
    if not lid:
        return ""
    return re.sub(r"[^\d]", "", str(lid))


def match_utility_to_rental(df: pd.DataFrame) -> list[dict]:
    """
    Match utility bills to rental invoices using shared identifiers.
    Returns a list of match dicts:  {utility_idx, rental_idx, matched_on, keys}
    Priority: Lease ID > Lot No > (Vendor + TIN No)
    """
    utilities = df[df["Types (Inv/CN)"] == "Utility"]
    rentals = df[df["Types (Inv/CN)"].isin(["Rental", "Inv"])]
    matches = []
    matched_rental_ids = set()

    for u_idx, u_row in utilities.iterrows():
        best_match = None
        best_score = 0
        best_keys = []

        u_lease = _normalize_lease_id(u_row.get("Lease ID", ""))
        u_lot = str(u_row.get("Lot No", "")).strip()
        u_vendor = str(u_row.get("Company Name", "")).strip().lower()
        u_tin = str(u_row.get("TIN No", "")).strip()
        u_acct = str(u_row.get("Account No", "")).strip()

        for r_idx, r_row in rentals.iterrows():
            if r_idx in matched_rental_ids:
                continue
            score = 0
            keys = []

            r_lease = _normalize_lease_id(r_row.get("Lease ID", ""))
            if u_lease and r_lease and u_lease == r_lease:
                score += 4
                keys.append(f"Lease ID ({u_row.get('Lease ID', '')} ↔ {r_row.get('Lease ID', '')})")

            r_lot = str(r_row.get("Lot No", "")).strip()
            if u_lot and r_lot and u_lot == r_lot:
                score += 3
                keys.append(f"Lot No ({u_lot})")

            r_vendor = str(r_row.get("Company Name", "")).strip().lower()
            if u_vendor and r_vendor and u_vendor == r_vendor:
                score += 2
                keys.append("Vendor Name")

            r_tin = str(r_row.get("TIN No", "")).strip()
            if u_tin and r_tin and u_tin == r_tin:
                score += 2
                keys.append(f"TIN No ({u_tin})")

            r_acct = str(r_row.get("Account No", "")).strip()
            if u_acct and r_acct and u_acct == r_acct:
                score += 1
                keys.append(f"Account No ({u_acct})")

            if score > best_score:
                best_score = score
                best_match = r_idx
                best_keys = keys

        if best_match is not None and best_score >= 2:
            matched_rental_ids.add(best_match)
            matches.append({
                "utility_idx": u_idx,
                "rental_idx": best_match,
                "score": best_score,
                "matched_on": best_keys,
            })

    return matches


# ═══════════════════════════════════════════════════════════════════════════
# HEADER
# ═══════════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="main-header">
    <h1>📄 SRKK Document Intelligence</h1>
    <p>AI-Powered OCR &bull; Document Classification &bull; Data Extraction &bull; Statement Matching</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    # ── Logo / Brand ──────────────────────────────────────────
    st.markdown(
        '<div class="sidebar-logo">'
        '<span class="sidebar-logo-icon">📄</span>'
        '<div class="sidebar-logo-text">SRKK Document Intelligence</div>'
        '<div class="sidebar-logo-sub">Document AI Platform</div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown("")

    # ── Dashboard ──────────────────────────────────────────────
    _dash_label = "🏠 Dashboard"
    _is_dash    = st.session_state.get("page") == _dash_label
    if st.button(_dash_label, key="nav_btn_dashboard", use_container_width=True,
                 type="primary" if _is_dash else "secondary"):
        st.session_state["page"] = _dash_label
        st.rerun()

    st.markdown("")

    # ── OCR Processing ────────────────────────────────────────────
    st.markdown('<div class="sidebar-section">OCR Processing</div>', unsafe_allow_html=True)
    _nav_pages = [
        "📤 Document Processing",
        "🔍 OCR Viewer",
        "📊 Extraction Viewer",
        "📋 Report Format",
    ]
    for _nav_label in _nav_pages:
        _is_active = st.session_state.get("page") == _nav_label
        if st.button(
            _nav_label,
            key=f"nav_btn_{_nav_label}",
            use_container_width=True,
            type="primary" if _is_active else "secondary",
        ):
            st.session_state["page"] = _nav_label
            st.rerun()

    # ── Reconciliation ────────────────────────────────────────────
    st.markdown('<div class="sidebar-section">Reconciliation</div>', unsafe_allow_html=True)
    for _recon_label in ["🔄 Reconciliation"]:
        _is_active = st.session_state.get("page") == _recon_label
        if st.button(
            _recon_label,
            key=f"recon_btn_{_recon_label}",
            use_container_width=True,
            type="primary" if _is_active else "secondary",
        ):
            st.session_state["page"] = _recon_label
            st.rerun()

    st.markdown("")
    st.divider()

    # ── Capabilities Card ─────────────────────────────────────
    st.markdown('<div class="sidebar-section">Capabilities</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sidebar-about-card">'
        '<div class="sidebar-about-item"><span>📄</span><span><strong>OCR</strong> — extract text from scanned docs</span></div>'
        '<div class="sidebar-about-item"><span>🏷️</span><span><strong>Classify</strong> — identify document types</span></div>'
        '<div class="sidebar-about-item"><span>📊</span><span><strong>Extract</strong> — pull structured financial data</span></div>'
        '<div class="sidebar-about-item"><span>📋</span><span><strong>Report</strong> — generate formatted reports</span></div>'
        '<div class="sidebar-about-item"><span>🔄</span><span><strong>Reconciliation</strong> — matching records to get balances</span></div>'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown("")
    st.divider()

    # ── Footer ─────────────────────────────────────────────────
    st.markdown(
        '<div class="sidebar-footer">'
        f'<span class="sidebar-badge">v1.0</span><br/>'
        f'<span style="color:#6ee7b7;font-size:0.7rem;">Powered by Azure OpenAI &bull; {datetime.now().strftime("%d %b %Y")}</span>'
        '</div>',
        unsafe_allow_html=True,
    )

# ═══════════════════════════════════════════════════════════════════════════
# STATE
# ═══════════════════════════════════════════════════════════════════════════
if "page" not in st.session_state:
    st.session_state["page"] = "🏠 Dashboard"
page = st.session_state["page"]

for key in ("ocr_result", "extraction_result", "doc_type", "uploaded_images", "processing_stage"):
    if key not in st.session_state:
        st.session_state[key] = None
if "doc_status" not in st.session_state:
    st.session_state["doc_status"] = {}  # {row_no: "verified"|"rejected"|"pending"}
if "processing_doc_status" not in st.session_state:
    st.session_state["processing_doc_status"] = {}
if "processing_selected_doc" not in st.session_state:
    st.session_state["processing_selected_doc"] = None
if "extraction_selected_file" not in st.session_state:
    st.session_state["extraction_selected_file"] = None
if "report_preview_source" not in st.session_state:
    st.session_state["report_preview_source"] = None
if "report_detail_row" not in st.session_state:
    st.session_state["report_detail_row"] = None
if "s2_results_ready" not in st.session_state:
    st.session_state["s2_results_ready"] = False


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
if page == "🏠 Dashboard":
    from core.page_tracker import (
        get_usage_by_source, MAX_PAGES,
    )

    st.markdown("### 🏠 Dashboard")
    st.markdown("Overview of pipeline page usage across the entire system.")
    st.markdown("---")

    _by_src    = get_usage_by_source()
    _ocr_used  = _by_src.get('ocr', 0)
    _ocr_remaining = max(0, MAX_PAGES - _ocr_used)
    _pct       = min(_ocr_used / MAX_PAGES, 1.0)

    # ── Quota progress bar (OCR only) ────────────────────────────────────
    st.markdown("#### 📊 OCR Page Quota")
    _bar_label = f"{_ocr_used:,} / {MAX_PAGES:,} pages used  •  {_ocr_remaining:,} remaining"
    st.progress(_pct, text=_bar_label)
    if _pct >= 1.0:
        st.error("⛔ OCR page quota exhausted. No more PDFs can be processed.")
    elif _pct >= 0.9:
        st.warning(f"⚠️ Only {_ocr_remaining:,} OCR pages remaining.")
    elif _pct >= 0.7:
        st.info(f"ℹ️ {_ocr_remaining:,} OCR pages remaining.")
    else:
        st.success(f"✅ {_ocr_remaining:,} OCR pages remaining.")

    st.markdown("")

    # ── KPI tiles ────────────────────────────────────────────────────────
    d1, d2 = st.columns(2)
    d1.metric("OCR Pages Used",      f"{_ocr_used:,}")
    d2.metric("Remaining OCR Pages", f"{_ocr_remaining:,}")

    st.markdown("")

    # ── Document Type Analytics ──────────────────────────────────────────
    _dt_df = _generate_ocr_doc_type_mock()
    _doc_colors = [
        _RC_TEAL, _RC_PURPLE, _RC_PINK, _RC_AMBER, _RC_GREEN,
        _RC_RED, _RC_PURPLE_DARK, _RC_TEAL_DARK,
        _RC_TEAL, _RC_PURPLE, _RC_PINK,
    ]

    _dt_left, _dt_right = st.columns(2)
    with _dt_left:
        _fig_pages = go.Figure(go.Bar(
            x=_dt_df["Document Type"],
            y=_dt_df["Avg Pages Uploaded"],
            marker=dict(color=_doc_colors, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Avg Pages: %{y}<extra></extra>",
        ))
        _fig_pages.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360,
            title=dict(text="Avg Pages Uploaded by Document Type", font=dict(size=14)),
            yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Avg Pages"),
            xaxis=dict(showgrid=False, tickangle=-30),
            margin=dict(l=50, r=20, t=50, b=90),
        )
        st.plotly_chart(_fig_pages, use_container_width=True, config={"displayModeBar": False})

    with _dt_right:
        _fig_fields = go.Figure(go.Bar(
            x=_dt_df["Document Type"],
            y=_dt_df["Avg Fields Extracted"],
            marker=dict(color=_doc_colors, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Avg Fields: %{y}<extra></extra>",
        ))
        _fig_fields.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360,
            title=dict(text="Avg Fields Extracted by Document Type", font=dict(size=14)),
            yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Avg Fields"),
            xaxis=dict(showgrid=False, tickangle=-30),
            margin=dict(l=50, r=20, t=50, b=90),
        )
        st.plotly_chart(_fig_fields, use_container_width=True, config={"displayModeBar": False})

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════
    # OCR SUMMARY
    # ════════════════════════════════════════════════════════════════════
    st.markdown("## 📄 OCR Summary")
    st.caption("Monthly breakdown of documents processed, pages processed, and success/failure rates.")

    _ocr_sum_df = _generate_ocr_summary_mock()

    # ── KPI tiles ────────────────────────────────────────────────────────
    _os_total_docs  = int(_ocr_sum_df["Documents Processed"].sum())
    _os_total_pages = int(_ocr_sum_df["Pages Processed"].sum())
    _os_total_ok    = int(_ocr_sum_df["Successful"].sum())
    _os_total_fail  = int(_ocr_sum_df["Failed"].sum())
    _os_success_pct = _os_total_ok / _os_total_docs * 100 if _os_total_docs else 0

    _ok1, _ok2, _ok3, _ok4, _ok5 = st.columns(5)
    _ok1.markdown(_rc_metric_card("Total Documents",  f"{_os_total_docs:,}",  _RC_TEAL),   unsafe_allow_html=True)
    _ok2.markdown(_rc_metric_card("Total Pages Processed", f"{_os_total_pages:,}", _RC_PURPLE), unsafe_allow_html=True)
    _ok3.markdown(_rc_metric_card("Successful",        f"{_os_total_ok:,}",    _RC_GREEN),  unsafe_allow_html=True)
    _ok4.markdown(_rc_metric_card("Failed",            f"{_os_total_fail:,}",  _RC_RED if _os_total_fail else _RC_GREEN), unsafe_allow_html=True)
    _ok5.markdown(_rc_metric_card("Success Rate",      f"{_os_success_pct:.1f}%", _RC_GREEN if _os_success_pct >= 95 else _RC_AMBER), unsafe_allow_html=True)

    st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

    # ── Monthly volume chart ──────────────────────────────────────────────
    _os_monthly = _ocr_sum_df.groupby("Month", sort=False)[["Documents Processed", "Pages Processed", "Successful", "Failed"]].sum().reset_index()
    _os_left, _os_right = st.columns(2)

    with _os_left:
        _fig_vol_m = go.Figure()
        _fig_vol_m.add_trace(go.Bar(
            x=_os_monthly["Month"], y=_os_monthly["Documents Processed"],
            name="Documents", marker=dict(color=_RC_TEAL, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Documents: %{y}<extra></extra>",
        ))
        _fig_vol_m.add_trace(go.Bar(
            x=_os_monthly["Month"], y=_os_monthly["Pages Processed"],
            name="Pages", marker=dict(color=_RC_PURPLE, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Pages: %{y}<extra></extra>",
        ))
        _fig_vol_m.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360, barmode="group",
            title=dict(text="Monthly Documents & Pages Processed", font=dict(size=14)),
            yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Count"),
            xaxis=dict(showgrid=False),
            legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
            margin=dict(l=50, r=20, t=50, b=70),
        )
        st.plotly_chart(_fig_vol_m, use_container_width=True, config={"displayModeBar": False})

    with _os_right:
        _fig_ok_fail = go.Figure()
        _fig_ok_fail.add_trace(go.Bar(
            x=_os_monthly["Month"], y=_os_monthly["Successful"],
            name="Successful", marker=dict(color=_RC_GREEN, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Successful: %{y}<extra></extra>",
        ))
        _fig_ok_fail.add_trace(go.Bar(
            x=_os_monthly["Month"], y=_os_monthly["Failed"],
            name="Failed", marker=dict(color=_RC_RED, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Failed: %{y}<extra></extra>",
        ))
        _fig_ok_fail.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360, barmode="stack",
            title=dict(text="Monthly Success vs Failed OCR", font=dict(size=14)),
            yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Documents"),
            xaxis=dict(showgrid=False),
            legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
            margin=dict(l=50, r=20, t=50, b=70),
        )
        st.plotly_chart(_fig_ok_fail, use_container_width=True, config={"displayModeBar": False})

    # ── Per document-type breakdown table ────────────────────────────────
    st.markdown(
        f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1.2rem 0 0.8rem 0;'
        f'display:flex;align-items:center;gap:0.5rem;">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_TEAL};display:inline-block;"></span>'
        f'By Document Type</div>', unsafe_allow_html=True,
    )
    _os_by_type = (
        _ocr_sum_df.groupby("Document Type")[["Documents Processed", "Pages Processed", "Successful", "Failed"]]
        .sum().reset_index()
    )
    _os_by_type["Success Rate (%)"] = (_os_by_type["Successful"] / _os_by_type["Documents Processed"] * 100).round(1)
    _os_by_type = _os_by_type.sort_values("Documents Processed", ascending=False).reset_index(drop=True)
    st.dataframe(
        _os_by_type.style.format({"Success Rate (%)": "{:.1f}%", "Documents Processed": "{:,}", "Pages Processed": "{:,}", "Successful": "{:,}", "Failed": "{:,}"}),
        use_container_width=True, hide_index=True,
    )

    st.markdown("---")

    # ════════════════════════════════════════════════════════════════════
    # RECONCILIATION SUMMARY  (mock data — integrated from Recon platform)
    # ════════════════════════════════════════════════════════════════════
    st.markdown("## 🔄 Reconciliation Summary")
    st.caption("Aggregated view of all reconciliation runs — filter, review, and drill into any run.")

    _rc_df = _generate_recon_mock_runs()

    # ── Dashboard Summary ────────────────────────────────────────────────
    st.markdown(
        f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1.2rem 0 0.8rem 0;'
        f'display:flex;align-items:center;gap:0.5rem;">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_TEAL};display:inline-block;"></span>'
        f'Dashboard Summary</div>', unsafe_allow_html=True,
    )
    _rc_total_runs       = len(_rc_df)
    _rc_avg_match        = _rc_df["Match Rate (%)"].mean()
    _rc_needs_review     = int(_rc_df["Needs Review"].sum())
    _rc_all_clear        = _rc_total_runs - _rc_needs_review
    _rc_sum_sales        = _rc_df["Total Sales (RM)"].sum()
    _rc_sum_payment      = _rc_df["Total Payment (RM)"].sum()
    _rc_sum_outstanding  = _rc_df["Total Outstanding (RM)"].sum()
    _rc_avg_duration     = _rc_df["Duration (s)"].mean()

    _rk1, _rk2, _rk3, _rk4 = st.columns(4)
    _rk1.markdown(_rc_metric_card("Total Runs",    f"{_rc_total_runs}", _RC_TEAL),  unsafe_allow_html=True)
    _rk2.markdown(_rc_metric_card("Avg Match Rate", f"{_rc_avg_match:.1f}%", _RC_GREEN if _rc_avg_match >= 95 else _RC_AMBER), unsafe_allow_html=True)
    _rk3.markdown(_rc_metric_card("Needs Review",  f"{_rc_needs_review}", _RC_RED if _rc_needs_review else _RC_GREEN), unsafe_allow_html=True)
    _rk4.markdown(_rc_metric_card("All Clear",     f"{_rc_all_clear}", _RC_GREEN), unsafe_allow_html=True)
    st.markdown("<div style='height:0.6rem'></div>", unsafe_allow_html=True)
    _rk5, _rk6, _rk7, _rk8 = st.columns(4)
    _rk5.markdown(_rc_metric_card("Total Sales",       f"RM {_rc_sum_sales:,.2f}",       _RC_PURPLE), unsafe_allow_html=True)
    _rk6.markdown(_rc_metric_card("Total Payment",     f"RM {_rc_sum_payment:,.2f}",     _RC_TEAL),   unsafe_allow_html=True)
    _rk7.markdown(_rc_metric_card("Total Outstanding", f"RM {_rc_sum_outstanding:,.2f}", _RC_RED),    unsafe_allow_html=True)
    _rk8.markdown(_rc_metric_card("Avg Duration",      f"{_rc_avg_duration:.1f}s",       _RC_PINK),   unsafe_allow_html=True)

    # ── Trends ───────────────────────────────────────────────────────────
    st.markdown(
        f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1.2rem 0 0.8rem 0;'
        f'display:flex;align-items:center;gap:0.5rem;">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_PURPLE};display:inline-block;"></span>'
        f'Trends</div>', unsafe_allow_html=True,
    )
    _rc_trend = _rc_df.sort_values("Run Date")
    _rt_left, _rt_right = st.columns(2)
    with _rt_left:
        _fig_match = go.Figure()
        _fig_match.add_trace(go.Scatter(
            x=_rc_trend["Run Date"], y=_rc_trend["Match Rate (%)"],
            mode="lines+markers", name="Match Rate",
            line=dict(color=_RC_TEAL, width=2.5), marker=dict(size=7),
            hovertemplate="<b>%{x}</b><br>Match Rate: %{y:.1f}%<extra></extra>",
        ))
        _fig_match.add_hline(y=95, line_dash="dash", line_color=_RC_GREEN,
                             annotation_text="Target 95%", annotation_position="top left")
        _fig_match.add_hline(y=92, line_dash="dot",  line_color=_RC_RED,
                             annotation_text="Review 92%", annotation_position="bottom left")
        _fig_match.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360,
            title=dict(text="Match Rate Over Time", font=dict(size=14)),
            yaxis=dict(range=[80, 102], showgrid=True, gridcolor="#F0F0F0", title="Match Rate (%)"),
            xaxis=dict(showgrid=False, title="Run Date"),
            margin=dict(l=50, r=20, t=50, b=50),
        )
        st.plotly_chart(_fig_match, use_container_width=True, config={"displayModeBar": False})
    with _rt_right:
        _fig_vol = go.Figure()
        _fig_vol.add_trace(go.Bar(
            x=_rc_trend["Run Date"], y=_rc_trend["Total Sales (RM)"],
            name="Sales", marker=dict(color=_RC_PURPLE, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Sales: RM %{y:,.0f}<extra></extra>",
        ))
        _fig_vol.add_trace(go.Bar(
            x=_rc_trend["Run Date"], y=_rc_trend["Total Payment (RM)"],
            name="Payment", marker=dict(color=_RC_TEAL, cornerradius=4),
            hovertemplate="<b>%{x}</b><br>Payment: RM %{y:,.0f}<extra></extra>",
        ))
        _fig_vol.update_layout(
            **_RC_PLOTLY_LAYOUT, height=360, barmode="group",
            title=dict(text="Sales vs Payment Per Run", font=dict(size=14)),
            yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Amount (RM)"),
            xaxis=dict(showgrid=False, title="Run Date"),
            legend=dict(orientation="h", yanchor="top", y=-0.18, xanchor="center", x=0.5),
            margin=dict(l=60, r=20, t=50, b=70),
        )
        st.plotly_chart(_fig_vol, use_container_width=True, config={"displayModeBar": False})

    # ── Reconciliation Run History ───────────────────────────────────────
    st.markdown(
        f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1.2rem 0 0.8rem 0;'
        f'display:flex;align-items:center;gap:0.5rem;">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_PINK};display:inline-block;"></span>'
        f'Reconciliation Run History</div>', unsafe_allow_html=True,
    )
    _rc_disp_cols = [
        "Run ID", "Run Date", "Period From", "Period To", "Status",
        "Match Rate (%)", "Recon Rows", "Outstanding Orders",
        "Total Sales (RM)", "Total Outstanding (RM)", "Duration (s)",
    ]
    _rc_disp_df = _rc_df[_rc_disp_cols].sort_values("Run Date", ascending=False).reset_index(drop=True)
    def _rc_highlight(row):
        if row["Status"] == "⚠️ Needs Review":
            return ["background-color:#FEF3C7;color:#92400E;"] * len(row)
        return [""] * len(row)
    _rc_styled = _rc_disp_df.style.apply(_rc_highlight, axis=1).format({
        "Match Rate (%)": "{:.2f}",
        "Total Sales (RM)": "RM {:,.2f}",
        "Total Outstanding (RM)": "RM {:,.2f}",
        "Duration (s)": "{:.2f}s",
    })
    st.markdown(
        f'<div style="display:flex;align-items:center;gap:0.5rem;margin:0.5rem 0;">'
        f'<span style="background:{_RC_TEAL};color:white;padding:0.15rem 0.6rem;border-radius:12px;font-size:0.78rem;font-weight:600;">'
        f'{len(_rc_disp_df):,}</span>'
        f'<span style="color:{_RC_MUTED};font-size:0.82rem;">runs shown</span>'
        f'<span style="background:{_RC_RED};color:white;padding:0.15rem 0.6rem;border-radius:12px;font-size:0.78rem;font-weight:600;margin-left:0.5rem;">'
        f'{_rc_needs_review}</span>'
        f'<span style="color:{_RC_MUTED};font-size:0.82rem;">need review</span></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(_rc_styled, use_container_width=True, height=400)

    # ── Run Detail View ──────────────────────────────────────────────────
    st.markdown(
        f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1.2rem 0 0.8rem 0;'
        f'display:flex;align-items:center;gap:0.5rem;">'
        f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_TEAL};display:inline-block;"></span>'
        f'Run Detail View</div>', unsafe_allow_html=True,
    )
    _rc_run_ids = _rc_df.sort_values("Run Date", ascending=False)["Run ID"].tolist()
    _rc_sel_run = st.selectbox("Select a run to inspect", _rc_run_ids, key="dash_rc_detail_run")

    if _rc_sel_run:
        _rc_row = _rc_df[_rc_df["Run ID"] == _rc_sel_run].iloc[0]
        _rc_is_review = _rc_row["Needs Review"]
        _rc_all_sorted = _rc_df.sort_values("Run Date", ascending=True).reset_index(drop=True)
        _rc_cur_idx = _rc_all_sorted[_rc_all_sorted["Run ID"] == _rc_sel_run].index[0]
        _rc_prev_row = _rc_all_sorted.iloc[_rc_cur_idx - 1] if _rc_cur_idx > 0 else None

        def _rc_delta(cur, prev, fmt=",.2f", prefix="", suffix="", invert=False):
            if prev is None:
                return ""
            diff = cur - prev
            if diff == 0:
                return f'<span style="font-size:0.72rem;color:{_RC_MUTED};margin-left:0.3rem;">— vs prev</span>'
            is_up = diff > 0
            color = (_RC_RED if is_up else _RC_GREEN) if invert else (_RC_GREEN if is_up else _RC_RED)
            arrow = "▲" if is_up else "▼"
            return f'<span style="font-size:0.72rem;color:{color};margin-left:0.3rem;">{arrow} {prefix}{abs(diff):{fmt}}{suffix} vs prev</span>'

        _rc_status_color = _RC_RED if _rc_is_review else _RC_GREEN
        st.markdown(f"""
        <div style="background:#fff;border:2px solid {_rc_status_color};border-radius:12px;
                    padding:1.25rem 1.5rem;margin-bottom:1rem;">
            <div style="display:flex;justify-content:space-between;align-items:center;">
                <div>
                    <span style="font-size:1.3rem;font-weight:700;color:{_RC_TEXT};">{_rc_row['Run ID']}</span>
                    <span style="margin-left:1rem;font-size:0.88rem;color:{_RC_MUTED};">Run Date: {_rc_row['Run Date']}</span>
                </div>
                <div style="background:{_rc_status_color};color:white;padding:0.3rem 1rem;
                            border-radius:20px;font-weight:600;font-size:0.85rem;">{_rc_row['Status']}</div>
            </div>
            <div style="margin-top:0.5rem;font-size:0.88rem;color:{_RC_MUTED};">
                Period: {_rc_row['Period From']} &rarr; {_rc_row['Period To']} &nbsp;|&nbsp; Duration: {_rc_row['Duration (s)']:.2f}s
            </div>
        </div>
        """, unsafe_allow_html=True)
        if _rc_is_review:
            st.warning(f"**Review Reasons:** {_rc_row['Review Reasons']}")

        # ── Financial Health ──────────────────────────────────────────
        st.markdown(
            f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1rem 0 0.6rem 0;'
            f'display:flex;align-items:center;gap:0.5rem;">'
            f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_PURPLE};display:inline-block;"></span>'
            f'Financial Health</div>', unsafe_allow_html=True,
        )
        _rc_match_val  = _rc_row["Match Rate (%)"]
        _rc_gauge_col  = _RC_TEAL if _rc_match_val >= 95 else (_RC_AMBER if _rc_match_val >= 92 else _RC_RED)
        _rc_prev_sales = _rc_prev_row["Total Sales (RM)"]       if _rc_prev_row is not None else None
        _rc_prev_pay   = _rc_prev_row["Total Payment (RM)"]     if _rc_prev_row is not None else None
        _rc_prev_fees  = _rc_prev_row["Total Fees (RM)"]        if _rc_prev_row is not None else None
        _rc_prev_out   = _rc_prev_row["Total Outstanding (RM)"] if _rc_prev_row is not None else None

        _fh1, _op1, _fh2, _op2, _fh3, _op3, _fh4 = st.columns([3, 0.5, 3, 0.5, 3, 0.5, 3])
        _fh1.markdown(_rc_metric_card("TOTAL SALES",   f"RM {_rc_row['Total Sales (RM)']:,.2f}", _RC_TEAL) +
                      _rc_delta(_rc_row['Total Sales (RM)'], _rc_prev_sales, prefix='RM '), unsafe_allow_html=True)
        _op1.markdown(f'<div style="display:flex;align-items:center;justify-content:center;font-size:1.5rem;font-weight:700;color:{_RC_MUTED};padding-top:1rem;">−</div>', unsafe_allow_html=True)
        _fh2.markdown(_rc_metric_card("TOTAL PAYMENT", f"RM {_rc_row['Total Payment (RM)']:,.2f}", _RC_PURPLE) +
                      _rc_delta(_rc_row['Total Payment (RM)'], _rc_prev_pay, prefix='RM '), unsafe_allow_html=True)
        _op2.markdown(f'<div style="display:flex;align-items:center;justify-content:center;font-size:1.5rem;font-weight:700;color:{_RC_MUTED};padding-top:1rem;">−</div>', unsafe_allow_html=True)
        _fh3.markdown(_rc_metric_card("TOTAL FEES",    f"RM {_rc_row['Total Fees (RM)']:,.2f}", _RC_AMBER) +
                      _rc_delta(_rc_row['Total Fees (RM)'], _rc_prev_fees, prefix='RM ', invert=True), unsafe_allow_html=True)
        _op3.markdown(f'<div style="display:flex;align-items:center;justify-content:center;font-size:1.5rem;font-weight:700;color:{_RC_RED};padding-top:1rem;">=</div>', unsafe_allow_html=True)
        _rc_out_color = _RC_GREEN if _rc_row['Total Outstanding (RM)'] == 0 else _RC_RED
        _fh4.markdown(_rc_metric_card("OUTSTANDING", f"RM {_rc_row['Total Outstanding (RM)']:,.2f}", _rc_out_color) +
                      _rc_delta(_rc_row['Total Outstanding (RM)'], _rc_prev_out, prefix='RM ', invert=True), unsafe_allow_html=True)

        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        _g_l, _g_c, _g_r = st.columns([1, 2, 1])
        with _g_c:
            _fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number", value=_rc_match_val,
                number=dict(suffix="%", font=dict(size=36)),
                title=dict(text="Match Rate", font=dict(size=14)),
                gauge=dict(
                    axis=dict(range=[0, 100]), bar=dict(color=_rc_gauge_col), bgcolor="#F0F0F0",
                    steps=[
                        dict(range=[0, 92],  color="#FDE8E8"),
                        dict(range=[92, 95], color="#FEF3C7"),
                        dict(range=[95, 100],color="#D1FAE5"),
                    ],
                ),
            ))
            _fig_gauge.update_layout(**_RC_PLOTLY_LAYOUT, height=260, margin=dict(l=30, r=30, t=60, b=10))
            st.plotly_chart(_fig_gauge, use_container_width=True, config={"displayModeBar": False})

        # ── Fee Analysis & Outstanding Trend ─────────────────────────
        st.markdown(
            f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1rem 0 0.6rem 0;'
            f'display:flex;align-items:center;gap:0.5rem;">'
            f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_AMBER};display:inline-block;"></span>'
            f'Fee Analysis & Outstanding Trend</div>', unsafe_allow_html=True,
        )
        _fee_col, _trend_col = st.columns(2)
        _rc_fees_pct = _rc_row.get("Fees % of Sales", 0)
        if _rc_fees_pct == 0 and _rc_row["Total Sales (RM)"] > 0:
            _rc_fees_pct = round(_rc_row["Total Fees (RM)"] / _rc_row["Total Sales (RM)"] * 100, 2)
        _rc_fees_color = _RC_GREEN if _rc_fees_pct <= 8 else (_RC_AMBER if _rc_fees_pct <= 12 else _RC_RED)
        _rc_prev_fees_pct = None
        if _rc_prev_row is not None and _rc_prev_row["Total Sales (RM)"] > 0:
            _rc_prev_fees_pct = round(_rc_prev_row["Total Fees (RM)"] / _rc_prev_row["Total Sales (RM)"] * 100, 2)

        with _fee_col:
            _fig_fee = go.Figure(go.Indicator(
                mode="gauge+number+delta", value=_rc_fees_pct,
                number=dict(suffix="%", font=dict(size=34)),
                title=dict(text="Fees % of Sales", font=dict(size=14)),
                delta=dict(
                    reference=_rc_prev_fees_pct if _rc_prev_fees_pct is not None else _rc_fees_pct,
                    increasing=dict(color=_RC_RED), decreasing=dict(color=_RC_GREEN),
                    suffix="%", font=dict(size=14),
                ),
                gauge=dict(
                    axis=dict(range=[0, 20]), bar=dict(color=_rc_fees_color), bgcolor="#F0F0F0",
                    steps=[
                        dict(range=[0, 5],   color="#D1FAE5"),
                        dict(range=[5, 8],   color="#E8FFE8"),
                        dict(range=[8, 12],  color="#FEF3C7"),
                        dict(range=[12, 20], color="#FDE8E8"),
                    ],
                    threshold=dict(line=dict(color=_RC_RED, width=3), thickness=0.8, value=12),
                ),
            ))
            _fig_fee.update_layout(**_RC_PLOTLY_LAYOUT, height=300, margin=dict(l=30, r=30, t=60, b=20))
            st.plotly_chart(_fig_fee, use_container_width=True, config={"displayModeBar": False})
            _fee_status = ("Within normal range" if _rc_fees_pct <= 8 else
                           ("Above average — verify contract" if _rc_fees_pct <= 12 else "Exceeds threshold — investigate"))
            _fee_icon = "✅" if _rc_fees_pct <= 8 else ("⚠️" if _rc_fees_pct <= 12 else "🚨")
            st.markdown(
                f'<div style="background:#fff;border:1px solid #E5E7EB;border-radius:8px;padding:0.75rem 1rem;font-size:0.85rem;">'
                f'{_fee_icon} <b>Fee Rate:</b> {_rc_fees_pct:.2f}% &nbsp;|&nbsp; '
                f'<b>Expected Shopee range:</b> 3–8% &nbsp;|&nbsp; <b>Assessment:</b> {_fee_status}</div>',
                unsafe_allow_html=True,
            )

        with _trend_col:
            _rc_trend_out = _rc_all_sorted[["Run Date", "Total Outstanding (RM)", "Run ID"]].tail(10)
            _bar_colors = [_RC_GREEN if v == 0 else _RC_RED for v in _rc_trend_out["Total Outstanding (RM)"]]
            _fig_out = go.Figure()
            _fig_out.add_trace(go.Bar(
                x=_rc_trend_out["Run Date"].astype(str),
                y=_rc_trend_out["Total Outstanding (RM)"],
                marker=dict(color=_bar_colors, cornerradius=4),
                text=[f"RM {v:,.0f}" for v in _rc_trend_out["Total Outstanding (RM)"]],
                textposition="outside", textfont=dict(size=10),
                hovertemplate="<b>%{x}</b><br>Outstanding: RM %{y:,.2f}<extra></extra>",
            ))
            _fig_out.add_hline(y=0, line_color=_RC_GREEN, line_width=2)
            _fig_out.update_layout(
                **_RC_PLOTLY_LAYOUT, height=300,
                title=dict(text="Outstanding Amount (Last Runs)", font=dict(size=14)),
                yaxis=dict(showgrid=True, gridcolor="#F0F0F0", title="Outstanding (RM)"),
                xaxis=dict(showgrid=False, title="Run Date", tickangle=-45),
                margin=dict(l=60, r=20, t=50, b=70),
            )
            st.plotly_chart(_fig_out, use_container_width=True, config={"displayModeBar": False})
            _last_two = _rc_trend_out["Total Outstanding (RM)"].tail(2).tolist()
            if len(_last_two) >= 2:
                if _last_two[-1] == 0:
                    st.success("✅ Outstanding is zero this run — fully reconciled.")
                elif _last_two[-1] < _last_two[-2]:
                    st.info(f"ℹ️ Outstanding decreased from RM {_last_two[-2]:,.2f} to RM {_last_two[-1]:,.2f} — improving.")
                elif _last_two[-1] > _last_two[-2]:
                    st.warning(f"⚠️ Outstanding increased from RM {_last_two[-2]:,.2f} to RM {_last_two[-1]:,.2f} — investigate.")
                else:
                    st.info(f"— Outstanding unchanged at RM {_last_two[-1]:,.2f}.")

        # ── Exceptions & Data Quality ─────────────────────────────────
        st.markdown(
            f'<div style="font-size:1.15rem;font-weight:700;color:{_RC_TEXT};margin:1rem 0 0.6rem 0;'
            f'display:flex;align-items:center;gap:0.5rem;">'
            f'<span style="width:10px;height:10px;border-radius:50%;background:{_RC_RED};display:inline-block;"></span>'
            f'Exceptions & Data Quality</div>', unsafe_allow_html=True,
        )
        _exc1, _exc2, _exc3, _exc4 = st.columns(4)
        _exc1.markdown(_rc_metric_card("OUTSTANDING ORDERS",    f"{_rc_row['Outstanding Orders']:,}",    _RC_RED if _rc_row['Outstanding Orders'] > 0 else _RC_GREEN), unsafe_allow_html=True)
        _exc2.markdown(_rc_metric_card("REFUND ORDERS",         f"{_rc_row['Refund Orders']:,}",         _RC_PINK), unsafe_allow_html=True)
        _exc3.markdown(_rc_metric_card("INCOME NOT IN BALANCE", f"{_rc_row['Income Not In Balance']:,}", _RC_RED_DARK), unsafe_allow_html=True)
        _exc4.markdown(_rc_metric_card("BALANCE NOT IN INCOME", f"{_rc_row['Balance Not In Income']:,}", _RC_PURPLE_DARK), unsafe_allow_html=True)

        st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
        with st.expander("📋 Raw Data Counts"):
            _raw1, _raw2, _raw3, _raw4 = st.columns(4)
            _raw1.markdown(_rc_metric_card("Income Rows",  f"{_rc_row['Income Rows']:,}",  _RC_PURPLE),      unsafe_allow_html=True)
            _raw2.markdown(_rc_metric_card("Balance Rows", f"{_rc_row['Balance Rows']:,}", _RC_TEAL),        unsafe_allow_html=True)
            _raw3.markdown(_rc_metric_card("Sales Rows",   f"{_rc_row['Sales Rows']:,}",   _RC_PINK),        unsafe_allow_html=True)
            _raw4.markdown(_rc_metric_card("Recon Rows",   f"{_rc_row['Recon Rows']:,}",   _RC_PURPLE_DARK), unsafe_allow_html=True)

    st.markdown("---")


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: DOCUMENT PROCESSING
# ═══════════════════════════════════════════════════════════════════════════
if page == "📤 Document Processing":

    st.markdown("### 📤 Document Processing Pipeline")
    st.markdown("Upload a PDF to run the full AI pipeline: **PDF → Images → OCR → Classification → Extraction**")

    _proc_mode = st.radio(
        "Document source",
        ["Upload new document", "Re-process existing document"],
        horizontal=True,
        key="proc_mode",
    )

    _uploads_dir_early = Path(__file__).resolve().parent / "docs" / "uploads"
    _uploads_dir_early.mkdir(parents=True, exist_ok=True)

    col_up, col_opt = st.columns([2, 1])
    with col_up:
        if _proc_mode == "Upload new document":
            uploaded_file = st.file_uploader(
                "Upload a PDF document", type=["pdf"],
                help="Invoices, Utility Bills, Bank Statements, Travel, Rental, SOA, etc.",
            )
            _existing_selection = None
        else:
            uploaded_file = None
            _existing_pdfs = sorted(
                [p for p in _uploads_dir_early.iterdir()
                 if p.is_file() and p.suffix.lower() == ".pdf"],
                key=lambda p: p.stat().st_mtime, reverse=True,
            )
            if _existing_pdfs:
                _existing_selection = st.selectbox(
                    "Select previously uploaded document",
                    options=[p.name for p in _existing_pdfs],
                    key="proc_existing_selector",
                )
            else:
                st.info("No previously uploaded documents found. Upload a document first.")
                _existing_selection = None
    with col_opt:
        force_type = st.selectbox("Force document type (optional)", ["Auto-detect"] + list(AGENT_REGISTRY.keys()))
        ocr_mode = st.radio("OCR Mode", ["Batch (all pages)", "Per-page"], index=0)
        upload_team_choice = st.selectbox("Document Team", ["Auto", "Sales", "Rental"], index=0)

    # ── Resolve active document (new upload OR existing selection) ────────────
    _active_bytes: bytes | None = None
    _active_upload_path: Path | None = None

    if _proc_mode == "Upload new document" and uploaded_file is not None:
        app_dir = Path(__file__).resolve().parent
        uploads_dir = app_dir / "docs" / "uploads"
        uploads_dir.mkdir(parents=True, exist_ok=True)

        _upload_key = f"_saved_upload_{getattr(uploaded_file, 'file_id', None) or (uploaded_file.name + str(uploaded_file.size))}"
        if _upload_key not in st.session_state:
            upload_path = uploads_dir / uploaded_file.name
            if upload_path.exists():
                counter = 1
                while True:
                    candidate = uploads_dir / f"{Path(uploaded_file.name).stem}_{counter}{Path(uploaded_file.name).suffix}"
                    if not candidate.exists():
                        upload_path = candidate
                        break
                    counter += 1
            upload_path.write_bytes(uploaded_file.getvalue())

            doc_team_map = load_doc_team_map()
            assigned_team = upload_team_choice.lower()
            if assigned_team == "auto":
                assigned_team = _team_from_doc_type(force_type if force_type != "Auto-detect" else "")
            doc_team_map[upload_path.name] = assigned_team
            save_doc_team_map(doc_team_map)
            st.session_state[_upload_key] = str(upload_path)

        _active_upload_path = Path(st.session_state[_upload_key])
        _active_bytes = _active_upload_path.read_bytes()

    elif _proc_mode == "Re-process existing document" and _existing_selection:
        _active_upload_path = _uploads_dir_early / _existing_selection
        _active_bytes = _active_upload_path.read_bytes()

    if _active_bytes is not None and _active_upload_path is not None:
        upload_path = _active_upload_path
        uploaded_bytes = _active_bytes
        assigned_team = load_doc_team_map().get(upload_path.name, "sales")

        with tempfile.TemporaryDirectory() as tmp_dir:
            pdf_path = Path(tmp_dir) / upload_path.name
            pdf_path.write_bytes(uploaded_bytes)
            st.markdown(f"**📄 Document:** `{upload_path.name}` ({len(uploaded_bytes) / 1024:.1f} KB)")
            st.caption(f"Stored in uploads: `{upload_path.name}` | Team: `{assigned_team.title()}`")

            if st.button("🚀 Run Full Pipeline", type="primary", use_container_width=True):
                with st.status("🔄 Processing document...", expanded=True) as status:
                    progress = st.progress(0)

                    # Step 1: PDF to Images
                    st.write("**Step 1/4:** Converting PDF to images...")
                    try:
                        image_dir = Path(__file__).resolve().parent / "output" / "images" / pdf_path.stem
                        image_dir.mkdir(parents=True, exist_ok=True)
                        image_paths = pdf_to_images(pdf_path, image_dir)
                        progress.progress(20)
                        st.write(f"  ✅ Converted to **{len(image_paths)} page(s)**")
                        tcols = st.columns(min(len(image_paths), 6))
                        for i, ip in enumerate(image_paths[:6]):
                            with tcols[i]:
                                st.image(str(ip), caption=f"Page {i+1}", width=110)
                        # ── Track pages in quota ────────────────────────────
                        from core.page_tracker import add_pages as _add_pages
                        _add_pages(len(image_paths), source="ocr")
                    except Exception as e:
                        st.error(f"❌ PDF conversion failed: {e}")
                        st.stop()

                    # Step 2: OCR
                    st.write("**Step 2/4:** Running AI-powered OCR...")
                    try:
                        user_prompt = (
                            "Transcribe ALL visible text from this document image exactly as it appears. "
                            "Output the result as a single valid JSON object following the schema in your instructions. "
                            "Do NOT interpret, summarize, or calculate anything. "
                            "Preserve all numbers, punctuation, and formatting exactly."
                        )

                        _total_pages = len(image_paths)
                        st.write(f"  📄 Total pages to process: **{_total_pages}**")

                        if ocr_mode.startswith("Batch"):
                            # ── Batch: send all images, then validate completeness ─
                            st.write(f"  📤 Sending all {_total_pages} page(s) to OCR in one request...")
                            raw_ocr = ocr_images_with_chat_model(image_paths, user_prompt)
                            ocr_parsed = _maybe_parse_json(raw_ocr)

                            # ── Page count validation ────────────────────────────
                            _returned_pages = []
                            if isinstance(ocr_parsed, dict):
                                _returned_pages = ocr_parsed.get("pages", [])
                            _returned_count = len(_returned_pages)
                            _returned_nums = {p.get("page_number") for p in _returned_pages if isinstance(p, dict)}

                            st.write(f"  📥 Pages received: **{_returned_count}** / {_total_pages}")

                            if _returned_count < _total_pages:
                                _missing_indices = [
                                    i for i, ip in enumerate(image_paths, start=1)
                                    if i not in _returned_nums
                                ]
                                st.warning(
                                    f"  ⚠️ Only {_returned_count}/{_total_pages} pages returned. "
                                    f"Re-processing {len(_missing_indices)} missing page(s) individually: {_missing_indices}"
                                )

                                for _miss_i in _missing_indices:
                                    _miss_path = image_paths[_miss_i - 1]
                                    st.write(f"  🔄 Retrying page {_miss_i} ({_miss_path.name})...")
                                    _raw_single = ocr_image_with_chat_model(_miss_path, user_prompt)
                                    _single_parsed = _maybe_parse_json(_raw_single)

                                    # Extract page data from per-page result and inject
                                    _new_page = None
                                    if isinstance(_single_parsed, dict):
                                        _sp_pages = _single_parsed.get("pages", [])
                                        if _sp_pages:
                                            _new_page = _sp_pages[0]
                                            _new_page["page_number"] = _miss_i
                                            _new_page["file_name"] = _miss_path.name
                                        else:
                                            # Model returned flat structure — wrap it
                                            _new_page = {
                                                "page_number": _miss_i,
                                                "file_name": _miss_path.name,
                                                "sections": _single_parsed.get("sections", []),
                                                "_raw": _single_parsed,
                                            }
                                    else:
                                        _new_page = {
                                            "page_number": _miss_i,
                                            "file_name": _miss_path.name,
                                            "sections": [],
                                            "_raw_text": str(_single_parsed),
                                        }

                                    if isinstance(ocr_parsed, dict) and "pages" in ocr_parsed:
                                        ocr_parsed["pages"].append(_new_page)
                                    st.write(f"    ✅ Page {_miss_i} recovered")

                                # Re-sort pages by page_number
                                if isinstance(ocr_parsed, dict) and "pages" in ocr_parsed:
                                    ocr_parsed["pages"].sort(key=lambda p: p.get("page_number", 999))
                                    # Update metadata
                                    if "metadata" in ocr_parsed:
                                        ocr_parsed["metadata"]["total_pages"] = len(ocr_parsed["pages"])
                                raw_ocr = json.dumps(ocr_parsed, ensure_ascii=False)

                            _final_count = len(ocr_parsed.get("pages", [])) if isinstance(ocr_parsed, dict) else 0
                            if _final_count == _total_pages:
                                st.write(f"  ✅ All {_total_pages} page(s) successfully OCR'd")
                            else:
                                st.warning(f"  ⚠️ Final page count: {_final_count}/{_total_pages} — some pages may still be missing")

                            ocr_json_str = raw_ocr if isinstance(raw_ocr, str) else json.dumps(ocr_parsed, ensure_ascii=False)

                        else:
                            # ── Per-page mode ─────────────────────────────────────
                            pages_list = []
                            for idx, ip in enumerate(image_paths):
                                st.write(f"  OCR page {idx+1}/{_total_pages} ({ip.name})...")
                                raw = ocr_image_with_chat_model(ip, user_prompt)
                                pages_list.append({"page_number": idx+1, "file": ip.name, "model_output": _maybe_parse_json(raw)})
                            ocr_parsed = {"mode": "per_image", "results": pages_list}
                            ocr_json_str = json.dumps(ocr_parsed, ensure_ascii=False)
                            st.write(f"  ✅ All {_total_pages} page(s) processed")

                        st.session_state.ocr_result = ocr_parsed
                        progress.progress(60)
                        st.write("  ✅ OCR complete")
                    except Exception as e:
                        st.error(f"❌ OCR failed: {e}")
                        st.stop()

                    # Step 3 & 4: Classify + Extract
                    st.write("**Step 3/4:** Classifying & extracting...")
                    try:
                        forced = None if force_type == "Auto-detect" else force_type
                        doc_type_result, extracted = orchestrator_run(ocr_json_str, forced_type=forced)
                        st.session_state.doc_type = doc_type_result
                        st.session_state.extraction_result = extracted
                        progress.progress(95)
                        st.write(f"  ✅ Classified as: **{doc_type_result.replace('_',' ').title()}**")
                    except Exception as e:
                        st.error(f"❌ Extraction failed: {e}")
                        st.stop()

                    # ── Duplicate invoice check ──────────────────────────────
                    _new_inv_no = (
                        extracted.get("invoice_number")
                        or extracted.get("document_number")
                        or extracted.get("statement_number")
                    ) if isinstance(extracted, dict) else None

                    _dup_source = None
                    if _new_inv_no:
                        _ext_dir = Path(__file__).resolve().parent / "output" / "extraction"
                        for _ef in _ext_dir.glob("*.json"):
                            try:
                                _ed = json.loads(_ef.read_text(encoding="utf-8"))
                                _existing_inv = (
                                    _ed.get("invoice_number")
                                    or _ed.get("document_number")
                                    or _ed.get("statement_number")
                                )
                                if _existing_inv and str(_existing_inv).strip() == str(_new_inv_no).strip():
                                    _dup_source = _ef.name
                                    break
                            except Exception:
                                pass

                    if _dup_source:
                        st.info(
                            f"ℹ️ Invoice **{_new_inv_no}** already exists (source: `{_dup_source}`). "
                            f"The existing record will be **overwritten** with the new extraction."
                        )
                        # Delete the old extraction file so the new one replaces it
                        _old_ef = Path(__file__).resolve().parent / "output" / "extraction" / _dup_source
                        try:
                            _old_ef.unlink(missing_ok=True)
                        except Exception:
                            pass

                    progress.progress(100)
                    status.update(label="✅ Pipeline complete!", state="complete", expanded=True)

                st.divider()
                st.markdown("### 📋 Results")
                tab_ext, tab_ocr, tab_json = st.tabs(["📊 Extracted Data", "🔍 OCR Output", "📝 Raw JSON"])
                with tab_ext:
                    display_extraction_result(extracted, doc_type_result)
                with tab_ocr:
                    display_ocr_result(ocr_parsed if isinstance(ocr_parsed, dict) else {"raw": ocr_parsed})
                with tab_json:
                    jc1, jc2 = st.columns(2)
                    with jc1:
                        st.markdown("**OCR Output**")
                        st.json(ocr_parsed)
                    with jc2:
                        st.markdown("**Extraction Output**")
                        st.json(extracted)

                save_base_name = upload_path.stem  # use the (possibly renamed) upload filename
                app_dir = Path(__file__).resolve().parent
                ocr_output_dir = app_dir / "output" / "ocr"
                extraction_output_dir = app_dir / "output" / "extraction"
                ocr_output_dir.mkdir(parents=True, exist_ok=True)
                extraction_output_dir.mkdir(parents=True, exist_ok=True)

                ocr_output_path = ocr_output_dir / f"{save_base_name}.json"
                extraction_output_path = extraction_output_dir / f"{save_base_name}.json"

                with open(ocr_output_path, "w", encoding="utf-8") as f:
                    ocr_with_meta = ocr_parsed if isinstance(ocr_parsed, dict) else {"raw": ocr_parsed}
                    if isinstance(ocr_with_meta, dict):
                        ocr_with_meta = {**ocr_with_meta, "_document_type": doc_type_result}
                    json.dump(ocr_with_meta, f, ensure_ascii=False, indent=2)
                _upload_ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                extracted_with_ts = {**extracted, "_uploaded_at": _upload_ts} if isinstance(extracted, dict) else extracted
                with open(extraction_output_path, "w", encoding="utf-8") as f:
                    json.dump(extracted_with_ts, f, ensure_ascii=False, indent=2)

                st.success(
                    f"Saved results to `{ocr_output_path.name}` and `{extraction_output_path.name}`. "
                    "These are now available in OCR Viewer, Extraction Viewer, and Report Format."
                )

                dc1, dc2 = st.columns(2)
                with dc1:
                    st.download_button("⬇️ Download OCR JSON",
                        data=json.dumps(ocr_parsed, ensure_ascii=False, indent=2),
                        file_name=f"{pdf_path.stem}_ocr.json", mime="application/json")
                with dc2:
                    st.download_button("⬇️ Download Extraction JSON",
                        data=json.dumps(extracted, ensure_ascii=False, indent=2),
                        file_name=f"{pdf_path.stem}_extracted.json", mime="application/json")

    st.divider()

    st.markdown("### Documents")
    st.caption("Upload and manage document processing")

    uploads_dir = Path(__file__).resolve().parent / "docs" / "uploads"
    source_docs = (
        sorted([p for p in uploads_dir.iterdir() if p.is_file() and p.suffix.lower() in {".pdf", ".png", ".jpg", ".jpeg", ".webp", ".tif", ".tiff"}], key=lambda p: p.stat().st_mtime, reverse=True)
        if uploads_dir.exists()
        else []
    )
    doc_team_map = load_doc_team_map()
    source_docs_with_team = [(p, infer_document_team(p, doc_team_map)) for p in source_docs]
    visible_source_docs = source_docs_with_team

    st.markdown(f"#### All Documents ({len(visible_source_docs)})")

    if visible_source_docs:
        h1, h2, h3, h4, h5, h6, h7 = st.columns([1.2, 3.2, 1.0, 1.0, 1.3, 1.7, 2.4])
        h1.markdown("**Doc ID**")
        h2.markdown("**File Name**")
        h3.markdown("**Type**")
        h4.markdown("**Size (MB)**")
        h5.markdown("**Upload Date**")
        h6.markdown("**Status**")
        h7.markdown("**Actions**")

        for idx, (file_path, _team) in enumerate(visible_source_docs, start=1):
            display_file_name = file_path.name
            file_type = file_path.suffix.replace(".", "").upper() or "FILE"
            file_size_mb = file_path.stat().st_size / (1024 * 1024)
            upload_date = datetime.fromtimestamp(file_path.stat().st_mtime).strftime("%Y-%m-%d")
            status = st.session_state["processing_doc_status"].get(file_path.name, "Ready for Review")

            c1, c2, c3, c4, c5, c6, c7 = st.columns([1.2, 3.2, 1.0, 1.0, 1.3, 1.7, 2.4])
            c1.markdown(f"**DOC-{idx:04d}**")
            c2.markdown(display_file_name)
            c3.markdown(file_type)
            c4.markdown(f"{file_size_mb:.1f}")
            c5.markdown(upload_date)
            c6.markdown(status)

            is_viewing = st.session_state.get("processing_selected_doc") == str(file_path)
            btn_label = "Hide" if is_viewing else "View"
            if c7.button(btn_label, key=f"view_doc_{file_path.name}", use_container_width=True):
                if is_viewing:
                    st.session_state["processing_selected_doc"] = None
                else:
                    st.session_state["processing_selected_doc"] = str(file_path)
                st.rerun()

            if st.session_state.get("processing_selected_doc") == str(file_path):
                st.markdown(f"##### Preview: {file_path.name}")
                display_processing_file_preview(file_path)

            st.markdown("---")

    else:
        st.info("No documents found.")

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: OCR VIEWER
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🔍 OCR Viewer":

    st.markdown("### 🔍 OCR Output Viewer")
    st.markdown("Browse previously processed OCR results with confidence scoring.")

    ocr_dir = Path(__file__).resolve().parent / "output" / "ocr"
    if ocr_dir.exists():
        ocr_files = sorted(ocr_dir.glob("*.json"))
        if ocr_files:
            file_names = [p.name for p in ocr_files]
            search_ocr = st.text_input(
                "Search OCR output",
                placeholder="Type to filter by filename...",
                label_visibility="collapsed",
            )
            filtered_names = [n for n in file_names if search_ocr.lower() in n.lower()] if search_ocr else file_names
            if filtered_names:
                sel_name = st.selectbox("Select OCR output", filtered_names, label_visibility="collapsed")
                sel = ocr_dir / sel_name
                if sel:
                    data = load_json_file(sel)
                    if isinstance(data, dict):
                        display_ocr_result(data)
                    else:
                        st.json(data)
            else:
                st.info(f"No OCR files match: `{search_ocr}`")
        else:
            st.info("No OCR output files found.")
    else:
        st.info("OCR output directory not found.")


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: EXTRACTION VIEWER
# ═══════════════════════════════════════════════════════════════════════════
elif page == "📊 Extraction Viewer":

    repo_items = load_extraction_repository_items()
    if not repo_items:
        st.info("No extraction files found.")
    else:
        search_query = st.text_input("", placeholder="Search invoice ID, vendor, or source file...", label_visibility="collapsed")
        f1, f2 = st.columns([1, 1])
        with f1:
            all_statuses = sorted({item["status"] for item in repo_items if item["status"]})
            status_filter = st.selectbox("Status", ["All Statuses"] + all_statuses)
        with f2:
            all_vendors = sorted({item["vendor"] for item in repo_items if item["vendor"] and item["vendor"] != "-"})
            vendor_filter = st.selectbox("Vendor", ["All Vendors"] + all_vendors)

        filtered_items = repo_items

        selected_file = st.session_state.get("extraction_selected_file")
        if selected_file and all(item["source_file"] != selected_file for item in filtered_items):
            st.session_state["extraction_selected_file"] = None

        if search_query:
            q = search_query.lower().strip()
            filtered_items = [
                item for item in filtered_items
                if q in item["invoice_id"].lower()
                or q in item["vendor"].lower()
                or q in item["source_file"].lower()
            ]
        if status_filter != "All Statuses":
            filtered_items = [item for item in filtered_items if item["status"] == status_filter]
        if vendor_filter != "All Vendors":
            filtered_items = [item for item in filtered_items if item["vendor"] == vendor_filter]

        st.markdown(f"#### Invoices ({len(filtered_items)})")
        h1, h2, h3, h4, h5, h6 = st.columns([2.3, 3.0, 1.5, 1.7, 2.0, 1.5])
        h1.markdown("**Invoice ID**")
        h2.markdown("**Vendor**")
        h3.markdown("**Date**")
        h4.markdown("**Total**")
        h5.markdown("**Status**")
        h6.markdown("**Last Updated**")

        for item in filtered_items:
            r1, r2, r3, r4, r5, r6 = st.columns([2.3, 3.0, 1.5, 1.7, 2.0, 1.5])
            if r1.button(item["invoice_id"], key=f"ext_row_{item['source_file']}", use_container_width=True):
                current = st.session_state.get("extraction_selected_file")
                st.session_state["extraction_selected_file"] = None if current == item["source_file"] else item["source_file"]

            r2.markdown(item["vendor"])
            r3.markdown(item["date"])
            r4.markdown(item["total"])
            r5.markdown(item["status"])
            r6.markdown(item["last_updated"])

            if st.session_state.get("extraction_selected_file") == item["source_file"]:
                st.markdown(f"##### Details: {item['invoice_id']} ({item['source_file']})")
                detail_tab, raw_tab = st.tabs(["📊 Structured View", "📝 Raw JSON"])
                with detail_tab:
                    display_extraction_result(item["data"], item["data"].get("document_type", "Unknown"))
                with raw_tab:
                    st.json(item["data"])

            st.markdown("---")


# ═══════════════════════════════════════════════════════════════════════════
# PAGE: REPORT FORMAT
# ═══════════════════════════════════════════════════════════════════════════
elif page == "📋 Report Format":

    st.markdown("### 📋 Extraction Report — Spreadsheet View")
    st.markdown(
        "All extracted documents are mapped to the standard report format below. "
        "You can **review, edit, and export** the data."
    )

    # ── Load all rows first so we can derive date bounds ──────────────────
    df = load_all_extraction_rows()

    # ── Date filter ───────────────────────────────────────────────────────
    def _parse_inv_date(val):
        if not val:
            return None
        for fmt in ["%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d %b %Y", "%d %B %Y", "%B %d, %Y", "%Y/%m/%d"]:
            try:
                return datetime.strptime(str(val).strip(), fmt).date()
            except Exception:
                pass
        return None

    _date_mode = st.radio(
        "Filter by date",
        ["Document Upload Date", "Invoice / PO Date"],
        horizontal=True,
        key="report_date_mode",
    )
    _dfc1, _dfc2 = st.columns(2)
    with _dfc1:
        _date_from = st.date_input(
            "From date",
            value=None,
            key="report_date_from",
            help="Leave empty for no lower bound",
        )
    with _dfc2:
        _date_to = st.date_input(
            "To date",
            value=None,
            key="report_date_to",
            help="Leave empty for no upper bound",
        )

    _apply_date_filter = _date_from is not None or _date_to is not None

    if _date_mode == "Document Upload Date":
        if _apply_date_filter:
            df = df[
                df["Upload Date"].apply(
                    lambda d: d is not None
                    and (_date_from is None or d >= _date_from)
                    and (_date_to is None or d <= _date_to)
                )
            ].reset_index(drop=True)
    else:  # Invoice / PO Date
        df["_inv_date_parsed"] = df["Invoice Date"].apply(_parse_inv_date)
        if _apply_date_filter:
            _include_no_date = st.checkbox(
                "Include documents with no invoice/PO date",
                value=True,
                key="report_include_no_date",
            )

            def _inv_date_filter(row):
                d = row["_inv_date_parsed"]
                if d is None:
                    return _include_no_date
                return (_date_from is None or d >= _date_from) and (_date_to is None or d <= _date_to)

            df = df[df.apply(_inv_date_filter, axis=1)].reset_index(drop=True)
        if "_inv_date_parsed" in df.columns:
            df = df.drop(columns=["_inv_date_parsed"])

    # Re-number after date filter
    for _ri, _rrow in enumerate(df.index, start=1):
        df.at[_rrow, "No"] = _ri

    if df.empty:
        st.info("No extraction files available.")
    else:
        # ── Summary metrics ──────────────────────────────────────────
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("📄 Total Documents", len(df))
        types_counts = df["Types (Inv/CN)"].value_counts()
        col2.metric("🧾 Invoices", types_counts.get("Inv", 0))
        col3.metric("⚡ Utility Bills", types_counts.get("Utility", 0))
        col4.metric("🏠 Rental", types_counts.get("Rental", 0))

        st.divider()

        # ── Run matching & enrich DataFrame ──────────────────────────
        matches = match_utility_to_rental(df)
        df["Matched To"] = ""
        df["Match Confidence"] = ""
        df["Matched On"] = ""
        for m in matches:
            u_idx, r_idx = m["utility_idx"], m["rental_idx"]
            conf = "High" if m["score"] >= 6 else ("Medium" if m["score"] >= 4 else "Low")
            keys_str = ", ".join(m["matched_on"])
            # Utility row → points to rental
            r_inv = df.loc[r_idx, "Invoice No"] if r_idx in df.index else ""
            r_name = df.loc[r_idx, "Company Name"] if r_idx in df.index else ""
            df.at[u_idx, "Matched To"] = f"Rental #{df.loc[r_idx, 'No']} — {r_inv} ({r_name})"
            df.at[u_idx, "Match Confidence"] = conf
            df.at[u_idx, "Matched On"] = keys_str
            # Rental row → points to utility
            u_inv = df.loc[u_idx, "Invoice No"] if u_idx in df.index else ""
            u_name = df.loc[u_idx, "Company Name"] if u_idx in df.index else ""
            df.at[r_idx, "Matched To"] = f"Utility #{df.loc[u_idx, 'No']} — {u_inv} ({u_name})"
            df.at[r_idx, "Match Confidence"] = conf
            df.at[r_idx, "Matched On"] = keys_str
            # Copy utility-specific fields into the rental row
            for fld in ["Electricity Amount", "Kwh Reading Before", "Kwh Reading After", "Current Reading / Total Units"]:
                u_val = df.loc[u_idx, fld] if u_idx in df.index else ""
                if u_val and (not df.at[r_idx, fld]):
                    df.at[r_idx, fld] = u_val

        # ── Tabs: Spreadsheet  |  Matched Pairs ──────────────────────
        tab_sheet, tab_matched = st.tabs(["📊 Spreadsheet View", "🔗 Utility ↔ Rental Matching"])

        # ── TAB 1: DOCUMENT REVIEW ────────────────────────────────
        with tab_sheet:
            # Type style lookup
            _type_style = {
                "Inv": ("🧾", "#d1fae5", "#065f46"),
                "Utility": ("⚡", "#ecfdf5", "#047857"),
                "Rental": ("🏠", "#d1fae5", "#022c22"),
                "Hotel": ("🏨", "#fde8e7", "#EE2D25"),
                "Travel": ("✈️", "#d1fae5", "#059669"),
                "SOA": ("📑", "#ecfdf5", "#047857"),
                "CN": ("📌", "#fde8e7", "#d68000"),
                "SRKK-Vendor": ("🏢", "#eff6ff", "#1d4ed8"),
                "SRKK-PO": ("📋", "#f0f4ff", "#4338ca"),
                "SRKK-MS Billing": ("☁️", "#e0f2fe", "#0369a1"),
            }

            # ── Toolbar ──────────────────────────────────────────────
            tb1, tb2, tb3, tb4 = st.columns([2, 2, 2, 1])
            with tb1:
                type_filter = st.multiselect(
                    "Document Type",
                    options=sorted(df["Types (Inv/CN)"].dropna().unique()),
                    default=[],
                    placeholder="All types",
                )
            with tb2:
                _company_opts = sorted([
                    c for c in df["Company Name"].dropna().unique()
                    if c and str(c).strip()
                ])
                company_filter = st.multiselect(
                    "Company Name",
                    options=_company_opts,
                    default=[],
                    placeholder="All companies",
                )
            with tb3:
                status_filter = st.multiselect(
                    "Review Status",
                    options=["Pending", "Verified", "Rejected"],
                    default=["Pending", "Verified", "Rejected"],
                )
            with tb4:
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Reset All", use_container_width=True):
                    st.session_state["doc_status"] = {}
                    st.rerun()

            # Apply filters (empty selection = show all)
            filtered = df.copy()
            if type_filter:
                filtered = filtered[filtered["Types (Inv/CN)"].isin(type_filter)]
            if company_filter:
                filtered = filtered[filtered["Company Name"].isin(company_filter)]
            # Apply status filter
            status_map_lower = {"Pending": "pending", "Verified": "verified", "Rejected": "rejected"}
            allowed_statuses = {status_map_lower[s] for s in status_filter}
            filtered = filtered[
                filtered["No"].apply(
                    lambda n: st.session_state["doc_status"].get(int(n), "pending") in allowed_statuses
                )
            ]

            # ── Counts bar ───────────────────────────────────────────
            n_verified = sum(1 for v in st.session_state["doc_status"].values() if v == "verified")
            n_rejected = sum(1 for v in st.session_state["doc_status"].values() if v == "rejected")
            n_pending = len(df) - n_verified - n_rejected
            cb1, cb2, cb3, cb4 = st.columns(4)
            cb1.markdown(f"**{len(filtered)}** of **{len(df)}** shown")
            cb2.markdown(f"✅ **{n_verified}** verified")
            cb3.markdown(f"❌ **{n_rejected}** rejected")
            cb4.markdown(f"⏳ **{n_pending}** pending")
            st.markdown("---")

            # ── Document cards ───────────────────────────────────────
            for _, row in filtered.iterrows():
                row_no = int(row["No"])
                doc_type = str(row.get("Types (Inv/CN)", "Inv"))
                icon, type_bg, accent = _type_style.get(doc_type, ("📄", "#f7fafc", "#4a5568"))
                company = row.get("Company Name", "") or "—"
                inv_no = row.get("Invoice No", "") or "—"
                inv_date = row.get("Invoice Date", "") or "—"
                total = row.get("Total Amount (incl. Tax)", "") or "—"
                matched_to = row.get("Matched To", "")
                status = st.session_state["doc_status"].get(row_no, "pending")
                status_class = f"doc-status-{status}"

                # ── Card header (HTML) ────────────────────────────────
                match_html = ""
                if matched_to:
                    conf = row.get("Match Confidence", "")
                    conf_colors = {"High": "#047857", "Medium": "#d68000", "Low": "#EE2D25"}
                    mc = conf_colors.get(conf, "#5a8a8f")
                    match_html = f'<span class="doc-card-match" style="background:{mc};color:white;">🔗 {conf}</span>'

                status_icons = {"verified": "✅", "rejected": "❌", "pending": "⏳"}
                status_labels = {"verified": "Verified", "rejected": "Rejected", "pending": "Pending"}
                s_icon = status_icons.get(status, "")
                s_label = status_labels.get(status, "")

                card_html = (
                    f'<div class="doc-card {status_class}">'
                    f'<div class="doc-card-header">'
                    f'<span class="doc-card-num">{row_no}</span>'
                    f'<span class="doc-card-type" style="background:{type_bg};color:{accent};">{icon} {doc_type}</span>'
                    f'<span class="doc-card-company">{company}</span>'
                    f'<span class="doc-card-detail">{inv_no}</span>'
                    f'<span class="doc-card-detail">{inv_date}</span>'
                    f'<span class="doc-card-amount">{total}</span>'
                    f'{match_html}'
                    f'<span class="doc-card-status">{s_icon} {s_label}</span>'
                    f'</div></div>'
                )
                row_left, row_right = st.columns([12, 2])
                with row_left:
                    st.markdown(card_html, unsafe_allow_html=True)
                with row_right:
                    if st.button("View", key=f"view_detail_{row_no}", use_container_width=True):
                        current_open = st.session_state.get("report_detail_row")
                        st.session_state["report_detail_row"] = None if current_open == row_no else row_no

                if st.session_state.get("report_detail_row") == row_no:
                    # Action buttons row
                    ac1, ac2, ac3 = st.columns([1, 1, 4])
                    with ac1:
                        if st.button("✅ Verify", key=f"verify_{row_no}", use_container_width=True,
                                     type="primary" if status != "verified" else "secondary"):
                            st.session_state["doc_status"][row_no] = "verified"
                            st.rerun()
                    with ac2:
                        if st.button("❌ Reject", key=f"reject_{row_no}", use_container_width=True,
                                     type="primary" if status != "rejected" else "secondary"):
                            st.session_state["doc_status"][row_no] = "rejected"
                            st.rerun()
                    with ac3:
                        if status != "pending":
                            if st.button("↩️ Reset to Pending", key=f"reset_{row_no}"):
                                st.session_state["doc_status"][row_no] = "pending"
                                st.rerun()

                    st.markdown("**Quick Reference (table order)**")
                    ordered_cols = [c for c in REPORT_COLUMNS if c in row.index]
                    detail_row = {c: (row.get(c, "") if row.get(c, "") not in (None, "") else "—") for c in ordered_cols}
                    st.dataframe(
                        pd.DataFrame([detail_row]),
                        use_container_width=True,
                        hide_index=True,
                    )

                    if matched_to:
                        st.caption(
                            f"Matched To: {matched_to} | Confidence: {row.get('Match Confidence', '')} | Matched On: {row.get('Matched On', '')}"
                        )

                    src = row.get("_source_file", "")
                    if src:
                        st.caption(f"Source: {src}")

                    if src:
                        pdf_match = find_source_pdf_for_extraction(src)
                        btn_label = "📄 View Original PDF"
                        if st.button(btn_label, key=f"view_original_pdf_{row_no}"):
                            current_preview = st.session_state.get("report_preview_source")
                            st.session_state["report_preview_source"] = None if current_preview == src else src

                        if st.session_state.get("report_preview_source") == src:
                            if pdf_match and pdf_match.exists():
                                st.markdown(f"##### Original PDF: {pdf_match.name}")
                                display_processing_file_preview(pdf_match)
                            else:
                                st.info("Matching PDF not found in src/docs/uploads.")

            st.markdown("---")

            # ── Table View ────────────────────────────────────────────
            st.markdown("#### 📊 Table View")
            display_cols = [c for c in REPORT_COLUMNS if c in filtered.columns]
            for extra in ["Matched To", "Match Confidence", "Matched On"]:
                if extra in filtered.columns:
                    display_cols.append(extra)
            table_df = filtered[display_cols].copy()
            if "No" in table_df.columns:
                table_df.insert(1, "Status", table_df["No"].apply(
                    lambda n: st.session_state["doc_status"].get(int(n), "pending").capitalize()
                ))
            else:
                table_df.insert(0, "Status", "Pending")

            if table_df.empty:
                st.info("No records match the selected filters.")

            # Styled HTML table with larger font in scrollable container
            st.markdown("""<style>
            .report-table-wrap {
                max-height: 500px;
                overflow-y: auto;
                overflow-x: auto;
                border: 2px solid #b0c4c8;
                border-radius: 8px;
            }
            .report-table-wrap table {
                width: 100%;
                border-collapse: collapse;
                font-size: 0.9rem;
            }
            .report-table-wrap thead { position: sticky; top: 0; z-index: 1; }
            .report-table-wrap th {
                background-color: #e4e8ec;
                font-weight: 700;
                padding: 10px 12px;
                border: 1px solid #b0c4c8;
                text-align: left;
                white-space: nowrap;
                font-size: 0.9rem;
            }
            .report-table-wrap td {
                padding: 8px 12px;
                border: 1px solid #c8d6da;
                font-size: 0.9rem;
            }
            .report-table-wrap tr:hover { background-color: #f0f6f7; }
            </style>""", unsafe_allow_html=True)

            table_html = table_df.to_html(index=False, escape=True, border=0)
            st.markdown(
                f'<div class="report-table-wrap">{table_html}</div>',
                unsafe_allow_html=True,
            )

            st.markdown("---")

            # ── Export ────────────────────────────────────────────────
            # Add status column to export
            display_cols = [c for c in REPORT_COLUMNS if c in filtered.columns]
            for extra in ["Matched To", "Match Confidence", "Matched On"]:
                if extra in filtered.columns:
                    display_cols.append(extra)
            export_df = filtered[display_cols].copy()
            if "No" in export_df.columns:
                export_df.insert(1, "Status", export_df["No"].apply(
                    lambda n: st.session_state["doc_status"].get(int(n), "pending").capitalize()
                ))
            else:
                export_df.insert(0, "Status", "Pending")

            st.markdown("#### 📥 Export Report")
            st.caption("Only verified & pending documents shown. Rejected documents are excluded from export.")

            # Filter out rejected for export
            export_clean = export_df[export_df["Status"] != "Rejected"]

            exp_col1, exp_col2, exp_col3 = st.columns(3)

            with exp_col1:
                csv_data = export_clean.to_csv(index=False).encode("utf-8-sig")
                st.download_button(
                    "⬇️ Download CSV",
                    data=csv_data,
                    file_name=f"extraction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                )

            with exp_col2:
                import io
                _excel_buf = io.BytesIO()
                export_clean.to_excel(_excel_buf, index=False, engine="openpyxl")
                _excel_bytes = _excel_buf.getvalue()
                st.download_button(
                    "⬇️ Download Excel",
                    data=_excel_bytes,
                    file_name=f"extraction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            with exp_col3:
                json_data = export_clean.to_json(orient="records", force_ascii=False, indent=2)
                st.download_button(
                    "⬇️ Download JSON",
                    data=json_data.encode("utf-8"),
                    file_name=f"extraction_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json",
                )

        # ── TAB 2: UTILITY ↔ RENTAL MATCHING ─────────────────────────
        with tab_matched:
            st.markdown(
                "Automatically matches **utility / electricity bills** to their corresponding "
                "**rental invoices** using shared identifiers: Lease ID, Lot No, Vendor Name, TIN No, Account No."
            )

            if not matches:
                st.info(
                    "No utility ↔ rental matches found. This requires utility bills and rental invoices "
                    "with shared Lease ID, Lot No, Vendor Name, or TIN No."
                )
            else:
                st.success(f"🔗 Found **{len(matches)}** utility ↔ rental match(es)")

                for i, m in enumerate(matches, 1):
                    u_row = df.loc[m["utility_idx"]]
                    r_row = df.loc[m["rental_idx"]]
                    match_tags = ", ".join(m["matched_on"])
                    confidence = "🟢 High" if m["score"] >= 6 else ("🟡 Medium" if m["score"] >= 4 else "🟠 Low")

                    with st.container():
                        st.markdown(f"---")
                        st.markdown(f"#### Match {i} — {confidence} confidence")
                        st.caption(f"Matched on: {match_tags}")

                        pair_col1, pair_col2 = st.columns(2)

                        with pair_col1:
                            st.markdown(
                                '<div style="background:#ecfdf5; border-left:4px solid #047857; '
                                'padding:0.8rem 1rem; border-radius:0 8px 8px 0; margin-bottom:0.6rem;">'
                                '<strong>⚡ Utility / Electricity Bill</strong></div>',
                                unsafe_allow_html=True,
                            )
                            st.markdown(f"**Vendor:** {u_row.get('Company Name', '')}")
                            st.markdown(f"**Invoice No:** {u_row.get('Invoice No', '')}")
                            st.markdown(f"**Invoice Date:** {u_row.get('Invoice Date', '')}")
                            st.markdown(f"**Lease ID:** {u_row.get('Lease ID', '')}")
                            st.markdown(f"**Lot No:** {u_row.get('Lot No', '')}")
                            st.markdown(f"**TIN No:** {u_row.get('TIN No', '')}")
                            st.markdown(f"**Electricity Amount:** {u_row.get('Electricity Amount', '')}")
                            st.markdown(f"**Total (incl. Tax):** {u_row.get('Total Amount (incl. Tax)', '')}")
                            kwh_b = u_row.get("Kwh Reading Before", "")
                            kwh_a = u_row.get("Kwh Reading After", "")
                            total_u = u_row.get("Current Reading / Total Units", "")
                            if kwh_b or kwh_a:
                                st.markdown(f"**kWh Before → After:** {kwh_b} → {kwh_a}")
                            if total_u:
                                st.markdown(f"**Total Units:** {total_u}")

                        with pair_col2:
                            st.markdown(
                                '<div style="background:#d1fae5; border-left:4px solid #022c22; '
                                'padding:0.8rem 1rem; border-radius:0 8px 8px 0; margin-bottom:0.6rem;">'
                                '<strong>🏠 Rental / Lease Invoice</strong></div>',
                                unsafe_allow_html=True,
                            )
                            st.markdown(f"**Vendor:** {r_row.get('Company Name', '')}")
                            st.markdown(f"**Invoice No:** {r_row.get('Invoice No', '')}")
                            st.markdown(f"**Invoice Date:** {r_row.get('Invoice Date', '')}")
                            st.markdown(f"**Lease ID:** {r_row.get('Lease ID', '')}")
                            st.markdown(f"**Lot No:** {r_row.get('Lot No', '')}")
                            st.markdown(f"**TIN No:** {r_row.get('TIN No', '')}")
                            st.markdown(f"**Total (incl. Tax):** {r_row.get('Total Amount (incl. Tax)', '')}")
                            st.markdown(f"**Description:** {r_row.get('Description', '')}")

                        # Side-by-side comparison table
                        compare_fields = [
                            "Company Name", "TIN No", "Invoice No", "Invoice Date", "Lease ID",
                            "Lot No", "Account No", "Unit No", "Location",
                            "Total Amount (incl. Tax)",
                        ]
                        compare_rows = []
                        for fld in compare_fields:
                            u_val = str(u_row.get(fld, ""))
                            r_val = str(r_row.get(fld, ""))
                            match_icon = "✅" if (u_val and r_val and u_val.strip() == r_val.strip()) else (
                                "🔶" if (u_val and r_val) else "—"
                            )
                            display_fld = fld.replace("\n", " ")
                            compare_rows.append({
                                "Field": display_fld,
                                "⚡ Utility": u_val,
                                "🏠 Rental": r_val,
                                "Match": match_icon,
                            })
                        with st.expander("📊 Field-by-Field Comparison", expanded=False):
                            st.dataframe(
                                pd.DataFrame(compare_rows),
                                use_container_width=True,
                                hide_index=True,
                            )

            # Show unmatched utility bills
            matched_util_idxs = {m["utility_idx"] for m in matches}
            unmatched_utils = df[
                (df["Types (Inv/CN)"] == "Utility") & (~df.index.isin(matched_util_idxs))
            ]
            if not unmatched_utils.empty:
                st.markdown("---")
                st.markdown("#### ⚠️ Unmatched Utility Bills")
                st.caption("These utility bills could not be matched to any rental invoice.")
                st.dataframe(
                    unmatched_utils[["No", "Company Name", "Invoice No", "Invoice Date",
                                     "Lease ID", "Lot No", "TIN No",
                                     "Total Amount (incl. Tax)"]].reset_index(drop=True),
                    use_container_width=True,
                    hide_index=True,
                )

# ═══════════════════════════════════════════════════════════════════════════
# PAGE: RECONCILIATION
# ═══════════════════════════════════════════════════════════════════════════
elif page == "🔄 Reconciliation":
    st.markdown("### 🔄 Reconciliation")
    st.markdown(
        "Select the **primary document** to reconcile. "
        "If it is a Microsoft Billing document, a second upload area will appear for the Purchase Order."
    )

    # ── Collect available uploaded files ────────────────────────────────────
    _recon_uploads_dir = Path(__file__).resolve().parent / "docs" / "uploads"
    _recon_extraction_dir = Path(__file__).resolve().parent / "output" / "extraction"

    _recon_pdf_files: list[Path] = []
    if _recon_uploads_dir.exists():
        _recon_pdf_files = sorted(
            [p for p in _recon_uploads_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"],
            key=lambda p: p.stat().st_mtime, reverse=True,
        )

    if not _recon_pdf_files:
        st.info(
            "No uploaded documents found. "
            "Go to **Document Processing** to upload and process your documents first."
        )
    else:
        _recon_file_names = [p.name for p in _recon_pdf_files]

        # ── File selector (single primary document) ──────────────────────────
        _recon_selected_name = st.selectbox(
            "Select primary document to reconcile",
            options=["— select —"] + _recon_file_names,
            index=0,
            key="recon_file_selector",
        )

        if _recon_selected_name == "— select —":
            st.warning("Select a document to run reconciliation.")
        else:
            _recon_selected_paths = [_recon_uploads_dir / _recon_selected_name]
            _recon_selected_names = [_recon_selected_name]

            # ── File preview cards ───────────────────────────────────────────
            st.markdown("#### Selected Documents")
            _prev_cols = st.columns(len(_recon_selected_paths))
            for _ci, (_cp, _col) in enumerate(zip(_recon_selected_paths, _prev_cols)):
                with _col:
                    # Check if extraction output exists
                    _ext_candidates = sorted(_recon_extraction_dir.glob(f"{_cp.stem}*.json")) if _recon_extraction_dir.exists() else []
                    _has_extraction = bool(_ext_candidates)
                    _status_icon = "✅" if _has_extraction else "⏳"
                    _status_text = "Extracted" if _has_extraction else "Not yet processed"
                    st.markdown(
                        f'<div style="background:#ecfdf5;border:1px solid #6ee7b7;border-radius:10px;'
                        f'padding:0.8rem 1rem;margin-bottom:0.5rem;">'
                        f'<div style="font-weight:700;color:#065f46;font-size:0.9rem;margin-bottom:0.3rem;">'
                        f'📄 {_cp.name}</div>'
                        f'<div style="font-size:0.78rem;color:#047857;">{_status_icon} {_status_text}</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )
                    _prev_key = f"recon_prev_{_cp.name}"
                    if st.button(f"👁 Preview", key=_prev_key, use_container_width=True):
                        st.session_state[f"recon_show_{_cp.name}"] = not st.session_state.get(f"recon_show_{_cp.name}", False)
                    if st.session_state.get(f"recon_show_{_cp.name}", False):
                        display_processing_file_preview(_cp)

            # ── Check all selected files have extraction output ──────────────
            _recon_extractions: list[dict] = []
            _missing_extractions: list[str] = []
            for _rp in _recon_selected_paths:
                _ext_files = sorted(_recon_extraction_dir.glob(f"{_rp.stem}*.json")) if _recon_extraction_dir.exists() else []
                if _ext_files:
                    try:
                        _d = json.loads(_ext_files[0].read_text(encoding="utf-8"))
                        _d["_source_file"] = _rp.name
                        _recon_extractions.append(_d)
                    except Exception:
                        _missing_extractions.append(_rp.name)
                else:
                    _missing_extractions.append(_rp.name)

            if _missing_extractions:
                st.warning(
                    f"The following files have not been processed yet — run the Document Processing "
                    f"pipeline on them first: **{', '.join(_missing_extractions)}**"
                )

            # ── Microsoft Billing Excel selector ────────────────────────────
            _has_ms_billing = any(
                "microsoft" in (d.get("document_type") or "").lower()
                or "cloud billing" in (d.get("document_type") or "").lower()
                or (isinstance(d.get("invoice"), dict) and "billing_summary" in (d.get("invoice") or {}))
                for d in _recon_extractions
            )

            _ms_excel_path: Path | None = None
            if _has_ms_billing:
                st.markdown("---")
                st.markdown("#### ☁️ Microsoft Billing — Excel Reference")
                st.caption(
                    "A Microsoft Billing document was detected. Provide the corresponding Excel "
                    "breakdown to cross-reference line items."
                )

                _ms_excel_dir = Path(__file__).resolve().parent / "docs" / "uploads" / "microsoft_billing"
                _ms_excel_dir.mkdir(parents=True, exist_ok=True)
                _existing_excels = sorted(
                    [f for f in _ms_excel_dir.iterdir() if f.suffix.lower() in (".xlsx", ".xls", ".csv")],
                    key=lambda f: f.stat().st_mtime, reverse=True,
                )

                _ms_excel_mode = st.radio(
                    "File source",
                    ["Upload new file", "Use previously uploaded file"],
                    horizontal=True,
                    key="ms_excel_mode",
                )

                if _ms_excel_mode == "Upload new file":
                    _ms_excel_upload = st.file_uploader(
                        "Upload Microsoft Billing file (.xlsx / .xls / .csv)",
                        type=["xlsx", "xls", "csv"],
                        key="ms_excel_uploader",
                    )
                    if _ms_excel_upload is not None:
                        _save_path = _ms_excel_dir / _ms_excel_upload.name
                        _save_path.write_bytes(_ms_excel_upload.read())
                        _ms_excel_path = _save_path
                        st.success(f"✅ Uploaded: **{_ms_excel_upload.name}**")
                else:
                    if _existing_excels:
                        _excel_names = [f.name for f in _existing_excels]
                        _selected_excel_name = st.selectbox(
                            "Select file",
                            options=_excel_names,
                            key="ms_excel_selector",
                        )
                        _ms_excel_path = _ms_excel_dir / _selected_excel_name
                        st.caption(f"Selected: `{_ms_excel_path.name}`")
                    else:
                        st.info("No previously uploaded files found. Upload one above.")

                if _ms_excel_path and _ms_excel_path.exists():
                    try:
                        if _ms_excel_path.suffix.lower() == ".csv":
                            _ms_preview_df = pd.read_csv(_ms_excel_path, nrows=5)
                        else:
                            _ms_preview_df = pd.read_excel(_ms_excel_path, nrows=5)
                        with st.expander("Preview file (first 5 rows)", expanded=False):
                            st.dataframe(_ms_preview_df, use_container_width=True, hide_index=True)
                    except Exception as _ex:
                        st.warning(f"Could not preview file: {_ex}")

                st.session_state["recon_ms_excel_path"] = str(_ms_excel_path) if _ms_excel_path else None

                # ── Purchase Order second upload area ────────────────────────
                st.markdown("---")
                st.markdown("#### 📋 Purchase Order Document")
                st.caption(
                    "Upload or select the Purchase Order PDF that corresponds to this billing. "
                    "The PO will be used for the second matching stage."
                )

                _po_uploads_dir = Path(__file__).resolve().parent / "docs" / "uploads"
                _existing_pos = sorted(
                    [p for p in _po_uploads_dir.iterdir()
                     if p.is_file() and p.suffix.lower() == ".pdf"
                     and p.name != _recon_selected_name],
                    key=lambda p: p.stat().st_mtime, reverse=True,
                )
                _po_file_names = [p.name for p in _existing_pos]

                _po_upload_mode = st.radio(
                    "PO file source",
                    ["Upload new PO PDF", "Select from uploaded documents"],
                    horizontal=True,
                    key="po_upload_mode",
                )

                _po_pdf_path: Path | None = None
                if _po_upload_mode == "Upload new PO PDF":
                    _po_pdf_upload = st.file_uploader(
                        "Upload Purchase Order PDF",
                        type=["pdf"],
                        key="po_pdf_uploader",
                    )
                    if _po_pdf_upload is not None:
                        _po_save_path = _po_uploads_dir / _po_pdf_upload.name
                        _po_save_path.write_bytes(_po_pdf_upload.read())
                        _po_pdf_path = _po_save_path
                        st.success(f"✅ Uploaded PO: **{_po_pdf_upload.name}**")
                else:
                    if _po_file_names:
                        _selected_po_name = st.selectbox(
                            "Select PO document",
                            options=_po_file_names,
                            key="po_pdf_selector",
                        )
                        _po_pdf_path = _po_uploads_dir / _selected_po_name
                        st.caption(f"Selected: `{_po_pdf_path.name}`")
                    else:
                        st.info("No other uploaded PDFs found. Upload the PO above.")

                # Resolve PO extraction JSON
                _po_extraction_dir = Path(__file__).resolve().parent / "output" / "extraction"
                _po_extraction_data: dict | None = None
                if _po_pdf_path and _po_pdf_path.exists():
                    _po_ext_candidates = sorted(_po_extraction_dir.glob(f"{_po_pdf_path.stem}*.json")) if _po_extraction_dir.exists() else []
                    if _po_ext_candidates:
                        try:
                            _po_extraction_data = json.loads(_po_ext_candidates[0].read_text(encoding="utf-8"))
                            _po_extraction_data["_source_file"] = _po_pdf_path.name
                            st.success(f"✅ PO extraction found: `{_po_ext_candidates[0].name}`")
                        except Exception as _pe:
                            st.warning(f"Could not load PO extraction: {_pe}")
                    else:
                        st.warning(
                            f"No extraction output found for `{_po_pdf_path.name}`. "
                            "Process this document in **Document Processing** first."
                        )
                st.session_state["recon_po_extraction"] = _po_extraction_data

                # ── Match mode selector ──────────────────────────────────────
                st.markdown("##### Match Mode")
                _ms_match_mode = st.radio(
                    "How should the Microsoft Billing document be reconciled?",
                    ["Match with Excel/CSV only",
                     "Match with Excel/CSV and Purchase Order (3-way)"],
                    index=1,
                    key="ms_match_mode",
                )
                st.session_state["recon_ms_match_mode"] = _ms_match_mode

            if _has_ms_billing and len(_recon_extractions) >= 1:
                st.markdown("---")

                if st.button("▶ Run Reconciliation", type="primary", use_container_width=True, key="recon_run_btn"):
                    st.session_state["recon_ms_ready"] = False
                    st.session_state["recon_po_ready"] = False
                    st.session_state["recon_po_summary_ready"] = False

                    _ms_match_mode = st.session_state.get("recon_ms_match_mode", "Match with Excel/CSV and Purchase Order (3-way)")
                    _ms_excel_path_str = st.session_state.get("recon_ms_excel_path")
                    _ms_excel_file = Path(_ms_excel_path_str) if _ms_excel_path_str else None
                    _po_ext_data = st.session_state.get("recon_po_extraction")

                    # ── Identify MS Billing extraction ───────────────────────
                    _ms_billing_data = next(
                        (d for d in _recon_extractions
                         if isinstance(d.get("invoice"), dict) and "billing_summary" in (d.get("invoice") or {})),
                        None,
                    )

                    # ── Stage 1: MS Billing × Excel ────────────────────────
                    _ms_result = None
                    if _ms_billing_data and _ms_excel_file and _ms_excel_file.exists():
                        try:
                            from core.reconcile.microsoft_billing_reconcile import reconcile as _ms_reconcile
                            st.write(f"  ☁️ Stage 1 — Billing × Excel: running against `{_ms_excel_file.name}`...")
                            _ms_result = _ms_reconcile(_ms_billing_data, _ms_excel_file)
                            st.session_state["recon_ms_result"] = _ms_result
                            st.session_state["recon_ms_ready"] = True
                            _meta = _ms_result["reconcile_meta"]
                            st.write(f"  ✅ Stage 1 done — Green: {_meta['green']}  Yellow: {_meta['yellow']}  Red: {_meta['red']}")
                        except Exception as _ms_err:
                            st.error(f"Stage 1 (Billing × Excel) failed: {_ms_err}")
                    else:
                        if not _ms_excel_file or not _ms_excel_file.exists():
                            st.warning("Stage 1 skipped: no Excel/CSV file provided.")

                    # ── Stage 2: Billing × PO ────────────────────────────────
                    if (
                        _ms_match_mode == "Match with Excel/CSV and Purchase Order (3-way)"
                        and _ms_billing_data
                        and _ms_result
                        and _po_ext_data
                    ):
                        try:
                            from core.reconcile.microsoft_billing_po_reconcile import reconcile_po as _po_reconcile
                            st.write("  📋 Stage 2 — Billing × PO: matching line items against Purchase Order...")
                            _po_result = _po_reconcile(_ms_billing_data, _ms_result, _po_ext_data)
                            st.session_state["recon_po_result"] = _po_result
                            st.session_state["recon_po_ready"] = True
                            _po_meta = _po_result["po_match_meta"]
                            _cust_icon = "✅" if _po_meta["customer_name_match"] else "⚠️"
                            st.write(
                                f"  {_cust_icon} Stage 2 done — Customer match: {_po_meta['customer_name_match']}  "
                                f"Found in Billing: {_po_meta['found_in_billing']}  "
                                f"Not Found in Billing: {_po_meta['not_found_in_billing']}"
                            )
                        except Exception as _po_err:
                            st.error(f"Stage 2 (Billing × PO) failed: {_po_err}")
                    elif _ms_match_mode == "Match with Excel/CSV and Purchase Order (3-way)" and not _po_ext_data:
                        st.warning("Stage 2 skipped: no Purchase Order extraction loaded.")

                    # ── Stage 3: Customer PO Coverage ───────────────────────
                    if (
                        _ms_match_mode == "Match with Excel/CSV and Purchase Order (3-way)"
                        and _ms_result
                    ):
                        try:
                            from core.reconcile.microsoft_billing_po_reconcile import reconcile_po_summary as _po_summary
                            _ext_dir = Path(__file__).resolve().parent / "output" / "extraction"
                            st.write("  🗂️ Stage 3 — PO Coverage: checking which customers have a matching PO...")
                            _po_summary_result = _po_summary(_ms_result, _ext_dir)
                            st.session_state["recon_po_summary_result"] = _po_summary_result
                            st.session_state["recon_po_summary_ready"] = True
                            st.write(
                                f"  ✅ Stage 3 done — "
                                f"Matched: {_po_summary_result['matched_count']}  "
                                f"Unmatched: {_po_summary_result['unmatched_count']}"
                            )
                        except Exception as _s3_err:
                            st.error(f"Stage 3 (PO Coverage) failed: {_s3_err}")

                # ── MS Billing × Excel Results ───────────────────────────────
                if st.session_state.get("recon_ms_ready"):
                    _ms_result = st.session_state["recon_ms_result"]
                    _ms_meta   = _ms_result["reconcile_meta"]
                    _ms_items  = _ms_result["line_items"]

                    st.markdown("### Stage 1: Microsoft Billing × Excel Reconciliation")
                    st.caption(
                        f"Billing # {_ms_meta.get('billing_number')} · "
                        f"Period {(_ms_meta.get('billing_period') or {}).get('from','')} – {(_ms_meta.get('billing_period') or {}).get('to','')} · "
                        f"Reference: `{_ms_meta.get('reference_file')}`"
                    )

                    _mc1, _mc2, _mc3, _mc4, _mc5, _mc6 = st.columns(6)
                    _mc1.metric("Total Line Items",       _ms_meta["total_line_items"])
                    _mc2.metric("🟢 Matched (Exact)",    _ms_meta.get("green", 0))
                    _mc3.metric("🟠 Near Match (±0.01)",  _ms_meta.get("near", 0))
                    _mc4.metric("🔵 Sum Match",           _ms_meta.get("flow2", 0))
                    _mc5.metric("🟡 Ambiguous",           _ms_meta["yellow"])
                    _mc6.metric("🔴 No Match",            _ms_meta["red"])

                    st.markdown("---")

                    _ms_status_filter = st.multiselect(
                        "Filter by status",
                        options=["green", "near", "flow2", "yellow", "red"],
                        default=["green", "near", "flow2", "yellow", "red"],
                        format_func=lambda s: {"green": "🟢 Matched (Exact)", "near": "🟠 Near Match (±0.01)", "flow2": "🔵 Sum Match", "yellow": "🟡 Ambiguous", "red": "🔴 No Match"}[s],
                        key="ms_recon_status_filter",
                    )

                    # ── Grouped display: one expander per billing line item ───
                    _STATUS_LABEL = {
                        "green":  "🟢 Matched (Exact)",
                        "near":   "🟠 Near Match (±0.01)",
                        "flow2":  "🔵 Sum Match",
                        "yellow": "🟡 Ambiguous",
                        "red":    "🔴 No Match",
                    }
                    _STATUS_BG = {
                        "green":  "#f0fdf4",
                        "near":   "#fff7ed",
                        "flow2":  "#eff6ff",
                        "yellow": "#fffbeb",
                        "red":    "#fde8e7",
                    }
                    _STATUS_COLOR = {
                        "green":  "#166534",
                        "near":   "#c2410c",
                        "flow2":  "#1d4ed8",
                        "yellow": "#b45309",
                        "red":    "#c0392b",
                    }

                    # ── Scrollable HTML card list ────────────────────────────
                    def _build_match_table(matches):
                        if not matches:
                            return "<p style='color:#6b7280;font-size:0.85rem;margin:4px 0'>No matching rows found in reference file.</p>"
                        hdr = "".join(f"<th style='padding:4px 8px;border-bottom:1px solid #d1d5db;text-align:left;font-size:0.8rem'>{c}</th>" for c in ["Customer Name", "CSV Amount", "Charge Type", "Date"])
                        rows_html = ""
                        for _mr in matches:
                            amt = _mr.get("Amount") or _mr.get("CSV Amount") or ""
                            row_cells = (
                                f"<td style='padding:3px 8px;font-size:0.8rem'>{_mr.get('Customer Name','')}</td>"
                                f"<td style='padding:3px 8px;font-size:0.8rem'>{amt}</td>"
                                f"<td style='padding:3px 8px;font-size:0.8rem'>{_mr.get('Charge Type','')}</td>"
                                f"<td style='padding:3px 8px;font-size:0.8rem'>{_mr.get('Date','')}</td>"
                            )
                            rows_html += f"<tr>{row_cells}</tr>"
                        return (
                            f"<table style='width:100%;border-collapse:collapse;margin-top:6px'>"
                            f"<thead><tr>{hdr}</tr></thead><tbody>{rows_html}</tbody></table>"
                        )

                    _cards_html = ""
                    for _item in _ms_items:
                        _s = _item["status"]
                        if _s not in _ms_status_filter:
                            continue
                        _bg    = _STATUS_BG[_s]
                        _color = _STATUS_COLOR[_s]
                        _label = _STATUS_LABEL[_s]
                        _prod  = _item["product"].replace("<", "&lt;").replace(">", "&gt;")
                        _amt   = _item["line_amount"]
                        _tbl   = _build_match_table(_item.get("matches", []))
                        # Compute CSV total for display
                        _csv_net = "—"
                        if _item.get("matches"):
                            try:
                                _csv_net = f"{sum(float(str(m.get('Amount', m.get('CSV Amount', 0))).replace(',', '')) for m in _item['matches']):,.2f}"
                            except Exception:
                                _csv_net = "—"
                        _cards_html += f"""
<details style="background:{_bg};border:1px solid {_color}30;border-radius:6px;padding:10px 14px;margin-bottom:8px;cursor:pointer;">
  <summary style="display:flex;justify-content:space-between;align-items:center;list-style:none;outline:none;">
    <span style="font-weight:600;font-size:0.9rem">{_prod}</span>
    <span style="display:flex;gap:10px;align-items:center;">
      <span style="font-size:0.85rem;color:#374151">Billing: <b>{_amt}</b></span>
      <span style="font-size:0.85rem;color:#374151">CSV Total: <b>{_csv_net}</b></span>
      <span style="background:{_bg};color:{_color};padding:2px 10px;border-radius:12px;font-size:0.75rem;font-weight:600;border:1px solid {_color}60">{_label}</span>
    </span>
  </summary>
  <div style="margin-top:8px">{_tbl}</div>
</details>"""

                    st.markdown(
                        f'<div style="max-height:600px;overflow-y:auto;border:1px solid #e5e7eb;'
                        f'border-radius:8px;padding:12px;background:#fafafa">{_cards_html}</div>',
                        unsafe_allow_html=True,
                    )

                    # Download
                    st.download_button(
                        "⬇️ Download Reconciliation JSON",
                        data=json.dumps(_ms_result, indent=2, ensure_ascii=False),
                        file_name=f"ms_billing_recon_{_ms_meta.get('billing_number','result')}.json",
                        mime="application/json",
                    )
                    st.markdown("---")

                # ── PO Match Results ─────────────────────────────────────────
                if st.session_state.get("recon_po_ready"):
                    _po_result = st.session_state.get("recon_po_result", {})
                    _po_meta   = _po_result.get("po_match_meta", {})
                    _po_items  = _po_result.get("line_items", [])

                    st.markdown("---")
                    st.markdown("### Stage 2: Stage 1 Result × Purchase Order Reconciliation")

                    # Customer name match banner
                    if _po_meta.get("customer_name_match"):
                        _matched_names = ", ".join(_po_meta.get("matched_customers", []))
                        st.success(
                            f"✅ **Customer Match Confirmed** — "
                            f"PO delivery recipient **{_po_meta.get('delivery_recipient')}** "
                            f"matched Excel customer(s): **{_matched_names}**"
                        )
                    else:
                        st.warning(
                            f"⚠️ **Customer Name Mismatch** — "
                            f"PO delivery recipient **{_po_meta.get('delivery_recipient')}** "
                            f"did not match any CustomerName in the Excel reconciliation result. "
                            "Line item matching proceeds regardless."
                        )

                    # Summary metrics
                    _poc1, _poc2, _poc3 = st.columns(3)
                    _poc1.metric("Total PO Items",           _po_meta.get("total_po_items", 0))
                    _poc2.metric("✅ Found in Billing",      _po_meta.get("found_in_billing", 0))
                    _poc3.metric("❌ Not Found in Billing",  _po_meta.get("not_found_in_billing", 0))

                    # Table 1: PO Items Found in Billing
                    _found_items = [r for r in _po_items if r["match_status"] == "found_in_billing"]
                    st.markdown(f"#### ✅ PO Item Found in Billing ({len(_found_items)})")
                    if _found_items:
                        _df_found = pd.DataFrame([{
                            "PO Line #":       r["po_line_no"],
                            "PO Description":  r["po_description"],
                            "PO Amount":       r["po_amount"],
                            "Billing Amount":  r.get("billing_amount") or "—",
                            "Order ID":        r.get("order_ids") or "—",
                        } for r in _found_items])
                        st.dataframe(_df_found, use_container_width=True, hide_index=True)
                    else:
                        st.info("No PO line items matched the billing records.")

                    # Table 2: PO Items Not Found in Billing
                    _not_found_items = [r for r in _po_items if r["match_status"] == "not_found_in_billing"]
                    st.markdown(f"#### ❌ PO Items Not Found in Billing ({len(_not_found_items)})")
                    if _not_found_items:
                        _df_nf = pd.DataFrame([{
                            "PO Line #":      r["po_line_no"],
                            "PO Description": r["po_description"],
                            "PO Amount":      r["po_amount"],
                        } for r in _not_found_items])
                        st.dataframe(_df_nf, use_container_width=True, hide_index=True)
                    else:
                        st.info("All PO line items were found in billing.")

                    st.caption(
                        f"PO #{_po_meta.get('po_number')}  |  "
                        f"Date: {_po_meta.get('po_date')}  |  "
                        f"Generated: {_po_meta.get('generated_at')}"
                    )

                # ── Stage 3: Customer PO Coverage ────────────────────────────
                if st.session_state.get("recon_po_summary_ready"):
                    _s3 = st.session_state.get("recon_po_summary_result", {})

                    st.markdown("---")
                    st.markdown("### Stage 3: Customer PO Coverage")

                    _s3c1, _s3c2, _s3c3 = st.columns(3)
                    _s3c1.metric("Total Customers",    _s3.get("total_customers", 0))
                    _s3c2.metric("✅ PO Matched",      _s3.get("matched_count", 0))
                    _s3c3.metric("❌ PO Missing",      _s3.get("unmatched_count", 0))

                    # Matched table
                    _s3_matched = _s3.get("matched", [])
                    st.markdown(f"#### ✅ Customers with Matching PO ({len(_s3_matched)})")
                    if _s3_matched:
                        _df_s3_match = pd.DataFrame([{
                            "Customer Name":    r["customer_name"],
                            "PO Number":        r["po_number"],
                            "Billing Amt":      r["billing_amount"],
                            "PO Amt":           r["po_amount"],
                            "Variance":         r["variance"],
                            "Status":           r["status"],
                        } for r in _s3_matched], index=range(1, len(_s3_matched) + 1))
                        st.dataframe(
                            _df_s3_match,
                            use_container_width=True,
                            column_config={
                                "Billing Amt": st.column_config.NumberColumn("Billing Amt", format="%.2f"),
                                "PO Amt":      st.column_config.NumberColumn("PO Amt",      format="%.2f"),
                                "Variance":    st.column_config.NumberColumn("Variance",    format="%.2f"),
                            },
                        )
                    else:
                        st.info("No customers matched a Purchase Order.")

                    # Unmatched table
                    _s3_unmatched = _s3.get("unmatched", [])
                    st.markdown(f"#### ❌ Customers with No PO ({len(_s3_unmatched)})")
                    if _s3_unmatched:
                        _df_s3_nm = pd.DataFrame([{
                            "Customer Name":  r["customer_name"],
                            "Billing Amt":    r["billing_amount"],
                            "Status":         r["status"],
                        } for r in _s3_unmatched], index=range(1, len(_s3_unmatched) + 1))
                        st.dataframe(
                            _df_s3_nm,
                            use_container_width=True,
                            column_config={
                                "Billing Amt": st.column_config.NumberColumn("Billing Amt", format="%.2f"),
                            },
                        )
                    else:
                        st.info("All customers have a matching Purchase Order.")

                    st.caption(f"Generated: {_s3.get('generated_at')}")