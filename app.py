import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime

# ─── PAGE CONFIG ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Order Studio", page_icon="📦", layout="wide")

# ─── GLOBAL STYLES — Card Minimalism ──────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

/* ════════════════════════════════════════════════════════════
   BASE RESET & FOUNDATION
════════════════════════════════════════════════════════════ */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html { -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility; }

.stApp {
    font-family: 'Inter', system-ui, sans-serif;
    background: #F5F5F7;
    min-height: 100vh;
}
#MainMenu, footer, header { visibility: hidden; display: none; height: 0; overflow: hidden; }
.block-container {
    padding-top: 2rem !important;
    padding-bottom: 5rem !important;
    max-width: 800px !important;
}

/* ════════════════════════════════════════════════════════════
   TOPBAR
════════════════════════════════════════════════════════════ */
.os-topbar {
    display: flex; align-items: center; justify-content: space-between;
    padding: 18px 0 20px;
    margin-bottom: 8px;
    border-bottom: 1px solid rgba(99,102,241,0.12);
}
.os-topbar-brand {
    display: flex; align-items: center; gap: 14px;
}
.os-topbar-logo {
    width: 40px; height: 40px;
    background: linear-gradient(135deg, #6366f1, #818cf8);
    border-radius: 12px;
    display: flex; align-items: center; justify-content: center;
    font-size: 18px;
    box-shadow: 0 4px 12px rgba(99,102,241,0.30);
    flex-shrink: 0;
}
.os-topbar-name {
    font-size: 17px; font-weight: 700;
    color: #1a1a2e; letter-spacing: -0.4px;
    margin-bottom: 3px;
}
.os-topbar-sub {
    font-size: 12px; color: #9ca3af; font-weight: 400;
}
.os-topbar-status {
    display: flex; align-items: center; gap: 7px;
    font-size: 12px; font-weight: 500; color: #6b7280;
    background: #fff;
    border: 1px solid rgba(99,102,241,0.12);
    border-radius: 100px;
    padding: 6px 14px;
    box-shadow: 0 1px 4px rgba(99,102,241,0.06);
}
.os-status-dot {
    width: 7px; height: 7px; border-radius: 50%;
    background: #22c55e;
    display: inline-block;
    animation: sdot 2s ease-in-out infinite;
}
@keyframes sdot {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.4; }
}

/* ════════════════════════════════════════════════════════════
   SECTION LABEL
════════════════════════════════════════════════════════════ */
.os-section {
    display: flex; align-items: center; gap: 12px;
    margin: 36px 0 16px;
    padding-bottom: 14px;
    border-bottom: 1px solid #EBEBEF;
}
.os-section-num {
    display: inline-flex; align-items: center; justify-content: center;
    width: 26px; height: 26px;
    background: #6366f1;
    border-radius: 7px;
    font-size: 11px; font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
    color: #fff;
    flex-shrink: 0;
}
.os-section-title {
    font-size: 16px; font-weight: 700; color: #374151;
    letter-spacing: -0.2px;
}

/* ════════════════════════════════════════════════════════════
   SETTINGS SELECTS
════════════════════════════════════════════════════════════ */
.stSelectbox > div > div {
    border: 1.5px solid #e5e8f4 !important;
    border-radius: 12px !important;
    background: #ffffff !important;
    font-size: 14px !important;
    font-family: 'Inter', sans-serif !important;
    transition: border-color 0.18s, box-shadow 0.18s !important;
    box-shadow: 0 1px 4px rgba(99,102,241,0.05) !important;
}
.stSelectbox > div > div:hover {
    border-color: #a5b4fc !important;
    box-shadow: 0 2px 8px rgba(99,102,241,0.10) !important;
}
.stSelectbox > div > div:focus-within {
    border-color: #6366f1 !important;
    box-shadow: 0 0 0 3px rgba(99,102,241,0.12) !important;
}
.stSelectbox label {
    font-size: 13px !important;
    font-weight: 600 !important;
    color: #6b7280 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    margin-bottom: 8px !important;
}

/* ════════════════════════════════════════════════════════════
   FILE UPLOADER
════════════════════════════════════════════════════════════ */
[data-testid="stFileUploader"] {
    background: #ffffff;
    border: 1.5px solid #E5E7EB !important;
    border-radius: 14px !important;
    padding: 0;
    transition: border-color 0.2s, box-shadow 0.2s;
    margin-top: 6px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: #a5b4fc !important;
    box-shadow: 0 0 0 3px rgba(99,102,241,0.08);
}
[data-testid="stFileUploader"] label {
    font-size: 13px !important;
    font-weight: 600 !important;
    color: #374151 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    text-align: center !important;
    display: block !important;
    padding-top: 8px !important;
}
[data-testid="stFileUploaderDropzone"] {
    border: none !important;
    background: transparent !important;
    padding: 20px !important;
}
/* Chỉ áp dụng layout dọc khi chưa upload (có instructions) */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) {
    flex-direction: column !important;
    align-items: center !important;
    gap: 12px !important;
    padding: 28px 20px !important;
}
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) > div {
    flex-direction: column !important;
    align-items: center !important;
    gap: 8px !important;
    width: 100% !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    font-size: 13px !important;
    color: #9ca3af !important;
    text-align: center !important;
    flex-direction: column !important;
    align-items: center !important;
}
/* Trạng thái đã upload file — compact horizontal */
[data-testid="stFileUploaderFile"] {
    padding: 6px 0 !important;
}
[data-testid="stFileUploaderFile"] button {
    font-size: inherit !important;
    min-width: unset !important;
    padding: 4px !important;
    border-radius: 50% !important;
    border: none !important;
    background: transparent !important;
}
[data-testid="stFileUploaderFile"] button::after {
    content: none !important;
}

/* badge chỉ dùng cho platform tag trong section header */
.badge {
    display: inline-flex; align-items: center; gap: 4px;
    font-size: 11px; font-weight: 600;
    padding: 3px 10px; border-radius: 6px;
    letter-spacing: 0.2px;
    font-family: 'Inter', sans-serif;
    vertical-align: middle;
    margin-left: 4px;
}
.badge-tiktok {
    background: #1a1a2e; color: #fff;
}
.badge-shopee {
    background: rgba(238,77,45,0.10); color: #dd4a1f;
    border: 1px solid rgba(238,77,45,0.20);
}

/* ════════════════════════════════════════════════════════════
   METRIC CARDS
════════════════════════════════════════════════════════════ */
.os-metrics {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
    margin: 20px 0;
}
.os-metric {
    background: #ffffff;
    border: 1px solid rgba(99,102,241,0.10);
    border-radius: 18px;
    padding: 22px 24px;
    position: relative; overflow: hidden;
    box-shadow: 0 2px 12px rgba(99,102,241,0.06);
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.os-metric:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(99,102,241,0.12);
}
.os-metric::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #6366f1, #818cf8, #a5b4fc);
    border-radius: 18px 18px 0 0;
}
.os-metric-icon {
    width: 36px; height: 36px;
    background: rgba(99,102,241,0.10);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 16px; margin-bottom: 14px;
}
.os-metric-label {
    font-size: 12px; font-weight: 600;
    color: #6b7280; text-transform: uppercase;
    letter-spacing: 0.6px; margin-bottom: 8px;
}
.os-metric-value {
    font-size: 42px; font-weight: 800;
    color: #1a1a2e; letter-spacing: -2px;
    line-height: 1; margin-bottom: 6px;
    font-variant-numeric: tabular-nums;
}
.os-metric-note {
    font-size: 13px; color: #9ca3af; font-weight: 400;
}

/* ════════════════════════════════════════════════════════════
   DEDUP BLOCK
════════════════════════════════════════════════════════════ */
.os-dedup {
    background: #ffffff;
    border: 1.5px solid rgba(99,102,241,0.12);
    border-radius: 16px;
    padding: 18px 22px;
    display: flex; align-items: center; gap: 20px;
    margin: 14px 0 22px;
    box-shadow: 0 2px 12px rgba(99,102,241,0.06);
}
.os-dedup.has-removed {
    border-left: 4px solid #f59e0b;
    background: linear-gradient(135deg, #fffbeb 0%, #fff 40%);
}
.os-dedup.no-removed {
    border-left: 4px solid #22c55e;
    background: linear-gradient(135deg, #f0fdf4 0%, #fff 40%);
}
.os-dedup-icon { font-size: 24px; flex-shrink: 0; }
.os-dedup-body { flex: 1; }
.os-dedup-title {
    font-size: 15px; font-weight: 700; color: #1a1a2e;
    margin-bottom: 8px;
}
.os-dedup-pills { display: flex; gap: 6px; flex-wrap: wrap; }
.os-dedup-pill {
    font-size: 12px; font-weight: 600;
    font-family: 'Inter', sans-serif;
    padding: 4px 12px; border-radius: 100px;
    background: #f0f2fa; color: #6b7280;
}
.os-dedup-pill.removed { background: #fef3c7; color: #b45309; }
.os-dedup-pill.kept    { background: #dcfce7; color: #16a34a; }
.os-dedup-count {
    text-align: right; flex-shrink: 0;
    font-size: 36px; font-weight: 800;
    letter-spacing: -1.5px;
    font-variant-numeric: tabular-nums;
}
.os-dedup-count.warn { color: #f59e0b; }
.os-dedup-count.ok   { color: #22c55e; }
.os-dedup-count-lbl {
    font-size: 12px; color: #9ca3af;
    margin-top: 2px; text-align: right; font-weight: 500;
}

/* ════════════════════════════════════════════════════════════
   SKU HEADER ROW
════════════════════════════════════════════════════════════ */
.os-sku-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 12px 16px;
    background: #ffffff;
    border: 1px solid rgba(99,102,241,0.12);
    border-radius: 12px;
    margin: 22px 0 6px;
    box-shadow: 0 1px 6px rgba(99,102,241,0.06);
}
.os-sku-left { display: flex; align-items: center; gap: 10px; }
.os-sku-code {
    font-family: 'JetBrains Mono', monospace;
    font-size: 13px; font-weight: 600;
    color: #6366f1;
    background: rgba(99,102,241,0.10);
    padding: 5px 12px; border-radius: 8px;
    letter-spacing: 0.3px;
}
.os-sku-label { font-size: 13px; color: #9ca3af; font-weight: 400; }
.os-sku-total {
    font-size: 16px; font-weight: 800;
    color: #1a1a2e; letter-spacing: -0.5px;
    font-variant-numeric: tabular-nums;
}

/* ════════════════════════════════════════════════════════════
   DOWNLOAD BUTTON
════════════════════════════════════════════════════════════ */
.stDownloadButton > button {
    background: linear-gradient(135deg, #6366f1, #818cf8) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 14px !important;
    padding: 14px 28px !important;
    font-weight: 700 !important;
    font-size: 13.5px !important;
    letter-spacing: 0.1px !important;
    font-family: 'Inter', sans-serif !important;
    box-shadow: 0 4px 20px rgba(99,102,241,0.35) !important;
    transition: all 0.2s cubic-bezier(.34,1.56,.64,1) !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    box-shadow: 0 8px 30px rgba(99,102,241,0.50) !important;
    transform: translateY(-2px) !important;
}
.stDownloadButton > button:active {
    transform: translateY(0) !important;
    box-shadow: 0 2px 10px rgba(99,102,241,0.30) !important;
}

/* ════════════════════════════════════════════════════════════
   DOWNLOAD WRAPPER
════════════════════════════════════════════════════════════ */
.os-dl-wrap {
    border: 1.5px solid rgba(99,102,241,0.15);
    border-radius: 18px;
    padding: 20px;
    background: #ffffff;
    margin: 20px 0;
    box-shadow: 0 2px 16px rgba(99,102,241,0.08);
}
.os-dl-meta {
    display: flex; align-items: center; justify-content: space-between;
    margin-top: 12px; padding-top: 12px;
    border-top: 1px solid #f0f2fa;
}
.os-dl-filename {
    font-size: 12px; color: #9ca3af;
    font-family: 'JetBrains Mono', monospace;
}
.os-dl-stats {
    font-size: 13px; color: #6b7280; font-weight: 500;
}

/* ════════════════════════════════════════════════════════════
   EMPTY STATE
════════════════════════════════════════════════════════════ */
.os-empty {
    text-align: center;
    padding: 72px 32px;
    border: 1.5px solid rgba(99,102,241,0.12);
    border-radius: 24px;
    background: #ffffff;
    margin: 14px 0 36px;
    box-shadow: 0 4px 24px rgba(99,102,241,0.08);
}
.os-empty-icon {
    font-size: 40px; margin-bottom: 18px; display: block;
    animation: emptyFloat 3s ease-in-out infinite;
}
@keyframes emptyFloat {
    0%,100% { transform: translateY(0); }
    50%      { transform: translateY(-8px); }
}
.os-empty-title {
    font-size: 20px; font-weight: 800; color: #1a1a2e;
    margin-bottom: 12px; letter-spacing: -0.5px;
}
.os-empty-sub {
    font-size: 15px; color: #9ca3af;
    max-width: 400px; margin: 0 auto 36px;
    line-height: 1.75; font-weight: 400;
}
.os-steps-row {
    display: flex; align-items: flex-start;
    justify-content: center; gap: 0;
}
.os-step-item {
    display: flex; flex-direction: column; align-items: center;
    gap: 8px; padding: 0 22px;
    position: relative;
}
.os-step-item:not(:last-child)::after {
    content: '';
    position: absolute;
    top: 16px; right: -2px;
    width: 30px; height: 2px;
    background: linear-gradient(90deg, rgba(99,102,241,0.4), rgba(99,102,241,0.2));
    border-radius: 1px;
}
.os-step-num {
    width: 34px; height: 34px;
    background: rgba(99,102,241,0.10);
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    font-size: 13px; font-weight: 700; color: #6366f1;
    font-family: 'JetBrains Mono', monospace;
    border: 1.5px solid rgba(99,102,241,0.20);
}
.os-step-txt {
    font-size: 12px; color: #9ca3af; font-weight: 500;
    white-space: nowrap;
}

/* ════════════════════════════════════════════════════════════
   UPLOAD NOTE
════════════════════════════════════════════════════════════ */
.os-note {
    font-size: 13px; color: #9ca3af;
    margin-top: 12px; line-height: 1.7;
    padding: 2px 4px;
}
.os-note strong { color: #6366f1; font-weight: 600; }
.os-note::before { content: 'ℹ️  '; }

/* ════════════════════════════════════════════════════════════
   DIVIDER
════════════════════════════════════════════════════════════ */
.os-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(99,102,241,0.20), transparent);
    margin: 32px 0;
}

/* ════════════════════════════════════════════════════════════
   DATAFRAME
════════════════════════════════════════════════════════════ */
.stDataFrame {
    border: 1px solid rgba(99,102,241,0.12) !important;
    border-radius: 12px !important;
    overflow: hidden !important;
    box-shadow: 0 2px 8px rgba(99,102,241,0.06) !important;
}
iframe { border-radius: 12px; }

/* ════════════════════════════════════════════════════════════
   SECTION TAG
════════════════════════════════════════════════════════════ */
.os-section-tag {
    font-size: 10px; font-weight: 700;
    color: #6366f1;
    background: rgba(99,102,241,0.08);
    padding: 3px 10px; border-radius: 6px;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    font-family: 'JetBrains Mono', monospace;
    margin-left: auto;
}
.os-section-sep { flex: 1; }

/* ════════════════════════════════════════════════════════════
   FILE UPLOADER — VIETNAMESE OVERRIDE
════════════════════════════════════════════════════════════ */
[data-testid="stFileUploaderDropzoneInstructions"] {
    visibility: hidden;
    position: relative;
}
[data-testid="stFileUploaderDropzoneInstructions"] * {
    display: none !important;
}
[data-testid="stFileUploaderDropzoneInstructions"]::before {
    content: 'Kéo thả file vào đây\A Giới hạn 200MB · CSV, XLSX';
    visibility: visible;
    display: block !important;
    white-space: pre-line;
    font-size: 13px; color: #9ca3af; font-weight: 400;
    text-align: center;
    line-height: 1.8;
}
/* Nút chọn file chính (khi chưa upload) */
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button {
    font-size: 0 !important;
    min-width: 120px;
    padding: 8px 20px !important;
    border-radius: 10px !important;
    border: 1.5px solid #e5e8f4 !important;
    background: #fff !important;
    transition: all 0.2s ease !important;
    cursor: pointer;
}
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button:hover {
    border-color: #a5b4fc !important;
    background: rgba(99,102,241,0.04) !important;
    box-shadow: 0 2px 8px rgba(99,102,241,0.10) !important;
}
[data-testid="stFileUploaderDropzone"]:has([data-testid="stFileUploaderDropzoneInstructions"]) button::after {
    content: 'Chọn file' !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    font-family: 'Inter', sans-serif !important;
    color: #6366f1 !important;
}
/* Nút chọn file nhỏ khi đã upload (compact) */
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) {
    padding: 8px 16px !important;
    flex-direction: row !important;
    align-items: center !important;
    gap: 8px !important;
}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) button {
    font-size: 0 !important;
    padding: 6px 14px !important;
    border-radius: 8px !important;
    border: 1px solid #e5e8f4 !important;
    background: #fff !important;
    min-width: unset !important;
}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"])) button::after {
    content: 'Chọn file khác' !important;
    font-size: 12px !important;
    font-weight: 500 !important;
    font-family: 'Inter', sans-serif !important;
    color: #6366f1 !important;
}

/* ════════════════════════════════════════════════════════════
   SECTION NUMBER HOVER
════════════════════════════════════════════════════════════ */
.os-section-num {
    transition: transform 0.2s ease, box-shadow 0.2s ease;
}
.os-section:hover .os-section-num {
    transform: scale(1.08);
    box-shadow: 0 2px 8px rgba(99,102,241,0.25);
}

/* ════════════════════════════════════════════════════════════
   RESPONSIVE
════════════════════════════════════════════════════════════ */
@media (max-width: 768px) {
    .os-topbar { flex-wrap: wrap; gap: 12px; }
    .os-topbar-status { margin-left: auto; }
    .os-metrics { grid-template-columns: 1fr; }
    .os-dedup { flex-direction: column; text-align: center; }
    .os-dedup-count { text-align: center; }
    .os-dedup-count-lbl { text-align: center; }
    .os-steps-row { flex-wrap: wrap; gap: 12px; }
    .os-step-item:not(:last-child)::after { display: none; }
}
</style>
""", unsafe_allow_html=True)

# ─── HÀM TÁCH SKU ─────────────────────────────────────────────────────────────
def parse_sku_from_col_g(val):
    val = str(val).strip()
    return val.split('_')[0] if '_' in val else val

# ─── HÀM TÁCH MÀU & SIZE ──────────────────────────────────────────────────────
def parse_color_size_from_col_i(val):
    val = str(val).strip()
    if ',' in val:
        last_comma_index = val.rfind(',')
        color = val[:last_comma_index].strip()
        size = val[last_comma_index + 1:].strip() if last_comma_index < len(val) - 1 else "F"
        if ':' in size:
            size = size.split(':')[0].strip()
        color = ' '.join(color.replace(',', ' ').split())
    else:
        color = val
        size = "F"
    if ':' in size:
        size = size.split(':')[0].strip()
    return pd.Series([color, size])

# ─── HÀM HELPER FIELD CODE ────────────────────────────────────────────────────
def add_page_number_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    run._element.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    t = OxmlElement('w:t')
    run._element.append(t)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

def add_num_pages_field(paragraph):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'NUMPAGES'
    run._element.append(instrText)
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    t = OxmlElement('w:t')
    run._element.append(t)
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

# ─── HÀM HELPER DOCX ──────────────────────────────────────────────────────────
def set_cell_shading(cell, color_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            element = OxmlElement(f'w:{edge}')
            element.set(qn('w:val'), val.get('val', 'single'))
            element.set(qn('w:sz'), val.get('sz', '4'))
            element.set(qn('w:color'), val.get('color', '000000'))
            element.set(qn('w:space'), '0')
            tcBorders.append(element)
    tcPr.append(tcBorders)

def set_row_height(row, height_cm=0.6):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('start', left), ('end', right)]:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

def format_cell(cell, text, font_size=10, bold=False, align='left', font_name='Arial', color=None, indent=False):
    cell.text = ''
    para = cell.paragraphs[0]
    para_format = para.paragraph_format
    para_format.space_before = Pt(1)
    para_format.space_after = Pt(1)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    prefix = '    ' if indent else ''
    run = para.add_run(prefix + str(text))
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.name = font_name
    if color:
        run.font.color.rgb = RGBColor(*color)
    if align == 'center':
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == 'right':
        para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    else:
        para.alignment = WD_ALIGN_PARAGRAPH.LEFT

# ─── HÀM XUẤT WORD ────────────────────────────────────────────────────────────
def export_to_word(detail_summary, total_orders, total_items, shop_name="TITIKID", platform="SHOPEE", shift="CHIỀU"):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)

    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_para.clear()
    run_page = add_page_number_field(footer_para)
    run_page.font.size = Pt(9)
    run_page.font.name = 'Arial'
    run_sep = footer_para.add_run(' / ')
    run_sep.font.size = Pt(9)
    run_sep.font.name = 'Arial'
    run_total = add_num_pages_field(footer_para)
    run_total.font.size = Pt(9)
    run_total.font.name = 'Arial'

    current_date = datetime.now().strftime('%d.%m')
    title_text = f"{shop_name} - {platform} - {shift} - {current_date} - {total_orders} ĐƠN - {total_items} ÁO"
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title_text)
    title_run.font.size = Pt(13)
    title_run.font.bold = True
    title_run.font.name = 'Arial'
    title_run.font.color.rgb = RGBColor(0, 0, 0)
    title_para.paragraph_format.space_after = Pt(6)

    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    col_widths = [Inches(1.5), Inches(2.5), Inches(2.2), Inches(1.0)]

    hdr_cells = table.rows[0].cells
    headers = ['SKU sản phẩm', 'Màu', 'Size', 'ÁO']
    header_bg = 'C5D0DC'
    for i, header_text in enumerate(headers):
        format_cell(hdr_cells[i], header_text, font_size=10, bold=True, align='left' if i < 3 else 'right', font_name='Arial')
        set_cell_shading(hdr_cells[i], header_bg)
    set_row_height(table.rows[0], 0.9)

    top_total_cells = table.add_row().cells
    top_total_bg = 'DCE6F1'
    format_cell(top_total_cells[0], 'Tổng số', font_size=10, bold=True, font_name='Arial')
    format_cell(top_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[3], '0', font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in top_total_cells:
        set_cell_shading(cell, top_total_bg)
    set_row_height(table.rows[-1], 0.85)

    unique_skus = detail_summary['SKU'].unique()
    for sku in unique_skus:
        sku_data = detail_summary[detail_summary['SKU'] == sku].copy()
        total_sku = int(sku_data['SL'].sum())
        sku_data = sku_data.sort_values(by=['Phân loại', 'Size'])
        for _, data_row in sku_data.iterrows():
            row_cells = table.add_row().cells
            format_cell(row_cells[0], sku, font_size=10, bold=True, font_name='Arial')
            format_cell(row_cells[1], str(data_row['Phân loại']), font_size=10, bold=False, font_name='Arial')
            format_cell(row_cells[2], str(data_row['Size']), font_size=10, font_name='Arial')
            format_cell(row_cells[3], str(int(data_row['SL'])), font_size=10, align='right', font_name='Arial')
            for cell in row_cells:
                set_cell_shading(cell, 'FFFFFF')
            set_row_height(table.rows[-1], 0.85)

        total_row_cells = table.add_row().cells
        format_cell(total_row_cells[0], f'Tổng số {sku}', font_size=10, bold=True, font_name='Arial')
        format_cell(total_row_cells[1], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[2], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[3], str(total_sku), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
        for cell in total_row_cells:
            set_cell_shading(cell, 'DCE6F1')
        set_row_height(table.rows[-1], 0.85)

    grand_total_cells = table.add_row().cells
    format_cell(grand_total_cells[0], 'Tổng cộng', font_size=10, bold=True, font_name='Arial')
    format_cell(grand_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[3], str(total_items), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in grand_total_cells:
        set_cell_shading(cell, 'DCE6F1')
    set_row_height(table.rows[-1], 0.85)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    for row_idx, row in enumerate(table.rows):
        if row_idx == 0:
            trPr = row._tr.get_or_add_trPr()
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)

    target = io.BytesIO()
    doc.save(target)
    return target.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
#  TIÊU ĐỀ
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="os-topbar">
    <div class="os-topbar-brand">
        <div class="os-topbar-logo">📦</div>
        <div>
            <div class="os-topbar-name">Xử lý đơn hàng</div>
            <div class="os-topbar-sub">TikTok &amp; Shopee · Tổng hợp · Lọc trùng · Xuất Word</div>
        </div>
    </div>
    <div class="os-topbar-status">
        <span class="os-status-dot"></span> Sẵn sàng
    </div>
</div>
''', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  01 — CÀI ĐẶT
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="os-section">
    <span class="os-section-num">01</span>
    <span class="os-section-title">Cài đặt</span>
</div>
''', unsafe_allow_html=True)


current_hour = datetime.now().hour
default_shift_index = 0 if current_hour < 12 else 1

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    shop_name = st.selectbox("Tên shop", ["TITIKID", "GIMME"], key="shop_name")
with col_s2:
    platform = st.selectbox("Sàn bán hàng", ["TIKTOK", "SHOPEE"], key="platform")
with col_s3:
    shift = st.selectbox("Ca làm việc", ["SÁNG", "CHIỀU"], index=default_shift_index, key="shift")


# ══════════════════════════════════════════════════════════════════════════════
#  02 — UPLOAD FILE
# ══════════════════════════════════════════════════════════════════════════════
platform_badge = (
    '<span class="badge badge-tiktok">TikTok</span>'
    if platform == "TIKTOK"
    else '<span class="badge badge-shopee">Shopee</span>'
)

st.markdown(f'''
<div class="os-section" style="margin-top: 20px;">
    <span class="os-section-num">02</span>
    <span class="os-section-title">Tải file đơn hàng {platform_badge}</span>
</div>
''', unsafe_allow_html=True)

col_f1, col_f2 = st.columns(2)
with col_f1:
    uploaded_file = st.file_uploader(
        f"📄 File ca hiện tại ({platform}) · Bắt buộc",
        type=["csv", "xlsx"],
        key="file_current",
        help="File đơn hàng ca này cần tổng hợp"
    )
with col_f2:
    prev_file = st.file_uploader(
        "📂 File ca trước · Tùy chọn (lọc đơn trùng)",
        type=["csv", "xlsx"],
        key="file_prev",
        help="Tải lên để loại bỏ đơn đã soạn ở ca trước"
    )

st.markdown('''
<p class="os-note">
    Chỉ cần <strong>file ca hiện tại</strong> là đủ để tổng hợp và xuất Word.
    Muốn <strong>lọc đơn trùng</strong> thì tải thêm file ca trước.
</p>
''', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  03 — KẾT QUẢ
# ══════════════════════════════════════════════════════════════════════════════
if not uploaded_file:
    st.markdown('''
    <div class="os-section" style="margin-top:28px;">
        <span class="os-section-num">03</span>
        <span class="os-section-title">Kết quả</span>
    </div>
    <div class="os-empty">
        <span class="os-empty-icon">📂</span>
        <div class="os-empty-title">Chưa có dữ liệu</div>
        <div class="os-empty-sub">Tải lên file đơn hàng phía trên để bắt đầu tổng hợp, lọc trùng và xuất file Word soạn hàng.</div>
        <div class="os-steps-row">
            <div class="os-step-item">
                <div class="os-step-num">1</div>
                <div class="os-step-txt">Chọn cửa hàng</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">2</div>
                <div class="os-step-txt">Tải file lên</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">3</div>
                <div class="os-step-txt">Lọc trùng ca trước</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">4</div>
                <div class="os-step-txt">Tải Word soạn hàng</div>
            </div>
        </div>
    </div>
    ''', unsafe_allow_html=True)

else:
    try:
        # ── Đọc file hiện tại ────────────────────────────────────────────────
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, engine='calamine', dtype=str)
        df = df.dropna(how='all').reset_index(drop=True)

        # ── Mapping cột theo sàn ─────────────────────────────────────────────
        if platform == "TIKTOK":
            sku_col_index, variation_col_index, qty_col_index = 6, 8, 9
        else:
            sku_col_index, variation_col_index, qty_col_index = 19, 20, 26

        max_col_needed = max(sku_col_index, variation_col_index, qty_col_index)
        if len(df.columns) <= max_col_needed:
            st.error(f"❌ File không đủ cột cho sàn {platform}. Cần ít nhất {max_col_needed + 1} cột, file chỉ có {len(df.columns)} cột.")
            st.stop()

        col_sku       = df.columns[sku_col_index]
        col_variation = df.columns[variation_col_index]
        col_qty       = df.columns[qty_col_index]
        id_col        = df.columns[0]

        total_raw     = df[id_col].nunique()
        removed_count = 0
        prev_total    = 0

        # ── Nhãn mục 03 ─────────────────────────────────────────────────
        st.markdown('''
        <div class="os-section" style="margin-top:28px;">
            <span class="os-section-num">03</span>
            <span class="os-section-title">Kết quả</span>
            <span class="os-section-sep"></span>
            <span class="os-section-tag">xuất dữ liệu</span>
        </div>
        ''', unsafe_allow_html=True)

        # ── Lọc trùng với file buổi trước ───────────────────────────────────
        if prev_file is not None:
            try:
                if prev_file.name.endswith('.csv'):
                    df_prev = pd.read_csv(prev_file, low_memory=False, dtype=str)
                else:
                    df_prev = pd.read_excel(prev_file, engine='calamine', dtype=str)
                df_prev = df_prev.dropna(how='all').reset_index(drop=True)

                prev_id_col  = df_prev.columns[0]
                prev_ids     = set(df_prev[prev_id_col].dropna().astype(str).str.strip())
                current_ids  = set(df[id_col].dropna().astype(str).str.strip())
                duplicated   = current_ids & prev_ids
                removed_count = len(duplicated)
                prev_total   = len(prev_ids)

                df = df[~df[id_col].astype(str).str.strip().isin(duplicated)].reset_index(drop=True)
                kept_count = df[id_col].nunique()

                if removed_count > 0:
                    st.markdown(f'''
                    <div class="os-dedup has-removed">
                        <div class="os-dedup-icon">⇄</div>
                        <div class="os-dedup-body">
                            <div class="os-dedup-title">Đã lọc trùng với file ca trước</div>
                            <div class="os-dedup-pills">
                                <span class="os-dedup-pill">Ca này: {total_raw} đơn</span>
                                <span class="os-dedup-pill">Ca trước: {prev_total} đơn</span>
                                <span class="os-dedup-pill removed">Loại: {removed_count} trùng</span>
                                <span class="os-dedup-pill kept">Giữ: {kept_count} mới</span>
                            </div>
                        </div>
                        <div>
                            <div class="os-dedup-count warn">−{removed_count}</div>
                            <div class="os-dedup-count-lbl">đơn trùng</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown(f'''
                    <div class="os-dedup no-removed">
                        <div class="os-dedup-icon">✓</div>
                        <div class="os-dedup-body">
                            <div class="os-dedup-title">Không có đơn trùng</div>
                            <div class="os-dedup-pills">
                                <span class="os-dedup-pill">Ca này: {total_raw} đơn</span>
                                <span class="os-dedup-pill">Ca trước: {prev_total} đơn</span>
                                <span class="os-dedup-pill kept">Tất cả đều là đơn mới</span>
                            </div>
                        </div>
                        <div>
                            <div class="os-dedup-count ok">0</div>
                            <div class="os-dedup-count-lbl">đơn trùng</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
            except Exception as e_prev:
                st.warning(f"⚠️ Không đọc được file buổi trước: {e_prev}")

        # ── Xử lý data ───────────────────────────────────────────────────────
        df['SKU_ID'] = df[col_sku].apply(parse_sku_from_col_g)
        df[['PL', 'SZ']] = df[col_variation].apply(parse_color_size_from_col_i)
        df['SL'] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0).astype(int)

        total_items  = int(df['SL'].sum())
        total_orders = df[id_col].nunique()

        detail_summary = df.groupby(['SKU_ID', 'PL', 'SZ'])['SL'].sum().reset_index()
        detail_summary.columns = ['SKU', 'Phân loại', 'Size', 'SL']
        unique_skus_count = detail_summary['SKU'].nunique()

        # ── Metric Cards ─────────────────────────────────────────────────────
        dedup_note = f"sau khi lọc {removed_count} trùng" if removed_count > 0 else "đơn hàng duy nhất"
        st.markdown(f'''
        <div class="os-metrics">
            <div class="os-metric">
                <div class="os-metric-icon">📦</div>
                <div class="os-metric-label">Tổng đơn hàng</div>
                <div class="os-metric-value">{total_orders}</div>
                <div class="os-metric-note">{dedup_note}</div>
            </div>
            <div class="os-metric">
                <div class="os-metric-icon">👕</div>
                <div class="os-metric-label">Tổng sản phẩm</div>
                <div class="os-metric-value">{total_items}</div>
                <div class="os-metric-note">tổng số lượng áo</div>
            </div>
            <div class="os-metric">
                <div class="os-metric-icon">🏷️</div>
                <div class="os-metric-label">Loại SKU</div>
                <div class="os-metric-value">{unique_skus_count}</div>
                <div class="os-metric-note">mã sản phẩm khác nhau</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

        # ── Download Word ─────────────────────────────────────────────────────
        current_date_file = datetime.now().strftime('%d.%m')
        word_filename = f"{shop_name}_{platform}_{shift}_{current_date_file}.docx"
        word_data = export_to_word(detail_summary, total_orders, total_items, shop_name, platform, shift)

        st.markdown('<div class="os-dl-wrap">', unsafe_allow_html=True)
        st.download_button(
            f"⬇️  Tải xuống phiếu soạn hàng  —  {total_orders} đơn · {total_items} áo",
            word_data,
            word_filename,
            use_container_width=True
        )
        st.markdown(f'''
            <div class="os-dl-meta">
                <span class="os-dl-filename">{word_filename}</span>
                <span class="os-dl-stats">{total_orders} đơn · {total_items} sản phẩm · {unique_skus_count} mã SKU</span>
            </div>
        </div>''', unsafe_allow_html=True)

        # ── Chi tiết theo SKU ─────────────────────────────────────────────────
        st.markdown('<div class="os-divider"></div>', unsafe_allow_html=True)

        unique_skus = detail_summary['SKU'].unique()
        for sku in unique_skus:
            sku_data  = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
            total_sku = int(sku_data['SL'].sum())

            st.markdown(f'''
            <div class="os-sku-header">
                <div class="os-sku-left">
                    <span class="os-sku-code">{sku}</span>
                    <span class="os-sku-label">sản phẩm</span>
                </div>
                <span class="os-sku-total">{total_sku} cái</span>
            </div>
            ''', unsafe_allow_html=True)

            st.dataframe(
                sku_data[['Phân loại', 'Size', 'SL']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Phân loại": st.column_config.TextColumn("Phân loại / Màu", width="large"),
                    "Size": st.column_config.TextColumn("Size", width="medium"),
                    "SL": st.column_config.NumberColumn("SL", width="small", format="%d")
                }
            )

    except Exception as e:
        st.error(f"❌ Lỗi xử lý file: {e}")