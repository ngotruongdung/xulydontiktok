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
st.set_page_config(page_title="Order Studio", page_icon="◈", layout="wide")

# ─── GLOBAL STYLES — 2026 Minimalism ──────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:ital,opsz,wght@0,14..32,300;0,14..32,400;0,14..32,500;0,14..32,600;0,14..32,700;1,14..32,400&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── BASE ────────────────────────────────────────────────────────────────── */
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
html { -webkit-font-smoothing: antialiased; }
.stApp { font-family: 'Inter', system-ui, sans-serif; background: #fafafa; }
#MainMenu, footer, header { visibility: hidden; }
.block-container {
    padding-top: 2rem !important;
    padding-bottom: 4rem !important;
    max-width: 900px !important;
}

/* ── HEADER ──────────────────────────────────────────────────────────────── */
.os-header {
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    padding-bottom: 28px;
    margin-bottom: 36px;
    border-bottom: 1px solid #e5e5e5;
}
.os-wordmark {
    display: flex; align-items: center; gap: 12px;
}
.os-logo {
    width: 36px; height: 36px;
    background: #111;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 16px; color: #fff; font-weight: 700;
    letter-spacing: -1px; font-family: 'JetBrains Mono', monospace;
    flex-shrink: 0;
}
.os-title {
    font-size: 17px; font-weight: 600;
    color: #111; letter-spacing: -0.3px;
    margin-bottom: 2px;
}
.os-sub {
    font-size: 12px; color: #999; font-weight: 400;
}
.os-status {
    display: flex; align-items: center; gap: 6px;
    font-size: 11px; font-weight: 500; color: #666;
    padding: 5px 10px;
    border: 1px solid #e5e5e5;
    border-radius: 6px;
    background: #fff;
}
.os-status-dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: #22c55e;
    animation: pulse 2s infinite;
}
@keyframes pulse {
    0%, 100% { opacity: 1; }
    50% { opacity: 0.4; }
}

/* ── SECTION LABEL ───────────────────────────────────────────────────────── */
.os-section {
    display: flex; align-items: center; gap: 10px;
    margin: 32px 0 16px;
}
.os-section-num {
    font-size: 10px; font-weight: 700;
    font-family: 'JetBrains Mono', monospace;
    color: #aaa; letter-spacing: 0.5px;
    min-width: 20px;
}
.os-section-title {
    font-size: 13px; font-weight: 600; color: #111;
    letter-spacing: -0.1px;
}
.os-section-sep {
    flex: 1; height: 1px; background: #e8e8e8;
}
.os-section-tag {
    font-size: 10px; font-weight: 500; color: #bbb;
    font-family: 'JetBrains Mono', monospace;
}

/* ── CARD BASE ───────────────────────────────────────────────────────────── */
.os-card {
    background: #fff;
    border: 1px solid #e5e5e5;
    border-radius: 10px;
    padding: 20px;
}

/* ── SETTINGS GRID ───────────────────────────────────────────────────────── */
.stSelectbox > div > div {
    border: 1px solid #e5e5e5 !important;
    border-radius: 8px !important;
    background: #fff !important;
    font-size: 13px !important;
    font-family: 'Inter', sans-serif !important;
    transition: border-color 0.15s !important;
}
.stSelectbox > div > div:hover {
    border-color: #bbb !important;
}
.stSelectbox > div > div:focus-within {
    border-color: #111 !important;
    box-shadow: none !important;
}
.stSelectbox label {
    font-size: 11px !important;
    font-weight: 600 !important;
    color: #999 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.6px !important;
    margin-bottom: 6px !important;
}

/* ── FILE UPLOADER ───────────────────────────────────────────────────────── */
[data-testid="stFileUploader"] {
    background: #fff;
    border: 1px dashed #ddd;
    border-radius: 10px;
    padding: 4px;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #aaa;
}
[data-testid="stFileUploader"] label {
    font-size: 11px !important;
    font-weight: 600 !important;
    color: #888 !important;
    text-transform: uppercase !important;
    letter-spacing: 0.6px !important;
}
[data-testid="stFileUploaderDropzone"] {
    border: none !important;
    background: transparent !important;
    padding: 12px 16px !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    font-size: 12px !important;
    color: #aaa !important;
}

/* ── BADGE CHIPS ─────────────────────────────────────────────────────────── */
.badge {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 10px; font-weight: 600;
    padding: 3px 8px; border-radius: 4px;
    letter-spacing: 0.3px; margin-bottom: 8px;
    font-family: 'JetBrains Mono', monospace;
}
.badge-required {
    background: #111; color: #fff;
}
.badge-optional {
    background: #f0f0f0; color: #888;
    border: 1px solid #e5e5e5;
}
.badge-tiktok {
    background: #000; color: #fff;
}
.badge-shopee {
    background: #ee4d2d14; color: #c0310d;
    border: 1px solid #ee4d2d22;
}

/* ── METRIC GRID ─────────────────────────────────────────────────────────── */
.os-metrics {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 12px;
    margin: 20px 0;
}
.os-metric {
    background: #fff;
    border: 1px solid #e5e5e5;
    border-radius: 10px;
    padding: 20px 22px;
    position: relative;
    overflow: hidden;
}
.os-metric::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 2px;
    background: #111;
    opacity: 0.08;
}
.os-metric-label {
    font-size: 10px; font-weight: 600;
    color: #aaa; text-transform: uppercase;
    letter-spacing: 0.7px; margin-bottom: 12px;
}
.os-metric-value {
    font-size: 36px; font-weight: 700;
    color: #111; letter-spacing: -2px;
    line-height: 1; margin-bottom: 4px;
    font-variant-numeric: tabular-nums;
}
.os-metric-note {
    font-size: 11px; color: #bbb; font-weight: 400;
}

/* ── DEDUP BLOCK ─────────────────────────────────────────────────────────── */
.os-dedup {
    background: #fff;
    border: 1px solid #e5e5e5;
    border-radius: 10px;
    padding: 16px 20px;
    display: flex; align-items: center; gap: 20px;
    margin: 12px 0 20px;
}
.os-dedup.has-removed {
    border-left: 3px solid #f59e0b;
}
.os-dedup.no-removed {
    border-left: 3px solid #22c55e;
}
.os-dedup-icon { font-size: 22px; flex-shrink: 0; }
.os-dedup-body { flex: 1; }
.os-dedup-title {
    font-size: 13px; font-weight: 600; color: #111;
    margin-bottom: 6px;
}
.os-dedup-pills { display: flex; gap: 6px; flex-wrap: wrap; }
.os-dedup-pill {
    font-size: 11px; font-weight: 500;
    font-family: 'JetBrains Mono', monospace;
    padding: 2px 8px; border-radius: 4px;
    background: #f5f5f5; color: #666;
}
.os-dedup-pill.removed { background: #fef3c7; color: #b45309; }
.os-dedup-pill.kept { background: #f0fdf4; color: #16a34a; }
.os-dedup-count {
    text-align: right; flex-shrink: 0;
    font-size: 28px; font-weight: 700;
    letter-spacing: -1px;
    font-variant-numeric: tabular-nums;
}
.os-dedup-count.warn { color: #f59e0b; }
.os-dedup-count.ok { color: #22c55e; }
.os-dedup-count-lbl {
    font-size: 10px; color: #bbb;
    margin-top: 2px; text-align: right;
}

/* ── SKU ROW ─────────────────────────────────────────────────────────────── */
.os-sku-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px;
    background: #fff;
    border: 1px solid #e5e5e5;
    border-radius: 8px;
    margin: 20px 0 4px;
}
.os-sku-left { display: flex; align-items: center; gap: 10px; }
.os-sku-code {
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px; font-weight: 500;
    color: #111;
    background: #f5f5f5;
    padding: 3px 8px; border-radius: 4px;
    letter-spacing: 0.2px;
}
.os-sku-label { font-size: 12px; color: #aaa; font-weight: 400; }
.os-sku-total {
    font-size: 14px; font-weight: 700;
    color: #111; letter-spacing: -0.5px;
    font-variant-numeric: tabular-nums;
}

/* ── DOWNLOAD BUTTON ─────────────────────────────────────────────────────── */
.stDownloadButton > button {
    background: #111 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 28px !important;
    font-weight: 600 !important;
    font-size: 13px !important;
    letter-spacing: 0.1px !important;
    font-family: 'Inter', sans-serif !important;
    box-shadow: none !important;
    transition: opacity 0.15s, transform 0.1s !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    opacity: 0.85 !important;
    transform: translateY(-1px) !important;
}
.stDownloadButton > button:active {
    opacity: 1 !important;
    transform: translateY(0) !important;
}

/* ── DOWNLOAD WRAPPER ────────────────────────────────────────────────────── */
.os-dl-wrap {
    border: 1px solid #e5e5e5;
    border-radius: 10px;
    padding: 16px;
    background: #fff;
    margin: 20px 0;
}
.os-dl-meta {
    display: flex; align-items: center; justify-content: space-between;
    margin-top: 10px;
}
.os-dl-filename {
    font-size: 11px; color: #bbb;
    font-family: 'JetBrains Mono', monospace;
}
.os-dl-stats {
    font-size: 11px; color: #aaa; font-weight: 500;
}

/* ── EMPTY STATE ─────────────────────────────────────────────────────────── */
.os-empty {
    text-align: center;
    padding: 64px 32px;
    border: 1px dashed #e0e0e0;
    border-radius: 12px;
    background: #fff;
    margin: 12px 0 32px;
}
.os-empty-icon {
    font-size: 32px; margin-bottom: 16px; display: block; opacity: 0.4;
}
.os-empty-title {
    font-size: 16px; font-weight: 600; color: #111;
    margin-bottom: 8px; letter-spacing: -0.3px;
}
.os-empty-sub {
    font-size: 13px; color: #bbb;
    max-width: 340px; margin: 0 auto 32px;
    line-height: 1.7; font-weight: 400;
}
.os-steps-row {
    display: flex; align-items: flex-start;
    justify-content: center; gap: 0;
}
.os-step-item {
    display: flex; flex-direction: column; align-items: center;
    gap: 8px; padding: 0 24px;
    position: relative;
}
.os-step-item:not(:last-child)::after {
    content: '';
    position: absolute;
    top: 16px; right: -1px;
    width: 2px; height: 2px;
    background: #d0d0d0;
    border-radius: 50%;
}
.os-step-num {
    width: 32px; height: 32px;
    border: 1px solid #e5e5e5;
    border-radius: 8px; background: #ffffff;
    display: flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 600; color: #999;
    font-family: 'JetBrains Mono', monospace;
}
.os-step-txt {
    font-size: 11px; color: #bbb; font-weight: 500;
    white-space: nowrap;
}

/* ── UPLOAD NOTE ─────────────────────────────────────────────────────────── */
.os-note {
    background: #fafafa;
    border: 1px solid #efefef;
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 12px; color: #999;
    margin-top: 12px; line-height: 1.6;
}
.os-note strong { color: #666; font-weight: 600; }

/* ── DIVIDER ─────────────────────────────────────────────────────────────── */
.os-divider {
    height: 1px; background: #efefef; margin: 28px 0;
}

/* ── DATAFRAME ───────────────────────────────────────────────────────────── */
.stDataFrame {
    border: 1px solid #e5e5e5 !important;
    border-radius: 8px !important;
    overflow: hidden !important;
}
iframe { border-radius: 8px; }
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
#  HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="os-header">
    <div class="os-wordmark">
        <div class="os-logo">OS</div>
        <div>
            <div class="os-title">Order Studio</div>
            <div class="os-sub">Tổng hợp · Lọc trùng · Xuất soạn hàng</div>
        </div>
    </div>
    <div class="os-status">
        <div class="os-status-dot"></div>
        Ready
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
    <span class="os-section-sep"></span>
    <span class="os-section-tag">SETUP</span>
</div>
''', unsafe_allow_html=True)

current_hour = datetime.now().hour
default_shift_index = 0 if current_hour < 12 else 1

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    shop_name = st.selectbox("Tên Shop", ["TITIKID", "GIMME"], key="shop_name")
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
<div class="os-section" style="margin-top: 28px;">
    <span class="os-section-num">02</span>
    <span class="os-section-title">Upload file {platform_badge}</span>
    <span class="os-section-sep"></span>
    <span class="os-section-tag">IMPORT</span>
</div>
''', unsafe_allow_html=True)

col_f1, col_f2 = st.columns(2)
with col_f1:
    st.markdown('<span class="badge badge-required">● Required</span>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        f"File ca hiện tại — {platform}",
        type=["csv", "xlsx"],
        key="file_current",
        help="File đơn hàng ca này cần tổng hợp"
    )
with col_f2:
    st.markdown('<span class="badge badge-optional">○ Optional</span>', unsafe_allow_html=True)
    prev_file = st.file_uploader(
        "File ca trước — Lọc trùng Order ID",
        type=["csv", "xlsx"],
        key="file_prev",
        help="Upload để loại bỏ đơn đã soạn ở ca trước"
    )

st.markdown('''
<div class="os-note">
    Chỉ cần <strong>file ca hiện tại</strong> là đủ để xuất báo cáo.
    File ca trước chỉ dùng khi muốn <strong>lọc bỏ đơn trùng</strong> giữa 2 ca.
</div>
''', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
#  03 — KẾT QUẢ
# ══════════════════════════════════════════════════════════════════════════════
if not uploaded_file:
    st.markdown('''
    <div class="os-section" style="margin-top:28px;">
        <span class="os-section-num">03</span>
        <span class="os-section-title">Kết quả</span>
        <span class="os-section-sep"></span>
        <span class="os-section-tag">OUTPUT</span>
    </div>
    <div class="os-empty">
        <span class="os-empty-icon">◈</span>
        <div class="os-empty-title">Chưa có dữ liệu</div>
        <div class="os-empty-sub">Upload file đơn hàng phía trên để bắt đầu tổng hợp, lọc trùng và xuất file Word soạn hàng.</div>
        <div class="os-steps-row">
            <div class="os-step-item">
                <div class="os-step-num">1</div>
                <div class="os-step-txt">Chọn shop & sàn</div>
            </div>
            <div class="os-step-item">
                <div class="os-step-num">2</div>
                <div class="os-step-txt">Upload CSV / XLSX</div>
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

        # ── Section 03 label ─────────────────────────────────────────────────
        st.markdown('''
        <div class="os-section" style="margin-top:28px;">
            <span class="os-section-num">03</span>
            <span class="os-section-title">Kết quả</span>
            <span class="os-section-sep"></span>
            <span class="os-section-tag">OUTPUT</span>
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
                <div class="os-metric-label">Tổng đơn hàng</div>
                <div class="os-metric-value">{total_orders}</div>
                <div class="os-metric-note">{dedup_note}</div>
            </div>
            <div class="os-metric">
                <div class="os-metric-label">Tổng sản phẩm</div>
                <div class="os-metric-value">{total_items}</div>
                <div class="os-metric-note">tổng số lượng áo</div>
            </div>
            <div class="os-metric">
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
            f"↓  Tải file Word soạn hàng  —  {total_orders} đơn · {total_items} áo",
            word_data,
            word_filename,
            use_container_width=True
        )
        st.markdown(f'''
            <div class="os-dl-meta">
                <span class="os-dl-filename">{word_filename}</span>
                <span class="os-dl-stats">{total_orders} orders · {total_items} items · {unique_skus_count} SKUs</span>
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