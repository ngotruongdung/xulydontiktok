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

# ─── CẤU HÌNH TRANG ───────────────────────────────────────────────────────────
st.set_page_config(page_title="Warehouse Pro", page_icon="📦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* ── Reset & Base ── */
*, *::before, *::after { box-sizing: border-box; }
.stApp { font-family: 'Inter', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.6rem; padding-bottom: 2.5rem; max-width: 980px; }

/* ── HEADER ─────────────────────────────────────────────────────────────── */
.app-header {
    display: flex; align-items: center; justify-content: space-between;
    margin-bottom: 28px; padding-bottom: 22px;
    border-bottom: 1px solid rgba(128,128,128,0.12);
}
.app-header-left { display: flex; align-items: center; gap: 14px; }
.app-header-icon {
    width: 48px; height: 48px; border-radius: 14px;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
    display: flex; align-items: center; justify-content: center;
    font-size: 24px; flex-shrink: 0;
    box-shadow: 0 4px 14px rgba(99,102,241,0.35);
}
.app-header-title { font-size: 22px; font-weight: 800; letter-spacing: -0.4px; margin: 0 0 2px; }
.app-header-sub { font-size: 12.5px; color: rgba(128,128,128,0.6); margin: 0; }
.app-header-badge {
    display: inline-flex; align-items: center; gap: 6px;
    background: rgba(99,102,241,0.1); border: 1px solid rgba(99,102,241,0.2);
    color: #6366f1; font-size: 11px; font-weight: 700;
    padding: 4px 12px; border-radius: 20px; letter-spacing: 0.4px;
}

/* ── STEP SYSTEM ─────────────────────────────────────────────────────────── */
.step-row {
    display: flex; align-items: center; gap: 10px;
    margin: 22px 0 12px;
}
.step-circle {
    width: 28px; height: 28px; border-radius: 50%; flex-shrink: 0;
    background: linear-gradient(135deg, #6366f1, #8b5cf6);
    color: white; font-size: 12px; font-weight: 800;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 2px 8px rgba(99,102,241,0.4);
}
.step-label { font-size: 13.5px; font-weight: 700; letter-spacing: -0.1px; }
.step-hint {
    font-size: 11.5px; color: rgba(128,128,128,0.55);
    margin-left: 4px;
}
.step-line {
    height: 1px; background: rgba(128,128,128,0.1); margin: 0 0 14px;
}

/* ── METRIC CARDS ────────────────────────────────────────────────────────── */
.metric-row {
    display: grid; grid-template-columns: 1fr 1fr 1fr;
    gap: 14px; margin: 4px 0 18px;
}
.metric-card {
    border-radius: 14px; padding: 18px 20px;
    border: 1px solid rgba(128,128,128,0.1);
    background: rgba(128,128,128,0.03);
    position: relative; overflow: hidden;
    transition: transform 0.15s, box-shadow 0.15s;
}
.metric-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.07);
}
.metric-card::after {
    content: ''; position: absolute;
    top: 0; left: 0; right: 0; height: 3px;
    border-radius: 14px 14px 0 0;
}
.metric-card.c-indigo::after { background: linear-gradient(90deg,#6366f1,#818cf8); }
.metric-card.c-emerald::after { background: linear-gradient(90deg,#10b981,#34d399); }
.metric-card.c-amber::after { background: linear-gradient(90deg,#f59e0b,#fbbf24); }
.metric-icon {
    font-size: 20px; margin-bottom: 10px; display: block;
    opacity: 0.85;
}
.metric-label {
    font-size: 10.5px; color: rgba(128,128,128,0.6);
    text-transform: uppercase; letter-spacing: 1px;
    font-weight: 700; margin-bottom: 5px;
}
.metric-value {
    font-size: 32px; font-weight: 800; letter-spacing: -1px;
    margin: 0; line-height: 1;
}
.metric-sub {
    font-size: 11px; color: rgba(128,128,128,0.5);
    margin: 5px 0 0;
}

/* ── DEDUP BANNER ────────────────────────────────────────────────────────── */
.dedup-wrap {
    border-radius: 14px; padding: 16px 20px;
    margin: 4px 0 18px;
    display: flex; align-items: center; gap: 16px;
}
.dedup-wrap.warn {
    border: 1px solid rgba(245,158,11,0.25);
    background: rgba(245,158,11,0.06);
}
.dedup-wrap.ok {
    border: 1px solid rgba(16,185,129,0.25);
    background: rgba(16,185,129,0.06);
}
.dedup-icon { font-size: 28px; flex-shrink: 0; }
.dedup-body { flex: 1; min-width: 0; }
.dedup-title { font-size: 14px; font-weight: 700; margin: 0 0 5px; }
.dedup-pills { display: flex; gap: 8px; flex-wrap: wrap; }
.dedup-pill {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 11.5px; font-weight: 600;
    padding: 3px 10px; border-radius: 20px;
    background: rgba(128,128,128,0.08);
    color: rgba(128,128,128,0.75);
}
.dedup-pill.removed { background: rgba(245,158,11,0.12); color: #d97706; }
.dedup-pill.kept { background: rgba(99,102,241,0.1); color: #6366f1; }
.dedup-stat-block { text-align: right; flex-shrink: 0; }
.dedup-stat-num { font-size: 28px; font-weight: 800; letter-spacing: -1px; line-height: 1; }
.dedup-stat-lbl { font-size: 10.5px; color: rgba(128,128,128,0.55); margin-top: 2px; }

/* ── SKU SECTION ─────────────────────────────────────────────────────────── */
.sku-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 10px 14px; margin: 18px 0 6px;
    border-radius: 10px;
    background: rgba(99,102,241,0.06);
    border: 1px solid rgba(99,102,241,0.12);
    border-left: 4px solid #6366f1;
}
.sku-header-left { display: flex; align-items: center; gap: 10px; }
.sku-tag {
    background: linear-gradient(135deg,#6366f1,#8b5cf6);
    color: white; font-size: 11px; font-weight: 700;
    padding: 3px 10px; border-radius: 6px;
    letter-spacing: 0.3px;
}
.sku-name-label { font-size: 13.5px; font-weight: 600; }
.sku-total {
    font-size: 14px; font-weight: 800;
    color: #6366f1; letter-spacing: -0.3px;
}

/* ── DOWNLOAD BUTTON ─────────────────────────────────────────────────────── */
.stDownloadButton > button {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%) !important;
    color: white !important; border: none !important;
    border-radius: 12px !important;
    padding: 13px 32px !important; font-weight: 700 !important;
    font-size: 14.5px !important; letter-spacing: 0.2px !important;
    box-shadow: 0 4px 18px rgba(99,102,241,0.4) !important;
    transition: opacity 0.2s, transform 0.15s, box-shadow 0.2s !important;
    width: 100% !important;
}
.stDownloadButton > button:hover {
    opacity: 0.92 !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 28px rgba(99,102,241,0.45) !important;
}
.stDownloadButton > button:active { transform: translateY(0) !important; }

/* ── CTA WRAP ────────────────────────────────────────────────────────────── */
.cta-wrap {
    border-radius: 14px; padding: 18px 20px;
    border: 1px solid rgba(99,102,241,0.15);
    background: rgba(99,102,241,0.04);
    margin-bottom: 20px;
}
.cta-info {
    font-size: 12px; color: rgba(128,128,128,0.6);
    margin: 8px 0 0; text-align: center;
}

/* ── EMPTY STATE ─────────────────────────────────────────────────────────── */
.empty-state {
    text-align: center; padding: 52px 24px 44px;
    border-radius: 18px;
    border: 2px dashed rgba(128,128,128,0.13);
    margin: 8px 0 24px;
    background: rgba(128,128,128,0.02);
}
.empty-icon { font-size: 52px; display: block; margin-bottom: 14px; }
.empty-title { font-size: 18px; font-weight: 800; margin: 0 0 8px; }
.empty-sub {
    font-size: 13px; color: rgba(128,128,128,0.6);
    max-width: 380px; margin: 0 auto 28px; line-height: 1.6;
}
.empty-steps { display: flex; justify-content: center; gap: 0; flex-wrap: wrap; }
.empty-step {
    display: flex; flex-direction: column; align-items: center;
    gap: 10px; padding: 0 20px;
    position: relative;
}
.empty-step:not(:last-child)::after {
    content: '→';
    position: absolute; right: -4px; top: 14px;
    font-size: 16px; color: rgba(128,128,128,0.25); font-weight: 700;
}
.empty-step-icon {
    width: 52px; height: 52px; border-radius: 14px;
    background: rgba(99,102,241,0.07);
    border: 1px solid rgba(99,102,241,0.14);
    display: flex; align-items: center; justify-content: center;
    font-size: 22px;
}
.empty-step-label { font-size: 12px; color: rgba(128,128,128,0.6); font-weight: 600; }

/* ── PLATFORM CHIP ───────────────────────────────────────────────────────── */
.chip-tiktok {
    display: inline-block;
    background: rgba(255,0,80,0.08); color: #e0003d;
    border: 1px solid rgba(255,0,80,0.18);
    font-size: 11px; font-weight: 700;
    padding: 2px 10px; border-radius: 20px;
}
.chip-shopee {
    display: inline-block;
    background: rgba(238,77,45,0.08); color: #c9360f;
    border: 1px solid rgba(238,77,45,0.18);
    font-size: 11px; font-weight: 700;
    padding: 2px 10px; border-radius: 20px;
}

/* ── OPTIONAL BADGE ─────────────────────────────────────────────────────── */
.upload-group { position: relative; }
.optional-badge {
    display: inline-flex; align-items: center; gap: 4px;
    background: rgba(128,128,128,0.08);
    border: 1px solid rgba(128,128,128,0.15);
    color: rgba(128,128,128,0.65);
    font-size: 10px; font-weight: 700;
    padding: 2px 8px; border-radius: 20px;
    letter-spacing: 0.4px; margin-bottom: 6px;
    text-transform: uppercase;
}
.required-badge {
    display: inline-flex; align-items: center; gap: 4px;
    background: rgba(99,102,241,0.1);
    border: 1px solid rgba(99,102,241,0.2);
    color: #6366f1;
    font-size: 10px; font-weight: 700;
    padding: 2px 8px; border-radius: 20px;
    letter-spacing: 0.4px; margin-bottom: 6px;
    text-transform: uppercase;
}
.upload-note {
    display: flex; align-items: center; gap: 8px;
    padding: 10px 14px; border-radius: 10px;
    background: rgba(128,128,128,0.04);
    border: 1px solid rgba(128,128,128,0.09);
    margin-top: 10px;
    font-size: 12px; color: rgba(128,128,128,0.65);
}
.upload-note .un-dot {
    width: 6px; height: 6px; border-radius: 50%;
    background: #6366f1; flex-shrink: 0;
}
/* ── MISC ────────────────────────────────────────────────────────────────── */
.stSelectbox label { font-weight: 600 !important; font-size: 12px !important; letter-spacing: 0.1px !important; }
[data-testid="stFileUploader"] label { font-weight: 600 !important; font-size: 12px !important; }
.stDataFrame { border-radius: 10px !important; }
.divider { height: 1px; background: rgba(128,128,128,0.1); margin: 24px 0; }
</style>
""", unsafe_allow_html=True)

# ─── HEADER ───────────────────────────────────────────────────────────────────
st.markdown('''
<div class="app-header">
    <div class="app-header-left">
        <div class="app-header-icon">📦</div>
        <div>
            <div class="app-header-title">Warehouse Pro</div>
            <div class="app-header-sub">Tổng hợp đơn hàng · Lọc trùng · Xuất Word soạn hàng</div>
        </div>
    </div>
    <div class="app-header-badge">⚡ Tự động hoá kho</div>
</div>
''', unsafe_allow_html=True)

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

# ═══════════════════════════════════════════════════════════════════════════════
#  BƯỚC 1 — CÀI ĐẶT
# ═══════════════════════════════════════════════════════════════════════════════
st.markdown('''
<div class="step-row">
    <div class="step-circle">1</div>
    <span class="step-label">Cài đặt xuất file</span>
    <span class="step-hint">· Chọn shop, sàn bán & ca làm việc</span>
</div>
<div class="step-line"></div>
''', unsafe_allow_html=True)

current_hour = datetime.now().hour
default_shift_index = 0 if current_hour < 12 else 1

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    shop_name = st.selectbox("🏪 Tên Shop", ["TITIKID", "GIMME"], key="shop_name")
with col_s2:
    platform = st.selectbox("🛒 Sàn bán hàng", ["TIKTOK", "SHOPEE"], key="platform")
with col_s3:
    shift = st.selectbox("🕐 Ca làm việc", ["SÁNG", "CHIỀU"], index=default_shift_index, key="shift")

# ═══════════════════════════════════════════════════════════════════════════════
#  BƯỚC 2 — UPLOAD FILE
# ═══════════════════════════════════════════════════════════════════════════════
platform_chip = f'<span class="chip-tiktok">TikTok</span>' if platform == "TIKTOK" else f'<span class="chip-shopee">Shopee</span>'
st.markdown(f'''
<div class="step-row" style="margin-top:28px">
    <div class="step-circle">2</div>
    <span class="step-label">Upload đơn hàng {platform_chip}</span>
    <span class="step-hint">· File buổi trước để lọc trùng (tuỳ chọn)</span>
</div>
<div class="step-line"></div>
''', unsafe_allow_html=True)

col_f1, col_f2 = st.columns(2)
with col_f1:
    st.markdown('<span class="required-badge">● Bắt buộc</span>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader(
        f"📂 File ca hiện tại — {platform}",
        type=["csv", "xlsx"],
        key="file_current",
        help="File đơn hàng ca này cần tổng hợp và soạn hàng"
    )
with col_f2:
    st.markdown('<span class="optional-badge">○ Tuỳ chọn</span>', unsafe_allow_html=True)
    prev_file = st.file_uploader(
        "🗂️ File ca trước — Lọc trùng Order ID",
        type=["csv", "xlsx"],
        key="file_prev",
        help="Không bắt buộc. Chỉ upload khi muốn loại bỏ đơn đã soạn ở ca trước"
    )

st.markdown('''
<div class="upload-note">
    <div class="un-dot"></div>
    <span>Chỉ cần upload <strong>file ca hiện tại</strong> là đủ để xuất báo cáo.
    File ca trước chỉ dùng khi muốn <strong>lọc bỏ đơn trùng</strong> giữa 2 ca.</span>
</div>
''', unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  BƯỚC 3 — KẾT QUẢ
# ═══════════════════════════════════════════════════════════════════════════════
if not uploaded_file:
    # ── Empty State ──────────────────────────────────────────────────────────
    st.markdown('''
    <div class="empty-state">
        <span class="empty-icon">📋</span>
        <div class="empty-title">Chưa có dữ liệu</div>
        <div class="empty-sub">Upload file đơn hàng ở trên để bắt đầu tổng hợp, lọc trùng và xuất file Word soạn hàng.</div>
        <div class="empty-steps">
            <div class="empty-step">
                <div class="empty-step-icon">⚙️</div>
                <div class="empty-step-label">Chọn shop & sàn</div>
            </div>
            <div class="empty-step">
                <div class="empty-step-icon">📂</div>
                <div class="empty-step-label">Upload file CSV / XLSX</div>
            </div>
            <div class="empty-step">
                <div class="empty-step-icon">🔁</div>
                <div class="empty-step-label">Lọc trùng ca trước</div>
            </div>
            <div class="empty-step">
                <div class="empty-step-icon">📥</div>
                <div class="empty-step-label">Tải Word soạn hàng</div>
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
            sku_col_index, variation_col_index, qty_col_index = 18, 19, 25

        max_col_needed = max(sku_col_index, variation_col_index, qty_col_index)
        if len(df.columns) <= max_col_needed:
            st.error(f"❌ File không đủ cột cho sàn {platform}. Cần ít nhất {max_col_needed + 1} cột, file chỉ có {len(df.columns)} cột.")
            st.stop()

        col_sku       = df.columns[sku_col_index]
        col_variation = df.columns[variation_col_index]
        col_qty       = df.columns[qty_col_index]
        id_col        = df.columns[0]

        # ── Lọc trùng với file buổi trước ───────────────────────────────────
        total_raw    = df[id_col].nunique()
        removed_count = 0
        prev_total    = 0

        st.markdown('''
        <div class="step-row" style="margin-top:28px">
            <div class="step-circle">3</div>
            <span class="step-label">Kết quả & Xuất file</span>
            <span class="step-hint">· Dữ liệu sau khi xử lý và lọc trùng</span>
        </div>
        <div class="step-line"></div>
        ''', unsafe_allow_html=True)

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
                    <div class="dedup-wrap warn">
                        <div class="dedup-icon">🔁</div>
                        <div class="dedup-body">
                            <div class="dedup-title">Đã lọc trùng với file buổi trước</div>
                            <div class="dedup-pills">
                                <span class="dedup-pill">📂 File hiện tại: {total_raw} đơn</span>
                                <span class="dedup-pill">🗂️ File ca trước: {prev_total} đơn</span>
                                <span class="dedup-pill removed">🗑️ Đã loại: {removed_count} đơn trùng</span>
                                <span class="dedup-pill kept">✅ Còn lại: {kept_count} đơn mới</span>
                            </div>
                        </div>
                        <div class="dedup-stat-block">
                            <div class="dedup-stat-num" style="color:#f59e0b">−{removed_count}</div>
                            <div class="dedup-stat-lbl">đơn trùng loại bỏ</div>
                        </div>
                    </div>
                    ''', unsafe_allow_html=True)
                else:
                    st.markdown(f'''
                    <div class="dedup-wrap ok">
                        <div class="dedup-icon">✅</div>
                        <div class="dedup-body">
                            <div class="dedup-title">Không có đơn trùng</div>
                            <div class="dedup-pills">
                                <span class="dedup-pill">📂 File hiện tại: {total_raw} đơn</span>
                                <span class="dedup-pill">🗂️ File ca trước: {prev_total} đơn</span>
                                <span class="dedup-pill kept">🎉 Tất cả đều là đơn mới</span>
                            </div>
                        </div>
                        <div class="dedup-stat-block">
                            <div class="dedup-stat-num" style="color:#10b981">0</div>
                            <div class="dedup-stat-lbl">đơn trùng</div>
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
        <div class="metric-row">
            <div class="metric-card c-indigo">
                <span class="metric-icon">🛒</span>
                <div class="metric-label">Tổng đơn hàng</div>
                <div class="metric-value">{total_orders}</div>
                <div class="metric-sub">{dedup_note}</div>
            </div>
            <div class="metric-card c-emerald">
                <span class="metric-icon">👕</span>
                <div class="metric-label">Tổng sản phẩm</div>
                <div class="metric-value">{total_items}</div>
                <div class="metric-sub">tổng số lượng áo</div>
            </div>
            <div class="metric-card c-amber">
                <span class="metric-icon">🏷️</span>
                <div class="metric-label">Loại SKU</div>
                <div class="metric-value">{unique_skus_count}</div>
                <div class="metric-sub">mã sản phẩm khác nhau</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)

        # ── Nút tải Word ─────────────────────────────────────────────────────
        current_date_file = datetime.now().strftime('%d.%m')
        word_filename = f"{shop_name}_{platform}_{shift}_{current_date_file}.docx"
        word_data = export_to_word(detail_summary, total_orders, total_items, shop_name, platform, shift)

        st.markdown('<div class="cta-wrap">', unsafe_allow_html=True)
        st.download_button(
            f"📥 TẢI FILE WORD SOẠN HÀNG  —  {total_orders} đơn · {total_items} áo",
            word_data,
            word_filename,
            use_container_width=True
        )
        st.markdown(f'<div class="cta-info">📄 {word_filename}</div></div>', unsafe_allow_html=True)

        # ── Chi tiết theo SKU ────────────────────────────────────────────────
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        unique_skus = detail_summary['SKU'].unique()
        for sku in unique_skus:
            sku_data  = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
            total_sku = int(sku_data['SL'].sum())

            st.markdown(f'''
            <div class="sku-header">
                <div class="sku-header-left">
                    <span class="sku-tag">{sku}</span>
                    <span class="sku-name-label">Sản phẩm</span>
                </div>
                <span class="sku-total">{total_sku} cái</span>
            </div>
            ''', unsafe_allow_html=True)

            st.dataframe(
                sku_data[['Phân loại', 'Size', 'SL']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Phân loại": st.column_config.TextColumn("🏷️ PHÂN LOẠI / MÀU", width="large"),
                    "Size": st.column_config.TextColumn("📏 SIZE", width="medium"),
                    "SL": st.column_config.NumberColumn("🔢 SL", width="small", format="%d")
                }
            )

    except Exception as e:
        st.error(f"❌ Lỗi xử lý file: {e}")