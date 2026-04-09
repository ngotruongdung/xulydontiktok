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

# 1. Cấu hình giao diện Web
st.set_page_config(page_title="Warehouse Pro", page_icon="📦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

.stApp { font-family: 'Inter', sans-serif; }
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 1.2rem; padding-bottom: 1rem; max-width: 960px; }

/* === COMPACT HEADER === */
.app-header {
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 16px; padding-bottom: 12px;
    border-bottom: 1px solid rgba(128,128,128,0.15);
}
.app-header .app-icon { font-size: 26px; }
.app-header .app-title {
    font-size: 20px; font-weight: 700; color: inherit; margin: 0;
}
.app-header .app-sub {
    font-size: 12px; color: rgba(128,128,128,0.7); margin: 0;
}

/* === METRIC ROW === */
.metric-row {
    display: grid; grid-template-columns: 1fr 1fr 1fr;
    gap: 12px; margin-bottom: 16px;
}
.metric-card {
    border-radius: 10px; padding: 14px 16px;
    border: 1px solid rgba(128,128,128,0.12);
    background: rgba(128,128,128,0.04);
}
.metric-card .metric-label {
    font-size: 11px; color: rgba(128,128,128,0.6);
    text-transform: uppercase; letter-spacing: 0.8px;
    margin-bottom: 2px; font-weight: 500;
}
.metric-card .metric-value {
    font-size: 24px; font-weight: 700; margin: 0; color: inherit;
}

/* === SKU HEADER === */
.sku-row {
    display: flex; align-items: center; justify-content: space-between;
    padding: 6px 0; margin: 10px 0 4px;
    border-bottom: 2px solid rgba(99,102,241,0.15);
}
.sku-row .sku-name {
    font-size: 14px; font-weight: 700; color: inherit;
    display: flex; align-items: center; gap: 6px;
}
.sku-row .sku-badge {
    background: #6366f1; color: #fff;
    font-size: 11px; font-weight: 600;
    padding: 2px 8px; border-radius: 4px;
}
.sku-row .sku-count { font-size: 13px; font-weight: 600; color: #6366f1; }

/* === DOWNLOAD BUTTON === */
.stDownloadButton > button {
    background: #6366f1 !important; color: white !important;
    border: none !important; border-radius: 8px !important;
    padding: 10px 24px !important; font-weight: 600 !important;
    font-size: 14px !important; transition: background 0.2s !important;
}
.stDownloadButton > button:hover { background: #4f46e5 !important; }

/* === SELECTBOX === */
.stSelectbox label { font-weight: 500 !important; font-size: 13px !important; }

/* === DATAFRAME === */
.stDataFrame { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# === COMPACT HEADER ===
st.markdown('''
<div class="app-header">
    <div class="app-icon">📦</div>
    <div>
        <div class="app-title">Warehouse Pro</div>
        <div class="app-sub">Xử lý đơn hàng &amp; soạn hàng</div>
    </div>
</div>
''', unsafe_allow_html=True)

# --- HÀM TÁCH SKU TỪ CỘT G (trước dấu _ đầu tiên) ---
def parse_sku_from_col_g(val):
    val = str(val).strip()
    if '_' in val:
        sku = val.split('_')[0]
    else:
        sku = val
    return sku

# --- HÀM TÁCH MÀU VÀ SIZE TỪ CỘT I (trước và sau dấu , cuối cùng) ---
def parse_color_size_from_col_i(val):
    val = str(val).strip()
    if ',' in val:
        # Tìm dấu phẩy cuối cùng để tách màu và size
        last_comma_index = val.rfind(',')
        color = val[:last_comma_index].strip()  # Tất cả trước dấu phẩy cuối cùng
        size = val[last_comma_index + 1:].strip() if last_comma_index < len(val) - 1 else "F"
        # Chỉ lấy phần trước dấu : (ví dụ: "Size 120 : 17 - 20kg" → "Size 120")
        if ':' in size:
            size = size.split(':')[0].strip()
        # Loại bỏ dấu phẩy trong phần màu (thay bằng khoảng trắng)
        color = color.replace(',', ' ').strip()
        # Loại bỏ khoảng trắng thừa
        color = ' '.join(color.split())
    else:
        color = val
        size = "F"
    # Cắt phần sau dấu : cho size (áp dụng cho cả trường hợp không có dấu phẩy)
    if ':' in size:
        size = size.split(':')[0].strip()
    return pd.Series([color, size])

# --- HÀM HELPER THÊM FIELD CODE VÀO PARAGRAPH ---
def add_page_number_field(paragraph):
    """Thêm field code số trang vào paragraph"""
    run = paragraph.add_run()
    # Begin field
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    
    # Instruction text
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'PAGE'
    run._element.append(instrText)
    
    # Separate
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    
    # Text placeholder
    t = OxmlElement('w:t')
    run._element.append(t)
    
    # End field
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

def add_num_pages_field(paragraph):
    """Thêm field code tổng số trang vào paragraph"""
    run = paragraph.add_run()
    # Begin field
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar1)
    
    # Instruction text
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'NUMPAGES'
    run._element.append(instrText)
    
    # Separate
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    run._element.append(fldChar2)
    
    # Text placeholder
    t = OxmlElement('w:t')
    run._element.append(t)
    
    # End field
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar3)
    return run

# --- HÀM SET CELL SHADING ---
def set_cell_shading(cell, color_hex):
    """Đặt màu nền cho ô trong bảng"""
    tcPr = cell._element.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), color_hex)
    shd.set(qn('w:val'), 'clear')
    tcPr.append(shd)

def set_cell_borders(cell, top=None, bottom=None, left=None, right=None):
    """Đặt border cho ô"""
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
    """Đặt chiều cao cố định cho hàng"""
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(int(height_cm * 567)))  # 1cm = 567 twips
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

def set_cell_margins(cell, top=30, bottom=30, left=60, right=60):
    """Đặt padding cho ô (đơn vị twips)"""
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('start', left), ('end', right)]:
        el = OxmlElement(f'w:{edge}')
        el.set(qn('w:w'), str(val))
        el.set(qn('w:type'), 'dxa')
        tcMar.append(el)
    tcPr.append(tcMar)

def format_cell(cell, text, font_size=10, bold=False, align='left', font_name='Arial', color=None, indent=False):
    """Helper format 1 cell trong bảng"""
    cell.text = ''
    para = cell.paragraphs[0]
    # Set spacing
    para_format = para.paragraph_format
    para_format.space_before = Pt(1)
    para_format.space_after = Pt(1)
    
    # Vertical alignment - căn giữa dọc
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

# --- HÀM XUẤT WORD THEO MẪU DỌC ---
def export_to_word(detail_summary, total_orders, total_items, shop_name="TITIKID", platform="SHOPEE", shift="CHIỀU"):
    doc = Document()
    section = doc.sections[0]
    
    # === PORTRAIT (DỌC) ===
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Inches(8.27)   # A4 width
    section.page_height = Inches(11.69)  # A4 height
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.4)
    section.bottom_margin = Inches(0.4)
    
    # Thêm số trang vào footer
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
    
    # === TIÊU ĐỀ CHÍNH ===
    current_date = datetime.now().strftime('%d.%m')
    title_text = f"{shop_name} - {platform} - {shift} - {current_date} - {total_orders} ĐƠN - {total_items} ÁO"
    
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(title_text)
    title_run.font.size = Pt(13)
    title_run.font.bold = True
    title_run.font.name = 'Arial'
    title_run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Giảm khoảng cách sau tiêu đề
    title_para.paragraph_format.space_after = Pt(6)
    
    # === TẠO 1 BẢNG DUY NHẤT ===
    # 4 cột: SKU sản phẩm | Màu | Size | Áo
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    table.autofit = False
    
    # Đặt độ rộng cột
    col_widths = [Inches(1.5), Inches(2.5), Inches(2.2), Inches(1.0)]
    
    # === HEADER ROW ===
    hdr_cells = table.rows[0].cells
    headers = ['SKU sản phẩm', 'Màu', 'Size', 'ÁO']
    header_bg = 'C5D0DC'  # Xám xanh nhạt giống mẫu
    
    for i, header_text in enumerate(headers):
        format_cell(hdr_cells[i], header_text, font_size=10, bold=True, align='left' if i < 3 else 'right', font_name='Arial')
        set_cell_shading(hdr_cells[i], header_bg)
    set_row_height(table.rows[0], 0.9)
    
    # === DÒNG "Tổng số" Ở ĐẦU (hiện 0) ===
    top_total_cells = table.add_row().cells
    top_total_bg = 'DCE6F1'
    format_cell(top_total_cells[0], 'Tổng số', font_size=10, bold=True, font_name='Arial')
    format_cell(top_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(top_total_cells[3], '0', font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in top_total_cells:
        set_cell_shading(cell, top_total_bg)
    set_row_height(table.rows[-1], 0.85)
    
    # === FILL DATA ===
    unique_skus = detail_summary['SKU'].unique()
    
    for sku in unique_skus:
        sku_data = detail_summary[detail_summary['SKU'] == sku].copy()
        total_sku = int(sku_data['SL'].sum())
        
        # Nhóm theo Màu và Size, sắp xếp
        sku_data = sku_data.sort_values(by=['Phân loại', 'Size'])
        
        for _, data_row in sku_data.iterrows():
            row_cells = table.add_row().cells
            row_bg = 'FFFFFF'
            
            # SKU - hiển thị trên MỖI dòng
            format_cell(row_cells[0], sku, font_size=10, bold=True, font_name='Arial')
            
            # Màu - hiển thị trên MỖI dòng
            format_cell(row_cells[1], str(data_row['Phân loại']), font_size=10, bold=False, font_name='Arial')
            
            # Size
            format_cell(row_cells[2], str(data_row['Size']), font_size=10, font_name='Arial')
            
            # SL
            format_cell(row_cells[3], str(int(data_row['SL'])), font_size=10, align='right', font_name='Arial')
            
            # Nền trắng
            for cell in row_cells:
                set_cell_shading(cell, row_bg)
            set_row_height(table.rows[-1], 0.85)
        
        # === DÒNG TỔNG SỐ CHO MỖI SKU ===
        total_row_cells = table.add_row().cells
        total_bg = 'DCE6F1'  # Xanh nhạt cho dòng tổng
        
        format_cell(total_row_cells[0], f'Tổng số {sku}', font_size=10, bold=True, font_name='Arial')
        format_cell(total_row_cells[1], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[2], '', font_size=10, font_name='Arial')
        format_cell(total_row_cells[3], str(total_sku), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
        
        for cell in total_row_cells:
            set_cell_shading(cell, total_bg)
        set_row_height(table.rows[-1], 0.85)
    
    # === DÒNG TỔNG CỘNG Ở CUỐI ===
    grand_total_cells = table.add_row().cells
    grand_total_bg = 'DCE6F1'
    format_cell(grand_total_cells[0], 'Tổng cộng', font_size=10, bold=True, font_name='Arial')
    format_cell(grand_total_cells[1], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[2], '', font_size=10, font_name='Arial')
    format_cell(grand_total_cells[3], str(total_items), font_size=10, bold=True, align='right', font_name='Arial', color=(0, 0, 139))
    for cell in grand_total_cells:
        set_cell_shading(cell, grand_total_bg)
    set_row_height(table.rows[-1], 0.85)
    
    # === ÉP ĐỘ RỘNG CỘT CHO TẤT CẢ CÁC HÀNG ===
    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = w
    
    # === SET TABLE PROPERTIES ===
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    
    # Cho phép bảng repeat header row khi sang trang mới
    for row_idx, row in enumerate(table.rows):
        if row_idx == 0:  # Header row
            trPr = row._tr.get_or_add_trPr()
            tblHeader = OxmlElement('w:tblHeader')
            trPr.append(tblHeader)

    target = io.BytesIO()
    doc.save(target)
    return target.getvalue()

# === THÔNG TIN XUẤT FILE ===

# Tự động chọn ca theo giờ hiện tại
current_hour = datetime.now().hour
default_shift_index = 0 if current_hour < 12 else 1

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1:
    shop_name = st.selectbox("🏪 Tên Shop", ["TITIKID", "GIMME"], key="shop_name")
with col_s2:
    platform = st.selectbox("🛒 Sàn", ["TIKTOK", "SHOPEE"], key="platform")
with col_s3:
    shift = st.selectbox("🕐 Ca", ["SÁNG", "CHIỀU"], index=default_shift_index, key="shift")

st.markdown("<br>", unsafe_allow_html=True)

uploaded_file = st.file_uploader(f"📂 Tải file đơn hàng ({platform})", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # Đọc dữ liệu
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, engine='calamine', dtype=str)

        df = df.dropna(how='all').reset_index(drop=True)
        
        # === CẤU HÌNH CỘT THEO SÀN ===
        if platform == "TIKTOK":
            # TikTok: Cột G (index 6) = SKU, Cột I (index 8) = Phân loại, Cột J (index 9) = SL
            sku_col_index = 6
            variation_col_index = 8
            qty_col_index = 9
        else:
            # Shopee: Cột S (index 18) = SKU, Cột T (index 19) = Phân loại, Cột Z (index 25) = SL
            sku_col_index = 18
            variation_col_index = 19
            qty_col_index = 25
        
        # Kiểm tra số cột có đủ không
        max_col_needed = max(sku_col_index, variation_col_index, qty_col_index)
        if len(df.columns) <= max_col_needed:
            st.error(f"❌ File không đủ cột cho sàn {platform}. Cần ít nhất {max_col_needed + 1} cột, file chỉ có {len(df.columns)} cột.")
            st.stop()
        
        # Lấy tên cột theo index
        col_sku = df.columns[sku_col_index]
        col_variation = df.columns[variation_col_index]
        col_qty = df.columns[qty_col_index]
        
        # Lấy cột Order ID để đếm số đơn
        id_col = df.columns[0]  # Cột A (Order ID)
        
        # Xử lý SKU (trước dấu _ đầu tiên)
        df['SKU_ID'] = df[col_sku].apply(parse_sku_from_col_g)
        
        # Xử lý màu và size (trước và sau dấu ,)
        df[['PL', 'SZ']] = df[col_variation].apply(parse_color_size_from_col_i)
        
        # Xử lý số lượng
        df['SL'] = pd.to_numeric(df[col_qty], errors='coerce').fillna(0).astype(int)
        
        # Tính tổng
        total_items = int(df['SL'].sum())
        total_orders = df[id_col].nunique()

        # Dashboard tổng quan - Metric Cards
        unique_skus_count = detail_summary['SKU'].nunique()
        st.markdown(f'''
        <div class="metric-row">
            <div class="metric-card">
                <div class="metric-label">🛒 Tổng đơn</div>
                <div class="metric-value">{total_orders}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">👕 Tổng SP</div>
                <div class="metric-value">{total_items}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">📋 Loại SKU</div>
                <div class="metric-value">{unique_skus_count}</div>
            </div>
        </div>
        ''', unsafe_allow_html=True)
        
        # Xử lý gôm đơn
        detail_summary = df.groupby(['SKU_ID', 'PL', 'SZ'])['SL'].sum().reset_index()
        detail_summary.columns = ['SKU', 'Phân loại', 'Size', 'SL']

        # Nút tải Word
        current_date_file = datetime.now().strftime('%d.%m')
        word_filename = f"{shop_name}_{platform}_{shift}_{current_date_file}.docx"
        word_data = export_to_word(detail_summary, total_orders, total_items, shop_name, platform, shift)
        st.download_button("📥 TẢI FILE WORD SOẠN HÀNG", word_data, word_filename)

        st.markdown("<br>", unsafe_allow_html=True)

        # --- HIỂN THỊ WEB APP CĂN CHỈNH ĐỀU ---
        unique_skus = detail_summary['SKU'].unique()
        
        for sku in unique_skus:
            sku_data = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
            total_sku = int(sku_data['SL'].sum())
            
            st.markdown(f'''
            <div class="sku-row">
                <div class="sku-name"><span class="sku-badge">{sku}</span> Sản phẩm</div>
                <div class="sku-count">{total_sku} cái</div>
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
        st.error(f"Lỗi: {e}")