import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from datetime import datetime

# 1. Cấu hình giao diện Web
st.set_page_config(page_title=" Warehouse Pro", layout="wide")

st.markdown("""
    <style>
    /* CSS để làm bảng Web trông sạch sẽ hơn */
    .stDataFrame { border: 1px solid #e6e9ef; border-radius: 10px; }
    .sku-title { color: #1f77b4; font-size: 20px; font-weight: bold; margin-top: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.title("👕 Hệ Thống Soạn Hàng")

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
        # Loại bỏ dấu phẩy trong phần màu (thay bằng khoảng trắng)
        color = color.replace(',', ' ').strip()
        # Loại bỏ khoảng trắng thừa
        color = ' '.join(color.split())
    else:
        color = val
        size = "F"
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

# --- HÀM XUẤT WORD ---
def export_to_word(detail_summary, total_orders, total_items):
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    
    # Thêm số trang vào footer
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_para.clear()
    
    # Thêm số trang hiện tại
    run_page = add_page_number_field(footer_para)
    run_page.font.size = Pt(10)
    run_page.font.name = 'Arial'
    
    # Thêm text " / "
    run_sep = footer_para.add_run(' / ')
    run_sep.font.size = Pt(10)
    run_sep.font.name = 'Arial'
    
    # Thêm tổng số trang
    run_total = add_num_pages_field(footer_para)
    run_total.font.size = Pt(10)
    run_total.font.name = 'Arial'
    
    # Tiêu đề chính với ngày tháng
    current_date = datetime.now().strftime('%d/%m/%Y')
    title = doc.add_heading(f'DANH SÁCH SOẠN HÀNG - {current_date}', 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.runs[0]
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.name = 'Arial'
    
    # Thông tin tổng quan
    info_para = doc.add_paragraph()
    info_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info_para.add_run(f'Tổng đơn: {total_orders} | Tổng áo: {total_items} cái')
    info_run.font.size = Pt(11)
    info_run.font.name = 'Arial'
    info_run.font.bold = True
    
    doc.add_paragraph()  # Khoảng trắng
    
    unique_skus = detail_summary['SKU'].unique()
    for idx, sku in enumerate(unique_skus):
        sku_data = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
        total_sku = int(sku_data['SL'].sum())
        
        # Tiêu đề SKU
        sku_para = doc.add_paragraph()
        sku_run = sku_para.add_run(f'📦 SKU: {sku} — Tổng: {total_sku} cái')
        sku_run.font.size = Pt(12)
        sku_run.font.bold = True
        sku_run.font.name = 'Arial'
        sku_run.font.color.rgb = RGBColor(0, 51, 102)  # Màu xanh đậm
        
        # Tạo bảng
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Light Grid Accent 1'
        
        # Header của bảng
        hdr_cells = table.rows[0].cells
        headers = ['PHÂN LOẠI / MÀU SẮC', 'SIZE', 'SL']
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            hdr_run = hdr_cells[i].paragraphs[0].runs[0]
            hdr_run.font.size = Pt(10)
            hdr_run.font.bold = True
            hdr_run.font.name = 'Arial'
            hdr_run.font.color.rgb = RGBColor(255, 255, 255)  # Màu trắng
            # Màu nền cho header
            tcPr = hdr_cells[i]._element.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:fill'), '4472C4')
            shd.set(qn('w:val'), 'clear')
            tcPr.append(shd)
        
        # Dữ liệu trong bảng
        for _, row in sku_data.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Phân loại'])
            row_cells[1].text = str(row['Size'])
            row_cells[2].text = str(int(row['SL']))
            
            # Định dạng các ô
            for cell in row_cells:
                cell.paragraphs[0].runs[0].font.size = Pt(10)
                cell.paragraphs[0].runs[0].font.name = 'Arial'
                # Căn chỉnh
                if cell == row_cells[1]:  # Cột Size
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif cell == row_cells[2]:  # Cột SL
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Ép độ rộng cột
        widths = [Inches(5.8), Inches(2.0), Inches(1.2)]
        for r in table.rows:
            for i, w in enumerate(widths):
                r.cells[i].width = w
        
        # Khoảng trắng giữa các SKU (trừ SKU cuối)
        if idx < len(unique_skus) - 1:
            doc.add_paragraph()

    target = io.BytesIO()
    doc.save(target)
    return target.getvalue()

uploaded_file = st.file_uploader("Tải file đơn hàng", type=["csv", "xlsx"])

if uploaded_file:
    try:
        # Đọc dữ liệu
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, low_memory=False, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, engine='calamine', dtype=str)

        df = df.dropna(how='all').reset_index(drop=True)
        
        # Lấy các cột theo yêu cầu: G (index 6), I (index 8), J (index 9)
        # Cột G: Seller SKU (lấy phần trước dấu _ đầu tiên)
        # Cột I: Variation (màu trước dấu ,, size sau dấu ,)
        # Cột J: Quantity (số lượng)
        col_g_index = 6  # Cột G (Seller SKU)
        col_i_index = 8  # Cột I (Variation)
        col_j_index = 9  # Cột J (Quantity)
        
        # Kiểm tra số cột có đủ không
        if len(df.columns) <= max(col_g_index, col_i_index, col_j_index):
            st.error(f"File không đủ cột. Cần ít nhất {max(col_g_index, col_i_index, col_j_index) + 1} cột.")
            st.stop()
        
        # Lấy tên cột theo index
        col_g = df.columns[col_g_index]  # Seller SKU
        col_i = df.columns[col_i_index]  # Variation
        col_j = df.columns[col_j_index]  # Quantity
        
        # Lấy cột Order ID để đếm số đơn
        id_col = df.columns[0]  # Cột A (Order ID)
        
        # Xử lý SKU từ cột G (trước dấu _ đầu tiên)
        df['SKU_ID'] = df[col_g].apply(parse_sku_from_col_g)
        
        # Xử lý màu và size từ cột I (trước và sau dấu ,)
        df[['PL', 'SZ']] = df[col_i].apply(parse_color_size_from_col_i)
        
        # Xử lý số lượng từ cột J
        df['SL'] = pd.to_numeric(df[col_j], errors='coerce').fillna(0).astype(int)
        
        # Tính tổng
        total_items = int(df['SL'].sum())
        total_orders = df[id_col].nunique()

        # Dashboard tổng quan
        st.markdown(f"### 📊 Tổng đơn: **{total_orders}** | Tổng áo: **{total_items}** cái")
        
        # Xử lý gôm đơn
        detail_summary = df.groupby(['SKU_ID', 'PL', 'SZ'])['SL'].sum().reset_index()
        detail_summary.columns = ['SKU', 'Phân loại', 'Size', 'SL']

        # Nút tải Word
        word_data = export_to_word(detail_summary, total_orders, total_items)
        st.download_button("📥 TẢI FILE WORD CĂN CHỈNH ĐỀU", word_data, "Gimme_Kho.docx")

        st.divider()

        # --- HIỂN THỊ WEB APP CĂN CHỈNH ĐỀU ---
        unique_skus = detail_summary['SKU'].unique()
        
        for sku in unique_skus:
            sku_data = detail_summary[detail_summary['SKU'] == sku].sort_values(by='Size')
            total_sku = int(sku_data['SL'].sum())
            
            st.markdown(f'<div class="sku-title">📦 SKU: {sku} (Tổng: {total_sku} cái)</div>', unsafe_allow_html=True)
            
            # ĐÂY LÀ PHẦN CĂN CHỈNH WEB: Ép độ rộng các cột cố định
            st.dataframe(
                sku_data[['Phân loại', 'Size', 'SL']],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Phân loại": st.column_config.TextColumn("🏷️ PHÂN LOẠI / MÀU SẮC", width="large"),
                    "Size": st.column_config.TextColumn("📏 SIZE", width="medium"),
                    "SL": st.column_config.NumberColumn("🔢 SL", width="small", format="%d")
                }
            )

    except Exception as e:
        st.error(f"Lỗi: {e}")