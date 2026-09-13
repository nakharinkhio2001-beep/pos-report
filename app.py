import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import os
import urllib.request
from openpyxl.styles import Font
from fpdf import FPDF

# ==========================================
# 1. ตั้งค่าหน้าเว็บ
# ==========================================
st.set_page_config(page_title="ระบบจัดการหมวดหมู่อาหาร POS", page_icon="🍲", layout="wide")

DEFAULT_CATEGORIES = {
    "ครัวไทย": [
        "น้ำพริก", "LINEMAN ทานเล่นไทย", "ไข่เจียว", "อาหารทานเล่นไทย",
        "Lineman ยำ/ลาบ", "ยำ/ลาบ", "ส้มตำ", "หมู", "Lineman ผัดผัก",
        "ผัดผัก", "Lineman ต้มแกง", "ต้มแกง", "Lineman ข้าว+อาหารจานเดียว",
        "Lineman ข้าวผัด", "Lineman ข้าวราด", "ข้าว+อาหารจานเดียว", "ข้าวผัด",
        "ข้าวราด", "เนื้อปลา", "ปลาตัว", "กุ้ง", "เนื้อ", "Lineman กับข้าว",
        "ไก่", "ทะเล", "หมึก", "มังสวิรัติ", "อาหารเจ", "เมนูพิเศษปู",
        "Lineman ก๋วยเตี๋ยว", "เมนูเส้น"
    ],
    "ครัวยุโรป": [
        "Lineman สลัด", "สลัด", "Lineman พิซซ่า", "พิซซ่า", "Lineman พาสต้า",
        "พาสต้า", "Lineman สเต็ก", "สเต็ก", "LINEMAN ทานเล่นฝรั่ง",
        "อาหารทานเล่นฝรั่ง", "MINI พิซซ่า"
    ],
    "เครื่องดื่มและของหวาน": [
        "เหล้า", "เบียร์", "น้ำดื่ม มิกเซอร์", "กาแฟสด", "โกโก้/ชา",
        "น้ำผลไม้เพื่อสุขภาพ", "ของหวาน/เจลาโต้", "เบอเกอรี่", "ไวน์แดง",
        "อิตาเลี่ยนโซดา", "ไวน์ขาว", "Lineman เค้ก/ขนมหวาน", "Lineman กาแฟ",
        "Lineman ชา/นม", "Lineman น้ำผลไม้เพื่อสุขภาพ"
    ],
    "รายการอื่นๆ": [
        "อาการกล่อง", "ค่าห้อง", "โปรโมชั่น", "นอกเมนู", "Stock", "Add-On",
        "ขายวัสดุรีไซเคิล", "set วันแม่", "step 1", "step 2", "step 3",
        "step 4", "step 5", "step 6"
    ]
}

if 'categories' not in st.session_state:
    st.session_state['categories'] = DEFAULT_CATEGORIES

def setup_worksheet_for_printing(worksheet):
    worksheet.page_setup.paperSize = 9
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.column_dimensions['A'].width = 22  
    worksheet.column_dimensions['B'].width = 8   
    worksheet.column_dimensions['C'].width = 45  
    worksheet.column_dimensions['D'].width = 12  
    worksheet.column_dimensions['E'].width = 15  
    worksheet.column_dimensions['F'].width = 18  

# ==========================================
# 2. ฟังก์ชันสร้าง PDF
# ==========================================
@st.cache_resource
def download_fonts():
    regular_url = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Regular.ttf"
    bold_url = "https://github.com/google/fonts/raw/main/ofl/sarabun/Sarabun-Bold.ttf"
    try:
        if not os.path.exists("Sarabun-Regular.ttf"):
            urllib.request.urlretrieve(regular_url, "Sarabun-Regular.ttf")
        if not os.path.exists("Sarabun-Bold.ttf"):
            urllib.request.urlretrieve(bold_url, "Sarabun-Bold.ttf")
        return True
    except Exception as e:
        return False

# ปรับให้ฟังก์ชันรับข้อมูลมาเป็น Dictionary (แยกตามหมวดหมู่)
def generate_pdf(sheets_dict, header_data):
    download_fonts()
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    if os.path.exists("Sarabun-Regular.ttf"):
        pdf.add_font("Sarabun", style="", fname="Sarabun-Regular.ttf")
        pdf.add_font("Sarabun", style="B", fname="Sarabun-Bold.ttf")
        font_name = "Sarabun"
    else:
        font_name = "Arial"
        
    pdf.set_auto_page_break(auto=True, margin=15)
        
    for sheet_name, df_sheet in sheets_dict.items():
        pdf.add_page()
        
        # พิมพ์ส่วนหัวตาราง
        pdf.set_font(font_name, style="B", size=11)
        for row_vals in header_data:
            row_text = " ".join([str(val) for val in row_vals if pd.notna(val) and str(val).strip() != ""])
            pdf.cell(0, 7, txt=row_text, ln=True, align='L')
        
        pdf.ln(3)
        pdf.set_font(font_name, style="B", size=14)
        pdf.cell(0, 10, txt=f"สรุปยอดขาย: {sheet_name}", ln=True, align='C')
        pdf.ln(3)
        
        # หัวคอลัมน์ตาราง
        pdf.set_font(font_name, style="B", size=10)
        col_widths = [35, 12, 70, 15, 25, 30] 
        headers = ['หมวดอาหาร', 'No.', 'รายละเอียดสินค้า', 'จำนวน', 'ราคาต่อหน่วย', 'ราคารวม']
        
        for i, h in enumerate(headers):
            pdf.cell(col_widths[i], 8, txt=h, border=1, align='C')
        pdf.ln()
        
        # ข้อมูลตาราง
        pdf.set_font(font_name, style="", size=10)
        for idx, row in df_sheet.iterrows():
            v_cat = str(row.get('หมวดอาหาร', '')) if pd.notna(row.get('หมวดอาหาร')) else ''
            v_no = str(row.get('No.', '')) if pd.notna(row.get('No.')) else ''
            v_det = str(row.get('รายละเอียดสินค้า', '')) if pd.notna(row.get('รายละเอียดสินค้า')) else ''
            v_qty = row.get('จำนวน', '')
            v_unit = row.get('ราคาต่อหน่วย', '')
            v_tot = row.get('ราคาอาหารรวม', '')
            
            if len(v_det) > 45:
                v_det = v_det[:42] + "..."
            
            def format_num(val):
                if pd.isna(val) or str(val).strip() == '': return ''
                try:
                    f = float(val)
                    if f.is_integer(): return f"{int(f):,}"
                    return f"{f:,.2f}"
                except:
                    return str(val)

            pdf.cell(col_widths[0], 7, txt=v_cat, border=1)
            pdf.cell(col_widths[1], 7, txt=v_no, border=1, align='C')
            pdf.cell(col_widths[2], 7, txt=v_det, border=1)
            pdf.cell(col_widths[3], 7, txt=format_num(v_qty), border=1, align='R')
            pdf.cell(col_widths[4], 7, txt=format_num(v_unit), border=1, align='R')
            pdf.cell(col_widths[5], 7, txt=format_num(v_tot), border=1, align='R')
            pdf.ln()
            
    return bytes(pdf.output())

# ==========================================
# 3. ฟังก์ชันประมวลผล (Rebuild โครงสร้าง)
# ==========================================
def process_excel_data(uploaded_file, category_mapping):
    try:
        df_temp = pd.read_excel(uploaded_file, header=None, nrows=30)
        category_col_name = 'หมวดอาหาร'
        header_row_index = None
        
        for i in range(len(df_temp)):
            if (df_temp.iloc[i].astype(str).str.strip() == category_col_name).any():
                header_row_index = i
                break

        if header_row_index is None:
            return None, None, None, f"ข้อผิดพลาด: ไม่พบหัวข้อ '{category_col_name}' โปรดตรวจสอบไฟล์"

        header_data = df_temp.iloc[:header_row_index].values
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row_index)

        col_names = [str(c).strip() for c in df.columns]
        cat_idx, no_idx, detail_idx = -1, -1, -1
        
        for i, c in enumerate(col_names):
            if 'หมวดอาหาร' in c and cat_idx == -1: cat_idx = i
            elif ('No.' in c or 'ลำดับ' in c) and no_idx == -1: no_idx = i
            elif 'รายละเอียด' in c and detail_idx == -1: detail_idx = i
                
        if cat_idx == -1: cat_idx = 0
        if no_idx == -1: no_idx = 1
        if detail_idx == -1: detail_idx = 2
            
        cleaned_data = []
        for idx, row in df.iterrows():
            if row.isna().all(): continue
            cat_val = row.iloc[cat_idx] if cat_idx < len(row) else np.nan
            no_val = row.iloc[no_idx] if no_idx < len(row) else np.nan
            detail_val = row.iloc[detail_idx] if detail_idx < len(row) else np.nan
            
            num_vals = []
            for i in range(detail_idx + 1, len(row)):
                val = row.iloc[i]
                if pd.notna(val) and str(val).strip() != '' and not str(val).startswith('Unnamed:'):
                    if str(val).strip() == 'รวม': detail_val = 'รวม'
                    else: num_vals.append(val)
                    
            if pd.isna(cat_val) and pd.isna(no_val) and pd.isna(detail_val) and len(num_vals) == 0:
                continue
                
            qty, unit_price, total_price = np.nan, np.nan, np.nan
            if len(num_vals) >= 3:
                qty, unit_price, total_price = num_vals[0], num_vals[1], num_vals[-1] 
            elif len(num_vals) == 2:
                qty, total_price = num_vals[0], num_vals[1]
            elif len(num_vals) == 1:
                total_price = num_vals[0]
                
            if pd.isna(no_val) and (pd.isna(detail_val) or str(detail_val).strip() == '') and pd.notna(total_price):
                detail_val = 'รวม'
                
            cleaned_data.append({
                'หมวดอาหาร': cat_val,
                'No.': no_val,
                'รายละเอียดสินค้า': detail_val,
                'จำนวน': qty,
                'ราคาต่อหน่วย': unit_price,
                'ราคาอาหารรวม': total_price
            })

        df_clean = pd.DataFrame(cleaned_data)
        all_defined_categories = [cat for sublist in category_mapping.values() for cat in sublist]

        def assign_category(val):
            val_str = str(val).strip()
            return val_str if val_str in all_defined_categories else np.nan 

        df_clean['__Category_Tracker__'] = df_clean['หมวดอาหาร'].apply(assign_category).ffill()

        # สร้าง Dictionary แยกข้อมูลแต่ละหมวดหมู่
        sheets_data = {}
        for sheet_name, categories in category_mapping.items():
            filtered_df = df_clean[df_clean['__Category_Tracker__'].isin(categories)].copy()
            if not filtered_df.empty:
                sheets_data[sheet_name] = filtered_df
                
        uncategorized_df = df_clean[~df_clean['__Category_Tracker__'].isin(all_defined_categories)].copy()
        if not uncategorized_df.empty:
            sheets_data["ไม่ระบุหมวดหมู่"] = uncategorized_df

        # 1. สร้างไฟล์ Excel 
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            for sheet_name, df_sheet in sheets_data.items():
                df_to_write = df_sheet.drop(columns=['__Category_Tracker__'])
                df_to_write.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_row_index)
                worksheet = writer.sheets[sheet_name]
                setup_worksheet_for_printing(worksheet)
                for r_idx, row_vals in enumerate(header_data):
                    current_col = 1
                    for val in row_vals:
                        if pd.notna(val) and str(val).strip() != "":
                            if hasattr(val, 'strftime'): val = val.strftime('%d/%m/%Y')
                            cell = worksheet.cell(row=r_idx + 1, column=current_col, value=val)
                            cell.font = Font(bold=True)
                            current_col += 1
        excel_data = output_excel.getvalue()
        
        # 2. สร้าง PDF รวมทุกหมวดหมู่
        pdf_all = generate_pdf(sheets_data, header_data)
        
        # 3. สร้าง PDF แยกแต่ละหมวดหมู่
        pdf_individuals = {}
        for sheet_name, df_sheet in sheets_data.items():
            pdf_individuals[sheet_name] = generate_pdf({sheet_name: df_sheet}, header_data)
            
        return excel_data, pdf_all, pdf_individuals, None

    except Exception as e:
        return None, None, None, str(e)

# ==========================================
# 4. จัดทำหน้าเว็บ (UI)
# ==========================================
st.title("📊 ระบบแยกหมวดหมู่อาหารและสรุปยอดขาย (ออนไลน์)")
st.markdown("โปรแกรมนี้สามารถใช้งานได้บนทุกเครื่อง ไม่ว่าจะผ่านโทรศัพท์ แท็บเล็ต หรือคอมพิวเตอร์ระบบใดก็ตาม")

tab1, tab2 = st.tabs(["📁 แยกรวมไฟล์ Excel", "⚙️ ตั้งค่าหมวดหมู่อาหาร"])

with tab1:
    st.subheader("1. อัปโหลดรายงานจาก POS")
    uploaded_file = st.file_uploader("ลากไฟล์ Excel (.xls, .xlsx) มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์", type=['xls', 'xlsx'])
    
    if uploaded_file is not None:
        st.info("อัปโหลดไฟล์สำเร็จ! คลิกปุ่มด้านล่างเพื่อดำเนินการ")
        
        if st.button("🔍 สแกนตรวจสอบรายการเมนูในไฟล์"):
            with st.spinner("กำลังสแกนไฟล์..."):
                try:
                    df_temp = pd.read_excel(uploaded_file, header=None, nrows=30)
                    header_row_index = None
                    for i in range(len(df_temp)):
                        if (df_temp.iloc[i].astype(str).str.strip() == 'หมวดอาหาร').any():
                            header_row_index = i
                            break
                            
                    uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file, header=header_row_index)
                    first_col = df.columns[0]
                    raw_categories = df[first_col].dropna().astype(str).str.strip().unique()
                    
                    reverse_mapping = {}
                    for sheet, cats in st.session_state['categories'].items():
                        for cat in cats:
                            reverse_mapping[cat] = sheet
                            
                    scan_results = []
                    for item in raw_categories:
                        if item == "" or item.startswith("Unnamed:") or item == 'หมวดอาหาร' or item == "รวม":
                            continue
                        status = reverse_mapping.get(item, "❌ ยังไม่มีหมวดหมู่ (ตกหล่น)")
                        scan_results.append({"ชื่อเมนูที่พบ": item, "สถานะ / อยู่ในชีท": status})
                        
                    st.write("### ผลการสแกน:")
                    df_scan = pd.DataFrame(scan_results)
                    st.dataframe(df_scan, use_container_width=True)
                    uploaded_file.seek(0)
                except Exception as e:
                    st.error(f"สแกนไม่สำเร็จ: {e}")

        st.markdown("---")
        if st.button("⚡ ประมวลผลตารางและสร้าง PDF", type="primary"):
            with st.spinner("กำลังจัดระเบียบตาราง แยกหมวดหมู่ และสร้างไฟล์ PDF (อาจใช้เวลา 5-10 วินาที)..."):
                excel_data, pdf_all, pdf_individuals, error = process_excel_data(uploaded_file, st.session_state['categories'])
                
                if error:
                    st.error(f"เกิดข้อผิดพลาด: {error}")
                else:
                    st.success("ประมวลผลสำเร็จ 100%! กรุณาเลือกดาวน์โหลดไฟล์ที่ต้องการด้านล่างได้เลยครับ")
                    
                    st.markdown("### 📥 1. ดาวน์โหลดไฟล์ต้นฉบับ")
                    st.download_button(
                        label="📊 ดาวน์โหลด Excel (มีแยกชีทครบทุกหมวด)",
                        data=excel_data,
                        file_name="รายงานสรุปแยกชีท_ออนไลน์.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
                    st.markdown("---")
                    st.markdown("### 📄 2. ดาวน์โหลดไฟล์ PDF")
                    st.download_button(
                        label="📑 ดาวน์โหลด PDF (รวมทุกหมวดหมู่ไว้ในไฟล์เดียว)",
                        data=pdf_all,
                        file_name="รายงานสรุปยอดขาย_รวมทุกหมวด.pdf",
                        mime="application/pdf",
                        use_container_width=True
                    )
                    
                    st.markdown("**หรือดาวน์โหลด PDF แยกเฉพาะหมวดหมู่:**")
                    
                    # จัดเรียงปุ่มดาวน์โหลดแยกหมวดหมู่ออกเป็นแถวละ 3 ปุ่ม
                    cats = list(pdf_individuals.keys())
                    for i in range(0, len(cats), 3):
                        cols = st.columns(3)
                        for j in range(3):
                            if i + j < len(cats):
                                cat_name = cats[i+j]
                                pdf_data = pdf_individuals[cat_name]
                                with cols[j]:
                                    st.download_button(
                                        label=f"📄 โหลดเฉพาะ: {cat_name}",
                                        data=pdf_data,
                                        file_name=f"รายงานสรุปยอดขาย_{cat_name}.pdf",
                                        mime="application/pdf",
                                        use_container_width=True
                                    )

with tab2:
    st.subheader("⚙️ แก้ไขหมวดหมู่ (แก้ไขด้วยการพิมพ์รูปแบบ JSON)")
    st.markdown("คุณสามารถเพิ่ม/ลด ชื่อเมนูอาหารได้จากกล่องข้อความด้านล่างนี้ แล้วกดปุ่มบันทึก")
    
    current_json = json.dumps(st.session_state['categories'], ensure_ascii=False, indent=4)
    edited_json = st.text_area("โครงสร้างข้อมูลหมวดหมู่", value=current_json, height=400)
    
    if st.button("💾 บันทึกการตั้งค่า"):
        try:
            new_mapping = json.loads(edited_json)
            st.session_state['categories'] = new_mapping
            st.success("บันทึกการตั้งค่าสำเร็จ! ข้อมูลนี้จะถูกใช้ในการประมวลผลครั้งต่อไป")
        except json.JSONDecodeError:
            st.error("รูปแบบข้อความไม่ถูกต้อง กรุณาตรวจสอบวงเล็บ ปีกกา หรือเครื่องหมายลูกน้ำให้ครบถ้วน")
