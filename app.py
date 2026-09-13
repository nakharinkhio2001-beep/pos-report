import streamlit as st
import pandas as pd
import numpy as np
import io
import json
from openpyxl.styles import Font

# ==========================================
# 1. ตั้งค่าหน้าเว็บ
# ==========================================
st.set_page_config(page_title="ระบบจัดการหมวดหมู่อาหาร POS", page_icon="🍲", layout="wide")

# หมวดหมู่เริ่มต้น
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

# ดึงข้อมูลหมวดหมู่มาเก็บใน Session (ความจำของเว็บ)
if 'categories' not in st.session_state:
    st.session_state['categories'] = DEFAULT_CATEGORIES


def setup_worksheet_for_printing(worksheet):
    worksheet.page_setup.paperSize = 9
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0


# ==========================================
# 2. ฟังก์ชันประมวลผล Excel (ทำงานบน Memory)
# ==========================================
def process_excel_data(uploaded_file, category_mapping):
    try:
        # อ่านหาบรรทัดหัวตาราง
        df_temp = pd.read_excel(uploaded_file, header=None, nrows=30)
        category_col_name = 'หมวดอาหาร'
        header_row_index = None

        for i in range(len(df_temp)):
            if (df_temp.iloc[i].astype(str).str.strip() == category_col_name).any():
                header_row_index = i
                break

        if header_row_index is None:
            return None, f"ข้อผิดพลาด: ไม่พบหัวข้อคอลัมน์ '{category_col_name}' โปรดตรวจสอบไฟล์"

        header_data = df_temp.iloc[:header_row_index].values

        # รีเซ็ตตำแหน่งอ่านไฟล์ และโหลดข้อมูลเต็ม
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, header=header_row_index)

        # ทำความสะอาดข้อมูล
        df.replace(r'^\s*$', np.nan, regex=True, inplace=True)
        df = df.dropna(axis=0, how='all')
        df = df.dropna(axis=1, how='all')

        cols = list(df.columns)

        if len(cols) >= 6:
            cols[0], cols[1], cols[2] = 'หมวดอาหาร', 'No.', 'รายละเอียดสินค้า'
            cols[-3], cols[-2], cols[-1] = 'จำนวน', 'ราคาต่อหน่วย', 'ราคาอาหารรวม'

            space_count = 1
            for i in range(3, len(cols) - 3):
                cols[i] = " " * space_count
                space_count += 1

            df.columns = cols

            df['รายละเอียดสินค้า'] = df['รายละเอียดสินค้า'].astype(object)
            if 'No.' in df.columns:
                df['No.'] = df['No.'].astype(object)

            for idx in df.index:
                if pd.notna(df.loc[idx, 'จำนวน']) and pd.isna(df.loc[idx, 'No.']):
                    val_detail = df.loc[idx, 'รายละเอียดสินค้า']
                    if pd.isna(val_detail) or str(val_detail).strip() == '':
                        df.loc[idx, 'รายละเอียดสินค้า'] = 'รวม'
        else:
            new_cols = []
            sc = 1
            for c in cols:
                if str(c).startswith("Unnamed:"):
                    new_cols.append(" " * sc)
                    sc += 1
                else:
                    new_cols.append(str(c))
            df.columns = new_cols

        all_defined_categories = [cat for sublist in category_mapping.values() for cat in sublist]

        def assign_category(val):
            val_str = str(val).strip()
            return val_str if val_str in all_defined_categories else np.nan

        first_col = df.columns[0]
        df['__Category_Tracker__'] = df[first_col].apply(assign_category).ffill()

        # สร้างไฟล์ Excel บนหน่วยความจำ (BytesIO)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, categories in category_mapping.items():
                filtered_df = df[df['__Category_Tracker__'].isin(categories)].copy()
                if not filtered_df.empty:
                    filtered_df = filtered_df.drop(columns=['__Category_Tracker__'])
                    filtered_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=header_row_index)
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

            uncategorized_df = df[~df['__Category_Tracker__'].isin(all_defined_categories)].copy()
            uncategorized_df = uncategorized_df.dropna(how='all')
            if not uncategorized_df.empty:
                uncategorized_df = uncategorized_df.drop(columns=['__Category_Tracker__'])
                uncategorized_df.to_excel(writer, sheet_name="ไม่ระบุหมวดหมู่", index=False, startrow=header_row_index)
                worksheet = writer.sheets["ไม่ระบุหมวดหมู่"]
                setup_worksheet_for_printing(worksheet)

                for r_idx, row_vals in enumerate(header_data):
                    current_col = 1
                    for val in row_vals:
                        if pd.notna(val) and str(val).strip() != "":
                            if hasattr(val, 'strftime'): val = val.strftime('%d/%m/%Y')
                            cell = worksheet.cell(row=r_idx + 1, column=current_col, value=val)
                            cell.font = Font(bold=True)
                            current_col += 1

        processed_data = output.getvalue()
        return processed_data, None

    except Exception as e:
        return None, str(e)


# ==========================================
# 3. จัดทำหน้าเว็บ (UI)
# ==========================================
st.title("📊 ระบบแยกหมวดหมู่อาหารและสรุปยอดขาย (ออนไลน์)")
st.markdown("โปรแกรมนี้สามารถใช้งานได้บนทุกเครื่อง ไม่ว่าจะผ่านโทรศัพท์ แท็บเล็ต หรือคอมพิวเตอร์ระบบใดก็ตาม")

tab1, tab2 = st.tabs(["📁 แยกรวมไฟล์ Excel", "⚙️ ตั้งค่าหมวดหมู่อาหาร"])

# --- แถบที่ 1: อัปโหลดและประมวลผล ---
with tab1:
    st.subheader("1. อัปโหลดรายงานจาก POS")
    uploaded_file = st.file_uploader("ลากไฟล์ Excel (.xls, .xlsx) มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์",
                                     type=['xls', 'xlsx'])

    if uploaded_file is not None:
        st.info("อัปโหลดไฟล์สำเร็จ! คลิกปุ่มด้านล่างเพื่อดำเนินการ")

        # ฟังก์ชันสแกนไฟล์
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
                    uploaded_file.seek(0)  # รีเซ็ตไฟล์หลังสแกน
                except Exception as e:
                    st.error(f"สแกนไม่สำเร็จ: {e}")

        # ปุ่มประมวลผลหลัก
        st.markdown("---")
        if st.button("⚡ ประมวลผลและสร้างไฟล์แยกชีท", type="primary"):
            with st.spinner("กำลังจัดระเบียบตาราง และแยกหมวดหมู่..."):
                excel_data, error = process_excel_data(uploaded_file, st.session_state['categories'])

                if error:
                    st.error(f"เกิดข้อผิดพลาด: {error}")
                else:
                    st.success("ประมวลผลสำเร็จ 100%! ตารางเรียงตรงเป๊ะ พร้อมใช้งานแล้ว")
                    st.download_button(
                        label="📥 ดาวน์โหลดไฟล์ Excel (สรุปแยกชีท)",
                        data=excel_data,
                        file_name=f"รายงานสรุปแยกชีท_ออนไลน์.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

# --- แถบที่ 2: ตั้งค่าหมวดหมู่ ---
with tab2:
    st.subheader("⚙️ แก้ไขหมวดหมู่ (แก้ไขด้วยการพิมพ์รูปแบบ JSON)")
    st.markdown("คุณสามารถเพิ่ม/ลด ชื่อเมนูอาหารได้จากกล่องข้อความด้านล่างนี้ แล้วกดปุ่มบันทึก")

    # แสดงกล่องข้อความให้แก้แบบ JSON
    current_json = json.dumps(st.session_state['categories'], ensure_ascii=False, indent=4)
    edited_json = st.text_area("โครงสร้างข้อมูลหมวดหมู่", value=current_json, height=400)

    if st.button("💾 บันทึกการตั้งค่า"):
        try:
            new_mapping = json.loads(edited_json)
            st.session_state['categories'] = new_mapping
            st.success("บันทึกการตั้งค่าสำเร็จ! ข้อมูลนี้จะถูกใช้ในการประมวลผลครั้งต่อไป")
        except json.JSONDecodeError:
            st.error("รูปแบบข้อความไม่ถูกต้อง กรุณาตรวจสอบวงเล็บ ปีกกา หรือเครื่องหมายลูกน้ำให้ครบถ้วน")