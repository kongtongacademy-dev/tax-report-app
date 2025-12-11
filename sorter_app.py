import streamlit as st
import pandas as pd
import io
import re

# ---------------------------------------------------------
# ฟังก์ชันช่วยรันเลข Invoice
# ---------------------------------------------------------
def generate_invoice_map(df, start_inv, order_col="Order ID", date_col="Created Time"):
    # เรียงข้อมูลตามวันที่
    df_sorted = df.sort_values(by=date_col, ascending=True)
    # หา Order ID ที่ไม่ซ้ำกัน
    unique_orders = df_sorted[order_col].unique()
    
    # แยก Prefix (ตัวอักษร) และ Number (ตัวเลข)
    match = re.match(r"^(.*?)(\d+)$", start_inv)
    if not match:
        return None, "รูปแบบเลข Invoice ไม่ถูกต้อง (ต้องลงท้ายด้วยตัวเลข)"
    
    prefix = match.group(1)
    start_num_str = match.group(2)
    num_length = len(start_num_str)
    current_num = int(start_num_str)
    
    # สร้างการจับคู่ Order ID -> Invoice No
    inv_map = {}
    for order_id in unique_orders:
        new_inv = f"{prefix}{str(current_num).zfill(num_length)}"
        inv_map[order_id] = new_inv
        current_num += 1
        
    return inv_map, None

# ---------------------------------------------------------
# ตั้งค่าหน้าเว็บ
# ---------------------------------------------------------
st.set_page_config(page_title="Excel Tax Report Generator", layout="wide")
st.title("📊 ระบบจัดการไฟล์ Excel & รายงานภาษีขาย")

# ---------------------------------------------------------
# Sidebar
# ---------------------------------------------------------
with st.sidebar:
    st.header("1. อัปโหลดไฟล์")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel/CSV ที่นี่", type=['xlsx', 'csv'])
    
    st.markdown("---")
    st.header("2. ตั้งค่าการอ่านไฟล์")
    # ไฟล์ CSV ส่วนใหญ่มักจะมีหัวข้ออยู่บรรทัดแรก (0) แต่ถ้ามาจาก Shopee บางทีเป็น 1
    header_row = st.number_input(
        "หัวข้อตารางอยู่บรรทัดที่เท่าไหร่? (ปกติใช้เลข 0 หรือ 1)", 
        min_value=0, value=0, step=1,
        help="ลองเปลี่ยนเลขนี้ถ้าโหลดแล้วไม่เจอชื่อคอลัมน์"
    )

# ---------------------------------------------------------
# Main Logic
# ---------------------------------------------------------
if uploaded_file is not None:
    try:
        # อ่านไฟล์
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=header_row)
        else:
            df = pd.read_excel(uploaded_file, header=header_row)

        # [สำคัญ] ลบช่องว่างหัวท้ายชื่อคอลัมน์ (แก้ปัญหาชื่อ ' Order ID ' มีวรรค)
        df.columns = df.columns.str.strip()

        # แปลงวันที่ (รองรับทั้งแบบไทยและสากล)
        if "Created Time" in df.columns:
            df["Created Time"] = pd.to_datetime(df["Created Time"], dayfirst=True, errors='coerce')

        # สร้าง Tabs
        tab1, tab2 = st.tabs(["📑 รายงานภาษีขาย (Tax Report)", "🔍 เช็คข้อมูลต้นฉบับ"])

        # =========================================================
        # TAB 1: สร้างรายงานภาษีขาย
        # =========================================================
        with tab1:
            st.subheader("สร้างรายงานภาษีขาย (Invoice Running + Fix Shipping Fee)")
            
            col_input, col_btn = st.columns([2, 1])
            with col_input:
                start_invoice = st.text_input("ระบุเลข Invoice ใบแรก", value="TINV251100001")
            
            if st.button("🚀 สร้างรายงานภาษีขาย", type="primary"):
                # คอลัมน์ที่จำเป็น (Mapping ชื่อตามไฟล์ CSV ของคุณ)
                required_cols = [
                    "Order ID", "Created Time", "SKU ID", "Product Name", "Variation", 
                    "SKU Unit Original Price", "Quantity", "SKU Seller Discount", 
                    "Shipping Fee After Discount", "Order Status"
                ]
                
                # เช็คว่าคอลัมน์ไหนหายไปบ้าง
                missing = [c for c in required_cols if c not in df.columns]
                
                if missing:
                    st.error(f"❌ ไม่พบคอลัมน์: {missing}")
                    st.warning("คำแนะนำ: ลองเปลี่ยนตัวเลข 'บรรทัดหัวข้อ' ด้านซ้าย (ลองเปลี่ยนเป็น 0 หรือ 1)")
                    st.write("คอลัมน์ที่เจอในไฟล์ตอนนี้:", df.columns.tolist())
                else:
                    # 1. สร้าง Invoice Map
                    inv_map, error = generate_invoice_map(df, start_invoice)
                    
                    if error:
                        st.error(error)
                    else:
                        df_tax = df.copy()
                        
                        # 2. เรียงวันที่ (เก่า -> ใหม่)
                        df_tax = df_tax.sort_values(by="Created Time", ascending=True)
                        
                        # 3. ใส่เลข Invoice
                        df_tax['Invoice No'] = df_tax['Order ID'].map(inv_map)
                        
                        # 4. คำนวณยอดเงิน (แปลงเป็นตัวเลขก่อนคำนวณ)
                        for col in ['SKU Unit Original Price', 'Quantity', 'Shipping Fee After Discount', 'SKU Seller Discount']:
                            df_tax[col] = pd.to_numeric(df_tax[col], errors='coerce').fillna(0)
                            
                        df_tax['จำนวนเงิน'] = df_tax['SKU Unit Original Price'] * df_tax['Quantity']
                        
                        # ==========================================================
                        # 5. แก้ค่าขนส่งซ้ำ (ให้เหลือแค่แถวแรกของ Order ID นั้น)
                        # ==========================================================
                        is_duplicate_order = df_tax.duplicated(subset=['Order ID'], keep='first')
                        df_tax.loc[is_duplicate_order, 'Shipping Fee After Discount'] = 0
                        # ==========================================================

                        # 6. เลือกและเปลี่ยนชื่อคอลัมน์ (เรียงตามที่คุณต้องการ)
                        cols_mapping = {
                            'Invoice No': 'Invoice No',
                            'Order ID': 'Order ID',
                            'Created Time': 'Created Time',
                            'SKU ID': 'SKU ID',
                            'Product Name': 'Product Name',
                            'Variation': 'Variation',
                            'SKU Unit Original Price': 'SKU Unit Original Price',
                            'Quantity': 'Quantity',
                            'จำนวนเงิน': 'จำนวนเงิน',
                            'SKU Seller Discount': 'ส่วนลด',
                            'Shipping Fee After Discount': 'ค่าขนส่ง',
                            'Order Status': 'Order Status'
                        }
                        
                        final_cols = list(cols_mapping.keys())
                        df_final = df_tax[final_cols].rename(columns=cols_mapping)
                        
                        st.success("✅ สร้างรายงานสำเร็จ!")
                        st.dataframe(df_final.head(20))
                        
                        # ปุ่มดาวน์โหลด
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False)
                            
                        st.download_button(
                            label="⬇️ ดาวน์โหลดรายงาน (.xlsx)",
                            data=buffer.getvalue(),
                            file_name=f"Tax_Report_{start_invoice}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

        # =========================================================
        # TAB 2: ดูข้อมูลต้นฉบับ
        # =========================================================
        with tab2:
            st.write("ข้อมูลดิบที่อ่านได้จากไฟล์:")
            st.dataframe(df.head(50))

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
        st.info("ลองเปลี่ยนตัวเลข 'บรรทัดหัวข้อ' ที่เมนูด้านซ้ายดูครับ")
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ CSV ที่เมนูด้านซ้าย")
