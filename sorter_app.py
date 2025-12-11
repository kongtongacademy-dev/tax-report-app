import streamlit as st
import pandas as pd
import io
import re

# ฟังก์ชันรันเลข Invoice
def generate_invoice_map(df, start_inv, order_col="Order ID", date_col="Created Time"):
    df_sorted = df.sort_values(by=date_col, ascending=True)
    unique_orders = df_sorted[order_col].unique()
    match = re.match(r"^(.*?)(\d+)$", start_inv)
    if not match: return None, "รูปแบบเลข Invoice ไม่ถูกต้อง"
    prefix, start_num_str = match.group(1), match.group(2)
    current_num = int(start_num_str)
    inv_map = {order_id: f"{prefix}{str(current_num + i).zfill(len(start_num_str))}" for i, order_id in enumerate(unique_orders)}
    return inv_map, None

st.set_page_config(page_title="Tax Report Generator", layout="wide")
st.title("📊 ระบบรายงานภาษีขายฉบับสมบูรณ์")

with st.sidebar:
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ CSV/Excel", type=['xlsx', 'csv'])
    header_row = st.number_input("หัวข้อตารางอยู่บรรทัดที่เท่าไหร่?", min_value=0, value=0)

if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file, header=header_row) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file, header=header_row)
        df.columns = df.columns.str.strip()

        if st.button("🚀 สร้างรายงานภาษีขาย", type="primary"):
            start_invoice = "TINV251100001" # หรือเชื่อมกับ text_input
            inv_map, _ = generate_invoice_map(df, start_invoice)
            
            df_tax = df.copy()
            df_tax['Created Time'] = pd.to_datetime(df_tax['Created Time'], dayfirst=True, errors='coerce')
            df_tax = df_tax.sort_values(by="Created Time", ascending=True)
            df_tax['Invoice No'] = df_tax['Order ID'].map(inv_map)

            # --- ส่วนที่แก้ไข: ลบ THB และแปลงเป็นตัวเลข ---
            cols_to_fix = ['SKU Unit Original Price', 'SKU Seller Discount', 'Shipping Fee After Discount', 'Quantity']
            for col in cols_to_fix:
                if col in df_tax.columns:
                    # ลบ THB, เครื่องหมายคอมม่า และช่องว่างออก
                    df_tax[col] = df_tax[col].astype(str).str.replace('THB', '', regex=False).str.replace(',', '', regex=False).str.strip()
                    df_tax[col] = pd.to_numeric(df_tax[col], errors='coerce').fillna(0)

            # --- คำนวณตามเงื่อนไขของคุณ ---
            # 1. จำนวนเงิน = ราคา x จำนวน
            df_tax['จำนวนเงิน'] = df_tax['SKU Unit Original Price'] * df_tax['Quantity']
            
            # 2. ส่วนลด (ดึงมาจาก SKU Seller Discount)
            df_tax['ส่วนลด'] = df_tax['SKU Seller Discount']
            
            # 3. ค่าขนส่ง (ดึงมาและทำให้เหลือแถวเดียวต่อ Order ID)
            is_duplicate = df_tax.duplicated(subset=['Order ID'], keep='first')
            df_tax['ค่าขนส่ง'] = df_tax['Shipping Fee After Discount']
            df_tax.loc[is_duplicate, 'ค่าขนส่ง'] = 0

            # 4. ยอดก่อนภาษี = (จำนวนเงิน - ส่วนลด + ค่าขนส่ง) / 1.07 
            # (หมายเหตุ: ปกติส่วนลดต้องเอาไปลบยอดขายก่อนบวกค่าส่ง)
            df_tax['ยอดก่อนภาษี'] = (df_tax['จำนวนเงิน'] - df_tax['ส่วนลด'] + df_tax['ค่าขนส่ง']) / 1.07
            
            # 5. VAT = ยอดก่อนภาษี * 7%
            df_tax['VAT'] = df_tax['ยอดก่อนภาษี'] * 0.07

            # จัดการวันที่
            df_tax['Created Time'] = df_tax['Created Time'].dt.strftime('%d/%m/%Y')

            # เรียงคอลัมน์ใหม่
            final_columns = [
                'Invoice No', 'Order ID', 'Created Time', 'SKU ID', 'Product Name', 'Variation', 
                'SKU Unit Original Price', 'Quantity', 'จำนวนเงิน', 'ส่วนลด', 'ค่าขนส่ง', 
                'ยอดก่อนภาษี', 'VAT', 'Order Status'
            ]
            
            df_final = df_tax[final_columns]
            st.success("✅ คำนวณข้อมูลสำเร็จ (ลบค่า THB และคำนวณภาษีแล้ว)")
            st.dataframe(df_final.head(20))

            # ดาวน์โหลด
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False)
            st.download_button("⬇️ ดาวน์โหลดรายงาน", buffer.getvalue(), "Tax_Report_Final.xlsx")

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
