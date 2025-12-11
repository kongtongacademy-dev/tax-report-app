import streamlit as st
import pandas as pd
import io
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# ---------------------------------------------------------
# ฟังก์ชันทำความสะอาดตัวเลข (ลบ THB, ลบลูกน้ำ)
# ---------------------------------------------------------
def clean_currency(x):
    if pd.isna(x):
        return 0.0
    s = str(x)
    # ลบคำว่า THB, ตัวอักษร, เว้นวรรค, ลูกน้ำ (เก็บเฉพาะ 0-9 . -)
    s_clean = re.sub(r'[^\d.-]', '', s)
    try:
        return float(s_clean)
    except ValueError:
        return 0.0

# ---------------------------------------------------------
# ฟังก์ชันส่งอีเมล
# ---------------------------------------------------------
def send_email_with_attachment(sender_email, sender_password, receiver_email, subject, body, file_buffer, filename):
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_buffer.getvalue())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {filename}")
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, receiver_email, text)
        server.quit()
        return True, "✅ ส่งอีเมลสำเร็จ!"
    except Exception as e:
        return False, f"❌ ส่งไม่ผ่าน: {e}"

# ---------------------------------------------------------
# ฟังก์ชันรันเลข Invoice
# ---------------------------------------------------------
def generate_invoice_map(df, start_inv, order_col="Order ID", date_col="Created Time"):
    df_sorted = df.sort_values(by=date_col, ascending=True)
    unique_orders = df_sorted[order_col].unique()
    
    match = re.match(r"^(.*?)(\d+)$", start_inv)
    if not match:
        return None, "รูปแบบเลข Invoice ไม่ถูกต้อง (ต้องลงท้ายด้วยตัวเลข)"
    
    prefix = match.group(1)
    start_num_str = match.group(2)
    num_length = len(start_num_str)
    current_num = int(start_num_str)
    
    inv_map = {}
    for order_id in unique_orders:
        new_inv = f"{prefix}{str(current_num).zfill(num_length)}"
        inv_map[order_id] = new_inv
        current_num += 1
        
    return inv_map, None

# ---------------------------------------------------------
# ตั้งค่าหน้าเว็บ
# ---------------------------------------------------------
st.set_page_config(page_title="Excel Tax Report (Fixed)", layout="wide")
st.title("📊 ระบบจัดการไฟล์ Excel & รายงานภาษีขาย (รองรับ THB)")

# ---------------------------------------------------------
# Sidebar
# ---------------------------------------------------------
with st.sidebar:
    st.header("1. อัปโหลดไฟล์")
    uploaded_file = st.file_uploader("เลือกไฟล์ Excel/CSV ที่นี่", type=['xlsx', 'csv'])
    st.markdown("---")
    st.header("2. ตั้งค่าการอ่าน")
    header_row = st.number_input("หัวข้อตารางอยู่บรรทัดที่เท่าไหร่?", min_value=0, value=0, step=1)

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

        df.columns = df.columns.str.strip()

        # แปลงวันที่
        if "Created Time" in df.columns:
            df["Created Time"] = pd.to_datetime(df["Created Time"], dayfirst=True, errors='coerce')

        tab1, tab2 = st.tabs(["📑 รายงานภาษีขาย (Tax Report)", "🔍 ข้อมูลต้นฉบับ"])

        with tab1:
            st.subheader("สร้างรายงานภาษีขาย (คำนวณยอดเงิน + VAT)")
            
            col_input, _ = st.columns([2, 1])
            with col_input:
                start_invoice = st.text_input("ระบุเลข Invoice ใบแรก", value="TINV251100001")
            
            if 'tax_file_buffer' not in st.session_state:
                st.session_state.tax_file_buffer = None

            if st.button("🚀 ประมวลผลและสร้างรายงาน", type="primary"):
                required_cols = ["Order ID", "Created Time", "SKU ID", "Product Name", "Variation", 
                                 "SKU Unit Original Price", "Quantity", "SKU Seller Discount", 
                                 "Shipping Fee After Discount", "Order Status"]
                
                missing = [c for c in required_cols if c not in df.columns]
                
                if missing:
                    st.error(f"❌ ไม่พบคอลัมน์: {missing}")
                else:
                    inv_map, error = generate_invoice_map(df, start_invoice)
                    if error:
                        st.error(error)
                    else:
                        df_tax = df.copy()
                        df_tax = df_tax.sort_values(by="Created Time", ascending=True)
                        df_tax['Invoice No'] = df_tax['Order ID'].map(inv_map)
                        
                        # ล้างค่าเงิน (THB) ออกก่อนคำนวณ
                        cols_to_clean = ['SKU Unit Original Price', 'Quantity', 'Shipping Fee After Discount', 'SKU Seller Discount']
                        for col in cols_to_clean:
                            df_tax[col] = df_tax[col].apply(clean_currency)
                            
                        # คำนวณต่างๆ
                        df_tax['จำนวนเงิน'] = df_tax['SKU Unit Original Price'] * df_tax['Quantity']
                        
                        # แก้ค่าขนส่งซ้ำ
                        is_duplicate_order = df_tax.duplicated(subset=['Order ID'], keep='first')
                        df_tax.loc[is_duplicate_order, 'Shipping Fee After Discount'] = 0

                        # คำนวณ VAT
                        df_tax['ยอดรวมสุทธิ'] = (df_tax['จำนวนเงิน'] - df_tax['SKU Seller Discount']) + df_tax['Shipping Fee After Discount']
                        df_tax['ยอดก่อนภาษี'] = df_tax['ยอดรวมสุทธิ'] / 1.07
                        df_tax['VAT'] = df_tax['ยอดก่อนภาษี'] * 0.07

                        # จัดการวันที่ (ตัดเวลา)
                        df_tax['Created Time'] = df_tax['Created Time'].dt.strftime('%d/%m/%Y')
                        
                        # จัดเรียงคอลัมน์ (เอา Order Status ไปท้ายสุด)
                        cols_mapping = {
                            'Invoice No': 'Invoice No', 
                            'Order ID': 'Order ID', 
                            'Created Time': 'Created Time',
                            'SKU ID': 'SKU ID', 
                            'Product Name': 'Product Name', 
                            'Variation': 'Variation',
                            'SKU Unit Original Price': 'ราคาต่อหน่วย', 
                            'Quantity': 'จำนวน',
                            'จำนวนเงิน': 'จำนวนเงิน',
                            'SKU Seller Discount': 'ส่วนลด',
                            'Shipping Fee After Discount': 'ค่าขนส่ง',
                            'ยอดรวมสุทธิ': 'ยอดรวมสุทธิ',
                            'ยอดก่อนภาษี': 'ยอดก่อนภาษี',
                            'VAT': 'VAT',
                            'Order Status': 'Order Status'  # <--- ย้ายมาไว้ตรงนี้ครับ (ท้ายสุด)
                        }
                        
                        final_cols_keys = list(cols_mapping.keys())
                        df_final = df_tax[final_cols_keys].rename(columns=cols_mapping)
                        
                        st.success("✅ คำนวณเสร็จสมบูรณ์!")
                        st.dataframe(df_final.head(10))
                        
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False)
                        
                        st.session_state.tax_file_buffer = buffer
                        st.session_state.tax_filename = f"Tax_Report_{start_invoice}.xlsx"

            # --- ส่วนส่งอีเมล ---
            if st.session_state.tax_file_buffer is not None:
                st.divider()
                st.subheader("📧 ดาวน์โหลด / ส่งอีเมล")
                
                col_dl, col_em = st.columns(2)
                
                with col_dl:
                     st.download_button(
                        label="⬇️ ดาวน์โหลดไฟล์ (.xlsx)",
                        data=st.session_state.tax_file_buffer.getvalue(),
                        file_name=st.session_state.tax_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                
                with col_em:
                    with st.expander("ส่งอีเมล"):
                        if "EMAIL_USER" not in st.secrets:
                            st.warning("⚠️ ต้องตั้งค่า Secrets ก่อน")
                        else:
                            recipient = st.text_input("อีเมลปลายทาง")
                            if st.button("ส่งอีเมล"):
                                success, msg = send_email_with_attachment(
                                    st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"],
                                    recipient, f"Tax Report {start_invoice}", "Attached.",
                                    st.session_state.tax_file_buffer, st.session_state.tax_filename
                                )
                                if success: st.success(msg)
                                else: st.error(msg)

        with tab2:
            st.write("ตัวอย่างข้อมูลดิบ:")
            st.dataframe(df.head(50))

    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ที่เมนูด้านซ้าย")
