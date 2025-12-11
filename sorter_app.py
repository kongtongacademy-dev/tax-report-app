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
# ฟังก์ชันส่งอีเมล (Gmail SMTP)
# ---------------------------------------------------------
# (โค้ดส่วนนี้ไม่ได้เปลี่ยนแปลง)
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
# (โค้ดส่วนนี้ไม่ได้เปลี่ยนแปลง)
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
st.set_page_config(page_title="Excel Tax Report & Email", layout="wide")
st.title("📊 ระบบจัดการไฟล์ Excel & รายงานภาษีขาย (VAT)")

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
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, header=header_row)
        else:
            df = pd.read_excel(uploaded_file, header=header_row)

        df.columns = df.columns.str.strip() # ลบวรรคหัวท้ายชื่อคอลัมน์

        # แปลงวันที่
        if "Created Time" in df.columns:
            df["Created Time"] = pd.to_datetime(df["Created Time"], dayfirst=True, errors='coerce')

        tab1, tab2 = st.tabs(["📑 รายงานภาษีขาย (Tax Report)", "🔍 ข้อมูลต้นฉบับ"])

        with tab1:
            st.subheader("สร้างรายงานภาษีขายและคำนวณ VAT")
            
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
                        
                        # [NEW] แปลงคอลัมน์เงินให้เป็นตัวเลขและเติม 0 ถ้าว่าง (เพื่อการคำนวณ)
                        for col in ['SKU Unit Original Price', 'Quantity', 'Shipping Fee After Discount', 'SKU Seller Discount']:
                            df_tax[col] = pd.to_numeric(df_tax[col], errors='coerce').fillna(0)
                            
                        # 1. คำนวณจำนวนเงิน (ราคาต่อหน่วย * จำนวน)
                        df_tax['จำนวนเงิน'] = df_tax['SKU Unit Original Price'] * df_tax['Quantity']
                        
                        # 2. แก้ค่าขนส่งซ้ำ (ให้เหลือแค่แถวแรกของ Order ID นั้น)
                        is_duplicate_order = df_tax.duplicated(subset=['Order ID'], keep='first')
                        df_tax.loc[is_duplicate_order, 'Shipping Fee After Discount'] = 0

                        # 3. คำนวณยอดรวมสุทธิ (Total Net)
                        # Total Net = (จำนวนเงิน - ส่วนลด) + ค่าขนส่ง
                        df_tax['ยอดรวมสุทธิ'] = (
                            df_tax['จำนวนเงิน'] - df_tax['SKU Seller Discount']
                        ) + df_tax['Shipping Fee After Discount']
                        
                        # 4. คำนวณยอดก่อนภาษี (Tax Base)
                        # Tax Base = ยอดรวมสุทธิ / 1.07
                        df_tax['ยอดก่อนภาษี'] = df_tax['ยอดรวมสุทธิ'] / 1.07
                        
                        # 5. คำนวณ VAT (7%)
                        # VAT = ยอดก่อนภาษี * 0.07
                        df_tax['VAT'] = df_tax['ยอดก่อนภาษี'] * 0.07

                        # 6. จัดการวันที่: ตัดเวลาทิ้ง เหลือแค่ dd/mm/yyyy
                        df_tax['Created Time'] = df_tax['Created Time'].dt.strftime('%d/%m/%Y')
                        
                        # 7. จัดเรียงและเปลี่ยนชื่อคอลัมน์สุดท้าย
                        cols_mapping = {
                            # ข้อมูลหลัก
                            'Invoice No': 'Invoice No', 'Order ID': 'Order ID', 'Created Time': 'Created Time',
                            'SKU ID': 'SKU ID', 'Product Name': 'Product Name', 'Variation': 'Variation',
                            'Order Status': 'Order Status',
                            # ราคา/จำนวน
                            'SKU Unit Original Price': 'ราคาต่อหน่วย', 'Quantity': 'จำนวน',
                            # การเงิน
                            'จำนวนเงิน': 'จำนวนเงิน',
                            'SKU Seller Discount': 'ส่วนลด',
                            'Shipping Fee After Discount': 'ค่าขนส่ง',
                            # ผลลัพธ์บัญชี
                            'ยอดรวมสุทธิ': 'ยอดรวมสุทธิ',
                            'ยอดก่อนภาษี': 'ยอดก่อนภาษี',
                            'VAT': 'VAT'
                        }
                        
                        # เลือกเฉพาะคอลัมน์ที่ต้องการ และใช้ชื่อที่กำหนด
                        final_cols_keys = list(cols_mapping.keys())
                        df_final = df_tax[final_cols_keys].rename(columns=cols_mapping)
                        
                        st.success("✅ สร้างรายงานและคำนวณภาษีเสร็จสมบูรณ์!")
                        st.dataframe(df_final.head(10))
                        
                        # เตรียมไฟล์ลง Buffer
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df_final.to_excel(writer, index=False)
                        
                        st.session_state.tax_file_buffer = buffer
                        st.session_state.tax_filename = f"Tax_Report_{start_invoice}.xlsx"

            # --- ส่วนส่งอีเมล ---
            if st.session_state.tax_file_buffer is not None:
                st.divider()
                st.subheader("📧 ส่งไฟล์ทางอีเมล / ดาวน์โหลด")
                
                col_download, col_email = st.columns([1, 1])

                with col_download:
                    # ปุ่มดาวน์โหลดปกติ
                    st.download_button(
                        label="⬇️ ดาวน์โหลดไฟล์ลงเครื่อง",
                        data=st.session_state.tax_file_buffer.getvalue(),
                        file_name=st.session_state.tax_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary"
                    )

                with col_email:
                    if "EMAIL_USER" not in st.secrets or "EMAIL_PASSWORD" not in st.secrets:
                         st.warning("⚠️ ตั้งค่าอีเมลก่อนจึงจะส่งได้ (ดูวิธีด้านล่าง)")
                    else:
                        with st.popover("ส่งอีเมลพร้อมไฟล์แนบ"):
                            recipient = st.text_input("ส่งไปที่อีเมล:", placeholder="accountant@company.com")
                            email_subject = st.text_input("หัวข้ออีเมล:", value=f"รายงานภาษีขาย {start_invoice}")
                            
                            if st.button("📨 ยืนยันการส่งอีเมล"):
                                if not recipient:
                                    st.error("กรุณากรอกอีเมลปลายทาง")
                                else:
                                    with st.spinner("กำลังส่งอีเมล..."):
                                        success, msg = send_email_with_attachment(
                                            st.secrets["EMAIL_USER"],
                                            st.secrets["EMAIL_PASSWORD"],
                                            recipient,
                                            email_subject,
                                            "ไฟล์รายงานภาษีขายที่คำนวณ VAT แล้ว แนบมาพร้อมกับอีเมลนี้ครับ",
                                            st.session_state.tax_file_buffer,
                                            st.session_state.tax_filename
                                        )
                                        if success:
                                            st.success(msg)
                                        else:
                                            st.error(msg)
        with tab2:
            st.write("ข้อมูลดิบที่อ่านได้จากไฟล์:")
            st.dataframe(df.head(5))

    except Exception as e:
        st.error(f"Error: {e}")
        st.info("ลองเปลี่ยนตัวเลข 'บรรทัดหัวข้อ' ที่เมนูด้านซ้ายดูครับ")
else:
    st.info("👈 กรุณาอัปโหลดไฟล์ CSV ที่เมนูด้านซ้าย")
