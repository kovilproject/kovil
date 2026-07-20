import streamlit as st
import pandas as pd
from datetime import datetime
import io
import os
import sqlite3
from indic_transliteration import sanscript
from indic_transliteration.sanscript import transliterate
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ---------------------------------------------------------
# 1. PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="கோவில் கணக்கு மேலாண்மை", layout="wide", page_icon="🛕")

# ---------------------------------------------------------
# 2. SQLITE DATABASE SETUP
# ---------------------------------------------------------
DB_NAME = "kovil_kanakku.db"


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    # Receipts table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS receipts (
            receipt_no INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT,
            category TEXT,
            name TEXT,
            city TEXT,
            amount REAL
        )
    ''')
    # Users table for Login
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            username TEXT PRIMARY KEY,
            password TEXT
        )
    ''')
    # Create default admin if not exists
    cursor.execute("SELECT * FROM users WHERE username = 'admin'")
    if not cursor.fetchone():
        cursor.execute("INSERT INTO users (username, password) VALUES (?, ?)", ('admin', 'kovil123'))

    conn.commit()
    conn.close()


init_db()

# ---------------------------------------------------------
# 3. AUTHENTICATION (LOGIN SYSTEM)
# ---------------------------------------------------------
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False


def check_login(username, password):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
    user = cursor.fetchone()
    conn.close()
    return user is not None


def update_password(username, new_password):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("UPDATE users SET password = ? WHERE username = ?", (new_password, username))
    conn.commit()
    conn.close()


if not st.session_state['logged_in']:
    st.markdown("<h1 style='text-align: center;'>🛕 ஸ்ரீ மகா மாரியம்மன் கோவில்</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align: center;'>கணக்கு மேலாண்மை - உள்நுழைவு (Login)</h3>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("🔐 தொடர லாகின் செய்யவும்")
        user_input = st.text_input("பயனர் பெயர் (Username)")
        pass_input = st.text_input("கடவுச்சொல் (Password)", type="password")

        if st.button("🔓 உள்நுழைக (Login)", use_container_width=True):
            if check_login(user_input, pass_input):
                st.session_state['logged_in'] = True
                st.session_state['username'] = user_input
                st.success("✅ வெற்றிகரமாக உள்நுழைந்துவிட்டீர்கள்!")
                st.rerun()
            else:
                st.error("❌ தவறான பயனர் பெயர் அல்லது கடவுச்சொல்!")
    st.stop()

# ---------------------------------------------------------
# 4. SIDEBAR & LOGOUT
# ---------------------------------------------------------
st.sidebar.title(f"👤 வரவேற்கிறோம், {st.session_state['username']}!")
if st.sidebar.button("🚪 வெளியேறு (Logout)"):
    st.session_state['logged_in'] = False
    st.rerun()

st.sidebar.markdown("---")
st.sidebar.subheader("🔑 பாஸ்வேர்ட் மாற்ற")
new_pass = st.sidebar.text_input("புதிய கடவுச்சொல்", type="password")
if st.sidebar.button("பாஸ்வேர்ட் மாற்று"):
    if new_pass:
        update_password(st.session_state['username'], new_pass)
        st.sidebar.success("✅ பாஸ்வேர்ட் மாற்றப்பட்டது!")
    else:
        st.sidebar.warning("⚠️ புதிய பாஸ்வேர்ட் டைப் செய்யவும்!")


# ---------------------------------------------------------
# 5. ACCURATE TANGLISH TO TAMIL CONVERTER
# ---------------------------------------------------------
def to_tamil(text):
    if not text:
        return ""

    special_dict = {
        "muthu": "முத்து", "ramesh": "ரமேஷ்", "suresh": "சுரேஷ்", "kumar": "குமார்",
        "raja": "ராஜா", "karthik": "கார்த்திக்", "murugan": "முருகன்", "selvam": "செல்வம்",
        "madurai": "மதுரை", "melur": "மேலூர்", "chennai": "சென்னை", "kovil": "கோவில்"
    }

    clean_text = text.strip().lower()
    if clean_text in special_dict:
        return special_dict[clean_text]

    try:
        converted = transliterate(text, sanscript.ITRANS, sanscript.TAMIL)
        return converted
    except Exception:
        return text


# ---------------------------------------------------------
# 6. TAMIL FONT HELPER
# ---------------------------------------------------------
def get_pdf_font():
    font_name = 'Helvetica'
    font_paths = [
        "C:\\Windows\\Fonts\\Nirmala.ttf",
        "C:\\Windows\\Fonts\\latha.ttf",
        "Bamini.ttf",
        "NotoSansTamil-Regular.ttf"
    ]
    for path in font_paths:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont('TamilFont', path))
                return 'TamilFont'
            except:
                continue
    return font_name


# ---------------------------------------------------------
# 7. PDF GENERATION
# ---------------------------------------------------------
def generate_receipt_pdf(title, name, city, amount, receipt_no):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    font_name = get_pdf_font()

    c.setLineWidth(2)
    c.rect(20, 20, width - 40, height - 40)

    c.setFont(font_name, 20)
    c.drawCentredString(width / 2, height - 80, "ஸ்ரீ மகா மாரியம்மன் கோவில்")
    c.setFont(font_name, 12)
    c.drawCentredString(width / 2, height - 105, "நற்பணி மன்றம் மற்றும் தளபதி தளபதிகள்")
    c.drawCentredString(width / 2, height - 125, "மேலூர் மேல வீதி, மதுரை - 625020.")
    c.line(40, height - 145, width - 40, height - 145)

    current_date = datetime.now().strftime("%d-%m-%Y")
    c.setFont(font_name, 12)
    c.drawString(50, height - 175, f"ரசீது எண் :  {receipt_no}")
    c.drawString(width - 180, height - 175, f"தேதி :  {current_date}")

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 210, title)

    c.rect(40, height - 420, width - 80, 180)
    c.setFont(font_name, 13)
    c.drawString(60, height - 260, f"பக்தர் பெயர் (Name)     :   {name}")
    c.drawString(60, height - 300, f"ஊர் / பகுதி (City)        :   {city}")
    c.drawString(60, height - 380, f"தொகை (Amount)           :   Rs. {amount}/-")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


def generate_full_report_pdf(df):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=landscape(A4))
    width, height = landscape(A4)
    font_name = get_pdf_font()

    c.setFont(font_name, 16)
    c.drawCentredString(width / 2, height - 40, "ஸ்ரீ மகா மாரியம்மன் கோவில் - கணக்கு அறிக்கை (Report)")
    c.setFont(font_name, 10)
    c.drawCentredString(width / 2, height - 55, f"உருவாக்கப்பட்ட தேதி: {datetime.now().strftime('%d-%m-%Y %I:%M %p')}")
    c.line(30, height - 65, width - 30, height - 65)

    y = height - 90
    c.setFont(font_name, 11)
    c.drawString(40, y, "ரசீது எண்")
    c.drawString(120, y, "தேதி")
    c.drawString(220, y, "வகை")
    c.drawString(340, y, "பெயர்")
    c.drawString(550, y, "ஊர்")
    c.drawString(720, y, "தொகை (Rs.)")

    c.line(30, y - 5, width - 30, y - 5)
    y -= 25

    total_amt = 0
    c.setFont(font_name, 10)
    for _, row in df.iterrows():
        if y < 50:
            c.showPage()
            y = height - 50
            c.setFont(font_name, 10)

        c.drawString(40, y, str(row['ரசீது எண்']))
        c.drawString(120, y, str(row['தேதி']))
        c.drawString(220, y, str(row['வகை']))
        c.drawString(340, y, str(row['பெயர்']))
        c.drawString(550, y, str(row['ஊர்']))
        c.drawString(720, y, f"{row['தொகை (₹)']:,.2f}")

        total_amt += row['தொகை (₹)']
        y -= 20

    c.line(30, y + 10, width - 30, y + 10)
    c.setFont(font_name, 11)
    c.drawString(550, y - 10, "மொத்த வசூல்:")
    c.drawString(720, y - 10, f"Rs. {total_amt:,.2f}")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------
# 8. STREAMLIT MAIN APP UI
# ---------------------------------------------------------
st.title("🛕 ஸ்ரீ மகா மாரியம்மன் கோவில் - ரசீது பதிவேடு")

tab1, tab2, tab3 = st.tabs(
    ["📝 புதிய ரசீது (Entry)", "📊 கணக்கு அறிக்கைகள் (Reports)", "🗑️ என்ட்ரி நீக்குதல் (Delete Entry)"])

# --- TAB 1: ENTRY ---
with tab1:
    st.subheader("புதிய ரசீது பதிவு")
    col1, col2 = st.columns(2)

    with col1:
        category = st.selectbox("வகை", ["வைப்பு நிதி", "காணிக்கை"])
        name_eng = st.text_input("பக்தர் பெயர் (English / Tanglish-ல் டைப் செய்யவும்):",
                                 placeholder="எ.கா: muthu / ramesh / karthik")
        name_tam_auto = to_tamil(name_eng) if name_eng else ""
        name_final = st.text_input("தமிழ் பெயர் (தேவைப்பட்டால் மாற்றிக் கொள்ளலாம்):", value=name_tam_auto)

    with col2:
        city_eng = st.text_input("ஊர் / பகுதி (English / Tanglish-ல் டைப் செய்யவும்):",
                                 placeholder="எ.கா: madurai / melur")
        city_tam_auto = to_tamil(city_eng) if city_eng else ""
        city_final = st.text_input("தமிழ் ஊர் (தேவைப்பட்டால் மாற்றிக் கொள்ளலாம்):", value=city_tam_auto)
        amount = st.number_input("தொகை (₹)", min_value=1.0, step=10.0)

    st.markdown("---")

    if st.button("💾 சேமி & PDF உருவாக்க"):
        if not name_final or amount <= 0:
            st.warning("⚠️ தயவுசெய்து பெயர் மற்றும் தொகையை நிரப்பவும்!")
        else:
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                current_date = datetime.now().strftime("%d-%m-%Y")
                cursor.execute('''
                    INSERT INTO receipts (date, category, name, city, amount)
                    VALUES (?, ?, ?, ?, ?)
                ''', (current_date, category, name_final, city_final, amount))
                conn.commit()
                receipt_no = cursor.lastrowid
                conn.close()

                st.success(f"✅ வெற்றிகரமாக சேமிக்கப்பட்டது! ரசீது எண்: {receipt_no}")
                pdf_bytes = generate_receipt_pdf(category, name_final, city_final, amount, receipt_no)
                st.download_button(
                    label="📄 தனி நபர் PDF ரசீது டவுன்லோட் செய்",
                    data=pdf_bytes,
                    file_name=f"Receipt_{receipt_no}.pdf",
                    mime="application/pdf"
                )
            except Exception as e:
                st.error(f"❌ சேமிப்பதில் எர்ரர்: {e}")

# --- TAB 2: REPORTS ---
with tab2:
    st.subheader("📊 முன்னேறிய கணக்கு அறிக்கைகள் (Advanced Reports)")
    try:
        conn = sqlite3.connect(DB_NAME)
        df = pd.read_sql_query(
            "SELECT receipt_no as 'ரசீது எண்', date as 'தேதி', category as 'வகை', name as 'பெயர்', city as 'ஊர்', amount as 'தொகை (₹)' FROM receipts ORDER BY receipt_no DESC",
            conn)
        conn.close()

        if not df.empty:
            st.markdown("#### 🔍 ஃபில்டர் வசதிகள்")
            col_f1, col_f2, col_f3 = st.columns(3)

            with col_f1:
                search_term = st.text_input("🔎 பெயர் அல்லது ஊரை வைத்து தேட:", placeholder="எ.கா: முத்து / மதுரை")
            with col_f2:
                cat_filter = st.selectbox("வகை வாரியாக பிரிக்க:", ["அனைத்தும் (All)", "காணிக்கை", "வைப்பு நிதி"])
            with col_f3:
                date_filter = st.date_input("தேதி வாரியாக பார்க்க (Range):", [])

            filtered_df = df.copy()
            if search_term:
                filtered_df = filtered_df[
                    filtered_df['பெயர்'].astype(str).str.contains(search_term, case=False, na=False) |
                    filtered_df['ஊர்'].astype(str).str.contains(search_term, case=False, na=False)
                    ]
            if cat_filter != "அனைத்தும் (All)":
                filtered_df = filtered_df[filtered_df['வகை'] == cat_filter]
            if len(date_filter) == 2:
                start_date, end_date = date_filter
                filtered_df['temp_date'] = pd.to_datetime(filtered_df['தேதி'], format='%d-%m-%Y')
                filtered_df = filtered_df[
                    (filtered_df['temp_date'] >= pd.to_datetime(start_date)) &
                    (filtered_df['temp_date'] <= pd.to_datetime(end_date))
                    ].drop(columns=['temp_date'])

            st.markdown("---")
            st.markdown("#### 💰 கணக்கு சுருக்கம் (Summary)")
            m1, m2, m3, m4 = st.columns(4)

            total_collection = filtered_df["தொகை (₹)"].sum()
            kanikkai_tot = filtered_df[filtered_df['வகை'] == 'காணிக்கை']["தொகை (₹)"].sum()
            vaipu_tot = filtered_df[filtered_df['வகை'] == 'வைப்பு நிதி']["தொகை (₹)"].sum()
            total_entries = len(filtered_df)

            m1.metric("மொத்த வசூல்", f"₹ {total_collection:,.2f}")
            m2.metric("காணிக்கை மொத்தம்", f"₹ {kanikkai_tot:,.2f}")
            m3.metric("வைப்பு நிதி மொத்தம்", f"₹ {vaipu_tot:,.2f}")
            m4.metric("மொத்த பதிவுகள்", f"{total_entries} எண்கள்")

            st.markdown("---")
            st.dataframe(filtered_df, use_container_width=True)

            st.markdown("#### 📥 அறிக்கையை பதிவிறக்கம் செய்ய / பிரிண்ட் எடுக்க")
            col_p1, col_p2, col_p3 = st.columns(3)

            with col_p1:
                pdf_report_bytes = generate_full_report_pdf(filtered_df)
                st.download_button(
                    label="🖨️ முழு PDF Report (Print-Ready)",
                    data=pdf_report_bytes,
                    file_name=f"Kovil_Full_Report_{datetime.now().strftime('%d_%m_%Y')}.pdf",
                    mime="application/pdf"
                )
            with col_p2:
                buffer_excel = io.BytesIO()
                with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                    filtered_df.to_excel(writer, index=False, sheet_name='Receipts_Report')
                st.download_button(
                    label="📥 Excel Report டவுன்லோட்",
                    data=buffer_excel.getvalue(),
                    file_name=f"Kovil_Report_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_p3:
                csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📄 CSV File டவுன்லோட்",
                    data=csv,
                    file_name=f"Kovil_Report_{datetime.now().strftime('%d_%m_%Y')}.csv",
                    mime="text/csv"
                )
        else:
            st.info("இன்னும் பதிவுகள் எதுவும் செய்யப்படவில்லை.")
    except Exception as e:
        st.error(f"டேட்டாவை வாசிக்க முடியவில்லை: {e}")

# --- TAB 3: DELETE ---
with tab3:
    st.subheader("🗑️ தவறான பதிவை நீக்குதல்")
    conn = sqlite3.connect(DB_NAME)
    df_del = pd.read_sql_query("SELECT receipt_no, name, amount FROM receipts ORDER BY receipt_no DESC", conn)
    conn.close()

    if not df_del.empty:
        receipt_list = [f"ரசீது எண்: {row['receipt_no']} - {row['name']} (₹{row['amount']})" for _, row in
                        df_del.iterrows()]
        selected_item = st.selectbox("நீக்க வேண்டிய ரசீதைத் தேர்ந்தெடுக்கவும்:", receipt_list)
        target_id = int(selected_item.split("-")[0].replace("ரசீது எண்:", "").strip())

        if st.button("❌ இந்த ரசீதை நீக்கு (Delete Entry)"):
            try:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("DELETE FROM receipts WHERE receipt_no = ?", (target_id,))
                conn.commit()
                conn.close()
                st.success(f"✅ ரசீது எண் {target_id} நீக்கப்பட்டது!")
                st.rerun()
            except Exception as e:
                st.error(f"நீக்குவதில் எர்ரர்: {e}")
    else:
        st.info("நீக்குவதற்கு பதிவுகள் இல்லை.")