import streamlit as st
import pandas as pd
import re
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formataddr
from email import encoders

# --- EMAIL CONFIGURATION ---
SENDER_EMAIL = "atharvaujoshi@gmail.com"
SENDER_NAME = "Spydarr Package Reporter" 
APP_PASSWORD = "nybl zsnx zvdw edqr"  

def send_email(recipient_email, excel_data, filename):
    try:
        recipient_name = recipient_email.split('@')[0].replace('.', ' ').title()
        msg = MIMEMultipart()
        msg['From'] = formataddr((SENDER_NAME, SENDER_EMAIL))
        msg['To'] = recipient_email
        msg['Subject'] = "Property Package Report"
        body = f"Dear {recipient_name},\n\nPlease find the attached updated Report with Package calculations.\n\nRegards,\nAtharva Joshi"
        msg.attach(MIMEText(body, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(excel_data)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename={filename}")
        msg.attach(part)
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending email: {e}")
        return False

def extract_lower_carpet(value):
    if pd.isna(value): return 0
    numbers = re.findall(r"[-+]?\d*\.\d+|\d+", str(value))
    return float(numbers[0]) if numbers else 0

# --- UI SETUP ---
st.set_page_config(page_title="Package Calculator", layout="wide")
st.title("🏙️ Package Calculation & Reporter")

uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        target_sheet = next((s for s in xls.sheet_names if s.lower() == 'report'), None)
        
        if not target_sheet:
            st.error("Could not find a sheet named 'Report' or 'report'.")
        else:
            df = pd.read_excel(xls, sheet_name=target_sheet)
            cols = df.columns.tolist()
            
            # Find specific columns regardless of exact casing
            carpet_col = next((c for c in cols if "carpet area(sq.ft)" in c.lower()), None)
            apr_col = next((c for c in cols if "average of apr" in c.lower()), None)
            count_col = next((c for c in cols if "count of property" in c.lower()), None)

            if carpet_col and apr_col:
                # Calculation Logic
                lower_carpet = df[carpet_col].apply(extract_lower_carpet)
                # Formula: Lower Carpet * 1.4 * 1.12 * Average of APR
                df['Package'] = (lower_carpet * 1.4 * 1.12 * df[apr_col]).round(0)
                
                # Column Reordering
                if count_col:
                    current_cols = df.columns.tolist()
                    current_cols.remove('Package')
                    count_idx = current_cols.index(count_col)
                    current_cols.insert(count_idx, 'Package')
                    df = df[current_cols]

                st.success("Calculations complete!")
                st.dataframe(df.head(10))

                # Prepare Excel for Download/Email
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Report')
                file_content = output.getvalue()

                # --- SIDEBAR EMAIL LOGIC ---
                st.sidebar.header("📧 Email Report")
                recipient = st.sidebar.text_input("Recipient Name", placeholder="firstname.lastname")
                
                if st.sidebar.button("Send to Email") and recipient:
                    full_email = f"{recipient.strip().lower()}@beyondwalls.com"
                    with st.spinner(f'Sending to {full_email}...'):
                        if send_email(full_email, file_content, "Updated_Package_Report.xlsx"):
                            st.sidebar.success(f"Report sent to {full_email}")
                
              
            else:
                st.error("Required columns ('Carpet Area(SQ.FT)' or 'Average of APR') not found.")

    except Exception as e:
        st.error(f"Error: {e}")
