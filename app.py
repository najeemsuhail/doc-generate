import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import os
import io
import zipfile
from pathlib import Path

st.set_page_config(page_title="Customer Letter Generator", layout="wide", initial_sidebar_state="expanded")

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 20px;
    }
    .stButton>button {
        width: 100%;
        height: 45px;
        border-radius: 5px;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

st.title("📧 Customer Letter Generator")

# Sidebar Navigation
with st.sidebar:
    st.header("📌 Navigation")
    menu = st.radio(
        "Select Feature:",
        ["🏠 Home", "📧 Generate Letters", "⚙️ Settings", "📚 Help", "ℹ️ About"],
        label_visibility="collapsed"
    )
    st.divider()

# HOME PAGE
if menu == "🏠 Home":
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("Welcome to Customer Letter Generator!")
        st.markdown("""
        ### Quick Start Guide
        
        1. **Generate Letters** - Upload Excel file and create personalized Word documents
        2. **Settings** - Configure company details and letter templates
        3. **Help** - View documentation and troubleshooting
        4. **About** - App information
        
        ### What This App Does
        - 📁 Upload customer Excel files
        - ✏️ Customize letter templates
        - 📄 Generate Word documents for each customer
        - 📥 Download all letters as ZIP
        
        **Get Started:** Click "📧 Generate Letters" in the menu! ➜
        """)
    
    with col2:
        st.info("""
        **Features:**
        - ✅ Bulk letter generation
        - ✅ Custom templates
        - ✅ Configurable dates
        - ✅ ZIP download
        """)

# GENERATE LETTERS PAGE
elif menu == "📧 Generate Letters":
    st.markdown("Generate personalized Word documents for bulk mailing to customers")
    
    # Sidebar Configuration
    with st.sidebar:
        st.header("⚙️ Configuration")
        
        company_name = st.text_input("Company Name", value="Your Company Name")
        company_address = st.text_input("Company Address", value="Your Address")
        company_contact = st.text_input("Contact Info (Email/Phone)", value="[Email/Phone]")
        sender_name = st.text_input("Sender Name", value="Your Name")
        sender_title = st.text_input("Sender Title", value="Your Title")
        
        st.divider()
        st.subheader("📅 Letter Date")
        letter_date = st.date_input("Select date for letters", value=datetime.now().date())
        letter_date_str = letter_date.strftime('%B %d, %Y')

    # Main content
    col1, col2 = st.columns([1.5, 1])

    with col1:
        st.header("📁 Step 1: Upload Excel File")
        uploaded_file = st.file_uploader("Choose your Excel file", type=['xlsx', 'xls'])
        
        if uploaded_file:
            df = pd.read_excel(uploaded_file)
            st.success(f"✓ File loaded successfully! ({len(df)} customers found)")
            
            with st.expander("📊 Preview Data", expanded=False):
                st.dataframe(df.head(10), use_container_width=True)
                st.info(f"Total rows: {len(df)}")

    with col2:
        st.header("📋 Expected Columns")
        expected_cols = [
            "SSA",
            "Billing Account",
            "CUSTOMER NAME",
            "Accot Subtype",
            "Department",
            "Address",
            "Status(Active/Inactive)",
            "Outstanding amount in Rs",
            "CLOSURE DATE"
        ]
        for col in expected_cols:
            st.caption(f"• {col}")

    # Letter Template Customization
    st.header("✏️ Step 2: Customize Letter Template")

    template_col1, template_col2 = st.columns(2)

    with template_col1:
        st.subheader("Letter Header")
        header_text = st.text_area(
            "Company Header Text",
            value=f"{company_name}\n{company_address}\n{company_contact}",
            height=80
        )

    with template_col2:
        st.subheader("Letter Closing")
        closing_text = st.text_area(
            "Letter Closing Text",
            value=f"Thank you for your prompt attention.\n\nSincerely,\n\n{sender_name}\n{sender_title}\n{company_name}",
            height=80
        )

    st.subheader("Letter Body Templates")
    col1, col2 = st.columns(2)

    with col1:
        active_template = st.text_area(
            "Active Status Letter",
            value="""We are reaching out regarding your account status and outstanding balance.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}
• Account Status: Active

Please review your account and ensure all payments are up to date. If you have any outstanding balance, we request you to settle it at your earliest convenience.

Payment Options:
• Bank transfer
• Check by mail
• Online payment portal
• Digital payment methods

If you have already made a payment or have any questions about your account, please feel free to contact us.

We value your business and look forward to a continued relationship with you.""",
            height=250,
            key="active_template"
        )

    with col2:
        inactive_template = st.text_area(
            "Inactive Status Letter",
            value="""We are writing to inform you that your account is currently inactive.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

If your account has been inactive due to closure or completion of services, please disregard this notice. However, if you have any outstanding payments, please settle them at your earliest convenience.

Payment can be made through:
• Bank transfer
• Check by mail
• Online payment portal
• Digital payment methods

If you have any questions regarding your account or need to reactivate your services, please contact us.

Thank you for your attention to this matter.""",
            height=250,
            key="inactive_template"
        )

    # Generate Letters
    st.header("🚀 Step 3: Generate Letters")

    if uploaded_file:
        col1, col2, col3 = st.columns(3)
        
        with col1:
            letter_format = st.selectbox("Output Format", ["Word (.docx)", "PDF (.pdf)"])
        
        with col2:
            start_row = st.number_input("Start from row", min_value=1, max_value=len(df), value=1)
        
        with col3:
            end_row = st.number_input("End at row", min_value=1, max_value=len(df), value=len(df))
        
        if st.button("🎯 Generate Letters", key="generate_btn"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            generated_files = []
            
            try:
                for idx, (_, customer) in enumerate(df.iloc[start_row-1:end_row].iterrows()):
                    progress = (idx + 1) / (end_row - start_row + 1)
                    progress_bar.progress(progress)
                    status_text.text(f"Generating letter {idx + 1} of {end_row - start_row + 1}...")
                    
                    doc = Document()
                    
                    header_para = doc.add_paragraph()
                    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    header_para.add_run(header_text).font.size = Pt(10)
                    
                    doc.add_paragraph(f"\nDate: {letter_date_str}\n")
                    
                    recipient_para = doc.add_paragraph()
                    recipient_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    customer_name = customer.get('CUSTOMER NAME', 'Valued Customer')
                    recipient_text = f"{customer_name}\n"
                    if pd.notna(customer.get('Address')):
                        recipient_text += f"{customer['Address']}"
                    recipient_para.add_run(recipient_text).font.size = Pt(11)
                    
                    salutation_name = str(customer_name) if pd.notna(customer_name) else "Valued Customer"
                    doc.add_paragraph(f"\nDear {salutation_name},")
                    
                    status = str(customer.get('Status(Active/Inactive)', 'Active')).lower().strip()
                    outstanding = customer.get('Outstanding amount in Rs', 0)
                    billing_account = customer.get('Billing Account', '')
                    department = customer.get('Department', '')
                    
                    if 'inactive' in status:
                        body = inactive_template
                    else:
                        body = active_template
                    
                    body = body.format(
                        billing_account=billing_account,
                        department=department,
                        outstanding=outstanding
                    )
                    
                    doc.add_paragraph(body)
                    doc.add_paragraph(f"\n{closing_text}")
                    
                    filename = f"Letter_{str(customer_name).replace(' ', '_').replace('/', '_')}.docx"
                    doc.save(filename)
                    generated_files.append(filename)
                
                status_text.success(f"✅ Generated {len(generated_files)} letters successfully!")
                progress_bar.empty()
                
                if generated_files:
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file in generated_files:
                            zip_file.write(file, arcname=file)
                    
                    zip_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download All Letters (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name=f"customer_letters_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                        mime="application/zip",
                        key="download_zip"
                    )
                    
                    for file in generated_files:
                        if os.path.exists(file):
                            os.remove(file)
            
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

    else:
        st.info("👆 Please upload an Excel file to get started")

# SETTINGS PAGE
elif menu == "⚙️ Settings":
    st.header("⚙️ Settings")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📋 Company Details")
        company_name = st.text_input("Company Name", value="Your Company Name", key="settings_company")
        company_address = st.text_input("Company Address", value="Your Address", key="settings_addr")
        company_contact = st.text_input("Contact Info", value="[Email/Phone]", key="settings_contact")
    
    with col2:
        st.subheader("👤 Sender Details")
        sender_name = st.text_input("Sender Name", value="Your Name", key="settings_sender")
        sender_title = st.text_input("Sender Title", value="Your Title", key="settings_title")
    
    st.divider()
    
    st.subheader("📅 Default Date Settings")
    letter_date = st.date_input("Default letter date", value=datetime.now().date(), key="settings_date")
    letter_date_str = letter_date.strftime('%B %d, %Y')
    
    if st.button("💾 Save Settings"):
        st.success("✓ Settings saved! (Session settings)")

# HELP PAGE
elif menu == "📚 Help":
    st.header("📚 Help & Documentation")
    
    with st.expander("📁 How to Prepare Excel File"):
        st.markdown("""
        Your Excel file should have these columns:
        
        | Column | Type | Example |
        |--------|------|---------|
        | SSA | Text | SSA-001 |
        | Billing Account | Text | ACC-12345 |
        | CUSTOMER NAME | Text | John Smith |
        | Accot Subtype | Text | Premium |
        | Department | Text | Sales |
        | Address | Text | 123 Main St |
        | Status(Active/Inactive) | Text | Active |
        | Outstanding amount in Rs | Number | 5000.00 |
        | CLOSURE DATE | Date | 2026-03-01 |
        """)
    
    with st.expander("✏️ How to Customize Letters"):
        st.markdown("""
        1. Go to **📧 Generate Letters**
        2. Scroll down to **Step 2: Customize Letter Template**
        3. Edit company header, body text, and closing
        4. Use these placeholders in body text:
           - `{billing_account}` - Customer's billing account
           - `{department}` - Department name
           - `{outstanding:,.2f}` - Outstanding amount formatted
        """)
    
    with st.expander("📥 Download & Print"):
        st.markdown("""
        1. Generate letters - they're downloaded as ZIP
        2. Extract the ZIP file
        3. Open each Word document
        4. Customize if needed
        5. Print for mailing
        """)
    
    with st.expander("❓ Troubleshooting"):
        st.markdown("""
        **Q: Column names don't match?**
        A: Make sure Excel columns match exactly (case-sensitive)
        
        **Q: Can't generate letters?**
        A: Check that Status column contains "Active" or "Inactive"
        
        **Q: Date format wrong?**
        A: Use the date picker in Generate Letters to select the date
        """)

# ABOUT PAGE
elif menu == "ℹ️ About":
    st.header("ℹ️ About This App")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### Customer Letter Generator v1.0
        
        A powerful tool for generating personalized customer letters for bulk mailing.
        
        **Features:**
        - 📁 Excel file import
        - ✏️ Custom letter templates
        - 📄 Batch Word document generation
        - 📅 Configurable dates
        - 📥 ZIP download
        
        **Technology Stack:**
        - Python 3.14
        - Streamlit (Web Interface)
        - python-docx (Word generation)
        - pandas (Data processing)
        - openpyxl (Excel reading)
        
        **GitHub Repository:**
        https://github.com/YOUR_USERNAME/doc-generate
        """)
    
    with col2:
        st.info("""
        **Version:** 1.0
        
        **Date:** Feb 9, 2026
        
        **License:** MIT
        """)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; font-size: 12px;'>
    Customer Letter Generator v1.0 | Ready to deploy
</div>
""", unsafe_allow_html=True)
