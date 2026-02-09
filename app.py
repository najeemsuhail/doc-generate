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
from copy import deepcopy
import re

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

# Helper function to replace text in Word document
def replace_text_in_paragraph(paragraph, replacements):
    """Replace all placeholders in a paragraph, handling split runs"""
    # Get full paragraph text
    full_text = paragraph.text
    
    # Check if any replacement is needed
    needs_replacement = any(key in full_text for key in replacements.keys())
    
    if not needs_replacement:
        return
    
    # Replace all placeholders in the full text
    new_text = full_text
    for key, value in replacements.items():
        new_text = new_text.replace(key, str(value))
    
    # Clear all runs
    for _ in range(len(paragraph.runs)):
        r = paragraph.runs[0]._element
        r.getparent().remove(r)
    
    # Add the new text (preserves paragraph formatting)
    paragraph.text = new_text

def replace_text_in_document(doc, replacements):
    """Replace all placeholders in document with customer data"""
    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_in_paragraph(paragraph, replacements)
                # Also handle cell text directly (some cells might not have paragraphs)
                if cell.text:
                    for key, value in replacements.items():
                        cell.text = cell.text.replace(key, str(value))
    
    return doc

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
        2. **Settings** - Configure company details
        3. **Help** - View documentation and troubleshooting
        4. **About** - App information
        
        ### What This App Does
        - 📁 Upload customer Excel files
        - 📄 Use your own Word template OR text templates
        - 📄 Generate personalized Word documents
        - 📥 Download all letters as ZIP
        
        **Get Started:** Click "📧 Generate Letters" in the menu! ➜
        """)
    
    with col2:
        st.info("""
        **Features:**
        - ✅ Word template support
        - ✅ Text templates
        - ✅ Configurable dates
        - ✅ Bulk generation
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

    # Step 1: Upload Excel
    st.header("📁 Step 1: Upload Excel File")
    uploaded_file = st.file_uploader("Choose your Excel file", type=['xlsx', 'xls'])
    
    if not uploaded_file:
        st.info("👆 Please upload an Excel file to get started")
        st.stop()
    
    df = pd.read_excel(uploaded_file)
    st.success(f"✓ File loaded successfully! ({len(df)} customers found)")
    
    with st.expander("📊 Preview Data", expanded=False):
        st.dataframe(df.head(10), use_container_width=True)
        st.info(f"Total rows: {len(df)}")

    # Step 2: Choose Template Source
    st.header("📋 Step 2: Choose Template Source")
    
    template_source = st.radio(
        "Select template type:",
        ["📄 Upload Word Template (.docx)", "📝 Use Text Template"],
        horizontal=True
    )
    
    template_doc = None
    active_template = ""
    inactive_template = ""
    
    if template_source == "📄 Upload Word Template (.docx)":
        st.info("✓ Upload your own formatted Word document template with placeholders like {CUSTOMER_NAME}, {BILLING_ACCOUNT}, etc.")
        
        template_file = st.file_uploader(
            "Choose Word template file",
            type=['docx'],
            key="template_upload"
        )
        
        if template_file:
            try:
                template_doc = Document(template_file)
                st.success("✓ Template loaded successfully!")
                
                # Extract placeholders from template
                placeholders_found = set()
                for paragraph in template_doc.paragraphs:
                    matches = re.findall(r'\{[^}]+\}', paragraph.text)
                    placeholders_found.update(matches)
                
                for table in template_doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                matches = re.findall(r'\{[^}]+\}', paragraph.text)
                                placeholders_found.update(matches)
                
                if placeholders_found:
                    available_placeholders = sorted(list(placeholders_found))
                    st.info(f"Found placeholders: {', '.join(available_placeholders)}")
                else:
                    st.warning("No placeholders found in template. Use format: {PLACEHOLDER_NAME}")
                    
            except Exception as e:
                st.error(f"Error loading template: {str(e)}")
        else:
            st.warning("Please upload a Word template file (.docx)")
            st.stop()
    
    else:
        st.info("✓ Using text-based templates to generate letters")
        
        # Pre-defined templates
        templates = {
            "payment_reminder": {
                "label": "Payment Reminder",
                "active": """Dear {CUSTOMER_NAME},

We are reaching out regarding your account status and outstanding balance.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

Please review your account and ensure all payments are up to date. If you have any outstanding balance, we request you to settle it at your earliest convenience.

Payment Options:
• Bank transfer
• Check by mail
• Online payment portal
• Digital payment methods

If you have already made a payment or have any questions about your account, please feel free to contact us.

We value your business and look forward to a continued relationship with you.""",
                "inactive": """Dear {CUSTOMER_NAME},

We are writing to inform you regarding your account.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

Please take immediate action on your account. If you have any questions or need assistance, please contact us without delay."""
            },
            
            "collection_notice": {
                "label": "Collection Notice",
                "active": """URGENT: Payment Required

Dear {CUSTOMER_NAME},

Our records indicate an outstanding balance on your account that requires immediate payment.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

Failure to pay may result in suspension of services. Please remit payment within 7 days.

Payment Methods:
• Bank transfer
• Check by mail
• Online payment portal

Contact us immediately if you have any questions.""",
                "inactive": """FINAL NOTICE: Account Status

Dear {CUSTOMER_NAME},

Your account has been marked inactive with an outstanding balance.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

Please settle this amount immediately to reinstate your account."""
            },
            
            "account_status": {
                "label": "Account Status",
                "active": """Dear {CUSTOMER_NAME},

We are writing to confirm the current status of your account.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}
• Account Status: Active

Your account is currently active. Please ensure all outstanding amounts are settled to avoid service interruption.

If you have any questions regarding your account balance or need payment assistance, please contact us.

Thank you for your business.""",
                "inactive": """Dear {CUSTOMER_NAME},

We are writing to inform you that your account is currently inactive.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

If this is due to completion of services, no action is required. However, if you would like to reactivate your account or have outstanding payments, please contact us immediately."""
            },
            
            "service_closure": {
                "label": "Service Closure",
                "active": """Dear {CUSTOMER_NAME},

We are writing to inform you about the status of your account services.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

Please note that your services may be subject to closure if outstanding payments are not settled. We recommend immediate action.

For payment arrangements or further information, please contact our office.

We value your business and would appreciate the opportunity to continue serving you.""",
                "inactive": """Final Notice: Service Closure

Dear {CUSTOMER_NAME},

Your account has been deactivated. There is an outstanding balance that requires settlement.

Account Details:
• Billing Account: {billing_account}
• Department: {department}
• Outstanding Amount: ₹{outstanding:,.2f}

To prevent further action, please settle this amount immediately."""
            }
        }
        
        selected_template = st.selectbox(
            "Choose a template:",
            options=list(templates.keys()),
            format_func=lambda x: templates[x]["label"]
        )
        
        st.info(f"✓ Using **{templates[selected_template]['label']}** template. Customize below if needed.")

        # Letter Template Customization
        st.header("✏️ Step 3: Customize Letter Template")

        template_col1, template_col2 = st.columns(2)

        with template_col1:
            active_template = st.text_area(
                "Active Status Letter",
                value=templates[selected_template]["active"],
                height=300,
                key="active_template"
            )

        with template_col2:
            inactive_template = st.text_area(
                "Inactive Status Letter",
                value=templates[selected_template]["inactive"],
                height=300,
                key="inactive_template"
            )
    
    st.divider()

    # Step 3/4: Generate Letters
    st.header(f"🚀 Step {'3' if template_source.startswith('📄') else '4'}: Generate Letters")

    col1, col2, col3 = st.columns(3)
    
    with col1:
        start_row = st.number_input("Start from row", min_value=1, max_value=len(df), value=1)
    
    with col2:
        end_row = st.number_input("End at row", min_value=1, max_value=len(df), value=len(df))
    
    with col3:
        st.empty()
    
    if st.button("🎯 Generate Letters", key="generate_btn"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        generated_files = []
        
        try:
            for idx, (_, customer) in enumerate(df.iloc[start_row-1:end_row].iterrows()):
                progress = (idx + 1) / (end_row - start_row + 1)
                progress_bar.progress(progress)
                status_text.text(f"Generating letter {idx + 1} of {end_row - start_row + 1}...")
                
                customer_name = customer.get('CUSTOMER NAME', 'Valued Customer')
                outstanding = customer.get('Outstanding amount in Rs', 0)
                billing_account = customer.get('Billing Account', '')
                department = customer.get('Department', '')
                address = customer.get('Address', '')
                status = str(customer.get('Status(Active/Inactive)', 'Active')).lower().strip()
                
                # Create replacement dictionary
                replacements = {
                    # Uppercase versions (template standard)
                    '{CUSTOMER_NAME}': customer_name,
                    '{ADDRESS}': address,
                    '{BILLING_ACCOUNT}': billing_account,
                    '{DEPARTMENT}': department,
                    '{OUTSTANDING_AMOUNT}': f"₹{outstanding:,.2f}",
                    '{STATUS}': status.capitalize(),
                    '{DATE}': letter_date_str,
                    '{COMPANY_NAME}': company_name,
                    '{SENDER_NAME}': sender_name,
                    '{CLOSURE_DATE}': letter_date_str,  # Alternative name
                    '{CLOSURE DATE}': letter_date_str,  # With space
                    # Lowercase versions (backward compatibility)
                    '{customer_name}': customer_name,
                    '{address}': address,
                    '{billing_account}': billing_account,
                    '{department}': department,
                    '{outstanding}': f"{outstanding:,.2f}",
                    '{outstanding:,.2f}': f"{outstanding:,.2f}",
                    '{status}': status,
                    '{date}': letter_date_str,
                    '{company_name}': company_name,
                    '{sender_name}': sender_name,
                }
                
                if template_source == "📄 Upload Word Template (.docx)":
                    # Use Word template
                    doc = deepcopy(template_doc)
                    doc = replace_text_in_document(doc, replacements)
                else:
                    # Create from text template
                    doc = Document()
                    
                    # Header
                    header_para = doc.add_paragraph()
                    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    header_para.add_run(f"{company_name}\n{company_address}\n{company_contact}").font.size = Pt(10)
                    
                    # Date
                    doc.add_paragraph(f"\nDate: {letter_date_str}\n")
                    
                    # Recipient
                    recipient_para = doc.add_paragraph()
                    recipient_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    recipient_para.add_run(f"{customer_name}\n{address}").font.size = Pt(11)
                    
                    # Body
                    if 'inactive' in status:
                        body = inactive_template
                    else:
                        body = active_template
                    
                    body = body.format(
                        CUSTOMER_NAME=customer_name,
                        customer_name=customer_name,
                        billing_account=billing_account,
                        department=department,
                        outstanding=f"{outstanding:,.2f}"
                    )
                    
                    doc.add_paragraph(body)
                    
                    # Closing
                    doc.add_paragraph(f"\nSincerely,\n\n{sender_name}\n{sender_title}\n{company_name}")
                
                # Save document
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
    
    with st.expander("📄 How to Create Word Template"):
        st.markdown("""
        1. Open Microsoft Word
        2. Create your letter template with formatting, logo, etc.
        3. Add placeholders like:
           - `{CUSTOMER_NAME}` - Customer name
           - `{ADDRESS}` - Customer address
           - `{BILLING_ACCOUNT}` - Billing account
           - `{DEPARTMENT}` - Department
           - `{outstanding:,.2f}` - Outstanding amount
           - `{DATE}` - Letter date
           - Any other custom fields
        4. Save as .docx file
        5. Upload in app under "Upload Word Template"
        """)
    
    with st.expander("✏️ How to Use Text Templates"):
        st.markdown("""
        1. Go to **📧 Generate Letters**
        2. Select **📝 Use Text Template**
        3. Choose from 4 pre-built templates:
           - Payment Reminder
           - Collection Notice
           - Account Status
           - Service Closure
        4. Customize the text if needed
        5. Generate letters
        """)
    
    with st.expander("📥 Download & Print"):
        st.markdown("""
        1. Generate letters - downloaded as ZIP
        2. Extract the ZIP file
        3. Open each Word document
        4. Customize if needed
        5. Print for mailing
        """)
    
    with st.expander("❓ Troubleshooting"):
        st.markdown("""
        **Q: Column names don't match?**
        A: Make sure Excel columns match exactly (case-sensitive)
        
        **Q: Placeholders not being replaced?**
        A: Check spelling of placeholders and use curly braces {PLACEHOLDER}
        
        **Q: Can't upload Word template?**
        A: Make sure it's a .docx file (not .doc or .pdf)
        """)

# ABOUT PAGE
elif menu == "ℹ️ About":
    st.header("ℹ️ About This App")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### Customer Letter Generator v2.0
        
        A powerful tool for generating personalized customer letters for bulk mailing.
        
        **Features:**
        - 📄 Word template support (.docx)
        - 📝 4 pre-built text templates
        - 📁 Excel file import
        - 📅 Configurable dates
        - 📥 ZIP download
        - 🎨 Full customization
        
        **Technology Stack:**
        - Python 3.14
        - Streamlit (Web Interface)
        - python-docx (Word generation)
        - pandas (Data processing)
        - openpyxl (Excel reading)
        """)
    
    with col2:
        st.info("""
        **Version:** 2.0
        
        **Date:** Feb 9, 2026
        
        **Features:** Word Templates
        """)

# Footer
st.divider()
st.markdown("""
<div style='text-align: center; color: gray; font-size: 12px;'>
    Customer Letter Generator v2.0 | Ready to deploy
</div>
""", unsafe_allow_html=True)
