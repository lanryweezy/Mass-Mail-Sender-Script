import streamlit as st

def main():
    st.set_page_config(layout="wide", page_title="Mass Email Sender")

    st.title("📤 Mass Email Sender Deluxe")

    # Sidebar for navigation/mode selection
    st.sidebar.title("Navigation")
    app_mode = st.sidebar.radio(
        "Choose a section:",
        ("🚀 Send Emails", "📄 Landing Page")
    )

    if app_mode == "🚀 Send Emails":
        run_sender_app()
    elif app_mode == "📄 Landing Page":
        show_landing_page()

def run_sender_app():
    st.header("Configure, Compose, and Send Your Emails")

    # 1. Configuration Section
    with st.expander("⚙️ Step 1: Email Sender Configuration", expanded=True):
        st.subheader("Sender Email Account Details")

        # Session state to store configuration
        if 'config' not in st.session_state:
            st.session_state.config = {
                "sender_email": "",
                "email_password": "",
                "smtp_server": "",
                "smtp_port": 587,
                "smtp_security": "TLS"
            }

        st.session_state.config['sender_email'] = st.text_input(
            "Your Email Address (e.g., yourname@gmail.com)",
            value=st.session_state.config.get('sender_email', '')
        )
        st.session_state.config['email_password'] = st.text_input(
            "App Password / Email Account Password",
            type="password",
            value=st.session_state.config.get('email_password', '')
        )
        st.caption("Note: For Gmail, you might need to generate an 'App Password'. For other providers, use your regular email password.")

        st.subheader("SMTP Server Details")

        # Pre-fill common SMTP settings
        common_smtp = {
            "Gmail": ("smtp.gmail.com", 587, "TLS"),
            "Outlook/Hotmail": ("smtp.office365.com", 587, "TLS"),
            "Yahoo": ("smtp.mail.yahoo.com", 587, "TLS"), # Yahoo can also use port 465 with SSL
            "Other": ("", 587, "TLS")
        }

        selected_provider = st.selectbox(
            "Select Email Provider (for common settings) or choose 'Other'",
            list(common_smtp.keys())
        )

        if selected_provider != "Other":
            # When a common provider is selected, update session_state.config directly
            # The widgets will then read these values.
            # This ensures that changing the provider updates the config.
            st.session_state.config['smtp_server'] = common_smtp[selected_provider][0]
            st.session_state.config['smtp_port'] = common_smtp[selected_provider][1]
            st.session_state.config['smtp_security'] = common_smtp[selected_provider][2]
        # For "Other", we let the user input directly into the config fields below.
        # If the user switches from a provider to "Other", the last provider's settings will persist
        # which is reasonable behavior, allowing them to customize from a template.

        st.session_state.config['smtp_server'] = st.text_input(
            "SMTP Server Address",
            value=st.session_state.config.get('smtp_server', ''),
            # No specific key needed here if we are directly modifying st.session_state.config before this line
            # However, for explicit control and to ensure Streamlit's rerun logic works as expected with selectbox changes,
            # it's better to ensure these are treated as distinct fields that get their value from the updated config.
            key='smtp_server_input'
        )
        st.session_state.config['smtp_port'] = st.number_input(
            "SMTP Port",
            min_value=1, max_value=65535,
            value=st.session_state.config.get('smtp_port', 587),
            key='smtp_port_input'
        )

        security_options = ("TLS", "SSL", "None")
        current_security = st.session_state.config.get('smtp_security', "TLS")
        default_security_index = security_options.index(current_security) if current_security in security_options else 0

        st.session_state.config['smtp_security'] = st.selectbox(
            "SMTP Security",
            security_options,
            index=default_security_index,
            key='smtp_security_input'
        )

        st.info("Ensure your email account allows SMTP access. For Gmail, you may need to enable 'Less secure app access' or use an 'App Password'.")

        # Display current config for debugging (optional, can be removed)
        # st.write("Current Config (live):", st.session_state.config)

import pandas as pd
import json
# openpyxl will be needed for pd.read_excel to read .xlsx files
# No direct import needed here, but it's a dependency for pandas.

import re

# ... (other imports and code remain the same) ...

# Helper function for basic email validation
def is_valid_email(email: str) -> bool:
    if not email or not isinstance(email, str):
        return False
    # Basic regex for email validation - can be improved for stricter validation if needed
    # This regex checks for a common pattern: something@something.something
    pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    return bool(re.match(pattern, email))

# Helper function to parse uploaded files
def load_data(uploaded_file):
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        elif uploaded_file.name.endswith(('.xls', '.xlsx')):
            df = pd.read_excel(uploaded_file)
        else:
            st.error("Unsupported file type. Please upload a CSV or Excel file.")
            return None
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def run_sender_app():
    st.header("Configure, Compose, and Send Your Emails")

    # Initialize session state for data if not already present
    if 'recipient_df' not in st.session_state:
        st.session_state.recipient_df = None # Will hold the DataFrame of recipients
    if 'manual_data' not in st.session_state:
        # For manual entry, a list of dictionaries
        st.session_state.manual_data = [{"Email": "example@test.com", "Name": "Test User"}]

    # ... (Configuration section remains the same) ...
    # 1. Configuration Section
    with st.expander("⚙️ Step 1: Email Sender Configuration", expanded=True):
        st.subheader("Sender Email Account Details")

        # Session state to store configuration
        if 'config' not in st.session_state:
            st.session_state.config = {
                "sender_email": "",
                "email_password": "",
                "smtp_server": "",
                "smtp_port": 587,
                "smtp_security": "TLS"
            }

        st.session_state.config['sender_email'] = st.text_input(
            "Your Email Address (e.g., yourname@gmail.com)",
            value=st.session_state.config.get('sender_email', '')
        )
        st.session_state.config['email_password'] = st.text_input(
            "App Password / Email Account Password",
            type="password",
            value=st.session_state.config.get('email_password', '')
        )
        st.caption("Note: For Gmail, you might need to generate an 'App Password'. For other providers, use your regular email password.")

        st.subheader("SMTP Server Details")

        common_smtp = {
            "Gmail": ("smtp.gmail.com", 587, "TLS"),
            "Outlook/Hotmail": ("smtp.office365.com", 587, "TLS"),
            "Yahoo": ("smtp.mail.yahoo.com", 587, "TLS"),
            "Other": ("", 587, "TLS") # Default for 'Other'
        }

        # Initialize default provider if not set
        if 'selected_provider' not in st.session_state:
            st.session_state.selected_provider = "Gmail" # Default to Gmail or any other

        # Update selected_provider in session_state when selectbox changes
        selected_provider = st.selectbox(
            "Select Email Provider (for common settings) or choose 'Other'",
            list(common_smtp.keys()),
            index=list(common_smtp.keys()).index(st.session_state.selected_provider), # Use session state for index
            key='provider_selectbox' # Unique key for the selectbox
        )
        st.session_state.selected_provider = selected_provider # Store the selection

        # Update config based on selected provider
        if selected_provider != "Other":
            st.session_state.config['smtp_server'] = common_smtp[selected_provider][0]
            st.session_state.config['smtp_port'] = common_smtp[selected_provider][1]
            st.session_state.config['smtp_security'] = common_smtp[selected_provider][2]
        # If "Other", the config retains its current values, which the user can then edit in the fields below.
        # If switching from a provider to "Other", the provider's values will be pre-filled from the previous step.

        st.session_state.config['smtp_server'] = st.text_input(
            "SMTP Server Address",
            value=st.session_state.config.get('smtp_server', ""), # Default to empty if somehow not set
            key='smtp_server_input_field'
        )
        st.session_state.config['smtp_port'] = st.number_input(
            "SMTP Port",
            min_value=1, max_value=65535,
            value=int(st.session_state.config.get('smtp_port', 587)), # Ensure port is int
            key='smtp_port_input_field'
        )

        security_options = ("TLS", "SSL", "None")
        current_security_value = st.session_state.config.get('smtp_security', "TLS")
        # Ensure current_security_value is one of the options to prevent error in .index()
        if current_security_value not in security_options:
            current_security_value = "TLS" # Fallback
        default_security_index = security_options.index(current_security_value)

        st.session_state.config['smtp_security'] = st.selectbox(
            "SMTP Security",
            security_options,
            index=default_security_index,
            key='smtp_security_input_field'
        )

        st.info("Ensure your email account allows SMTP access. For Gmail, you may need to enable 'Less secure app access' or use an 'App Password'.")

        if st.button("🔄 Reset Configuration to Defaults", key="reset_config_button"):
            st.session_state.config = {
                "sender_email": "",
                "email_password": "",
                "smtp_server": common_smtp["Gmail"][0], # Default to Gmail settings
                "smtp_port": common_smtp["Gmail"][1],
                "smtp_security": common_smtp["Gmail"][2]
            }
            st.session_state.selected_provider = "Gmail" # Reset provider selection
            st.success("Configuration reset to defaults (Gmail).")
            st.rerun()

    # 2. Data Input Section
    with st.expander("📊 Step 2: Upload or Create Recipient Data", expanded=True): # Expanded by default

        data_input_method = st.radio(
            "Choose data input method:",
            ("Upload File", "Manual Entry"),
            key="data_input_method_radio"
        )

        if data_input_method == "Upload File":
            st.subheader("Upload Data File")
            uploaded_file = st.file_uploader(
                "Upload CSV or Excel file",
                type=['csv', 'xlsx', 'xls'],
                key="file_uploader"
            )
            if uploaded_file is not None:
                df = load_data(uploaded_file)
                if df is not None:
                    st.session_state.recipient_df = df
                    st.success(f"Successfully loaded {uploaded_file.name}")
                    # Clear manual data if file is uploaded
                    st.session_state.manual_data = []

            # Sample CSV Download Button
            sample_csv_data = "Email,Name,Company,Birthday,CustomField1\n" \
                              "johndoe@example.com,John Doe,Example Inc.,1990-05-15,ValueA\n" \
                              "janesmith@example.com,Jane Smith,Another Corp,1985-11-22,ValueB\n" \
                              "contact@company.com,Info Desk,Company LLC,,ValueC"
            st.download_button(
                label="📥 Download Sample CSV Template",
                data=sample_csv_data,
                file_name="sample_recipient_template.csv",
                mime="text/csv",
                key="download_sample_csv_button"
            )

            if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
                 st.caption("Uploaded Data Preview (first 5 rows):")
                 st.dataframe(st.session_state.recipient_df.head())


        elif data_input_method == "Manual Entry":
            st.subheader("Create or Edit Data Manually")

            # Use st.data_editor for a basic table editing experience
            # Initialize if empty or if switching from file upload
            if not st.session_state.manual_data or st.session_state.recipient_df is not None:
                 if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
                     # If switching from file upload with data, offer to convert it to manual editing
                     if st.button("Edit Uploaded Data Manually"):
                         st.session_state.manual_data = st.session_state.recipient_df.to_dict('records')
                         st.session_state.recipient_df = None # Clear uploaded df
                 elif not st.session_state.manual_data : # Initialize with a default row if empty
                     st.session_state.manual_data = [{"Email": "contact@example.com", "Name": "New Contact"}]


            if st.session_state.manual_data:
                edited_data = st.data_editor(
                    pd.DataFrame(st.session_state.manual_data), # Convert to DF for data_editor
                    num_rows="dynamic",
                    key="manual_data_editor"
                )
                # Update session state with edited data
                if isinstance(edited_data, pd.DataFrame):
                    st.session_state.manual_data = edited_data.to_dict('records')
                    # Convert manual data to recipient_df for consistent processing later
                    st.session_state.recipient_df = pd.DataFrame(st.session_state.manual_data)
            else: # if manual_data list is empty (e.g. after clearing an upload)
                 st.info("Click 'Add Row' in the table above (if available with st.data_editor) or re-initialize if needed.")
                 # Re-initialize button or automatically add a row
                 if st.button("Add first manual entry row"):
                     st.session_state.manual_data = [{"Email": "contact@example.com", "Name": "New Contact"}]
                     st.session_state.recipient_df = pd.DataFrame(st.session_state.manual_data)
                     st.rerun()


        # Display the final DataFrame that will be used for mailing
        if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
            st.subheader("Current Recipient Data for Mailing:")
            st.dataframe(st.session_state.recipient_df)
            st.caption(f"{len(st.session_state.recipient_df)} recipients loaded.")
            # Check for 'Email' column
            if 'Email' not in st.session_state.recipient_df.columns:
                st.error("🚨 Critical: The data does not contain an 'Email' column. This column is required to send emails.")
            else:
                # Perform email validation
                valid_email_count = 0
                invalid_format_count = 0
                empty_email_count = 0

                email_column = st.session_state.recipient_df['Email']
                for email_val in email_column:
                    if pd.isna(email_val) or str(email_val).strip() == "":
                        empty_email_count += 1
                    elif is_valid_email(str(email_val).strip()):
                        valid_email_count += 1
                    else:
                        invalid_format_count += 1

                if valid_email_count > 0:
                    st.success(f"✅ {valid_email_count} valid email addresses found.")
                if empty_email_count > 0:
                    st.warning(f"⚠️ {empty_email_count} rows have missing email addresses.")
                if invalid_format_count > 0:
                    st.error(f"🚫 {invalid_format_count} email addresses have an invalid format.")
                    # Optionally, list some invalid emails
                    # invalid_emails_sample = [str(e) for e in email_column if pd.notna(e) and str(e).strip() != "" and not is_valid_email(str(e).strip())][:5]
                    # if invalid_emails_sample:
                    #     st.expander("Show sample invalid emails").write(invalid_emails_sample)

            if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
                if st.button("🗑️ Clear All Recipient Data", key="clear_all_data_button"):
                    st.session_state.recipient_df = None
                    st.session_state.manual_data = [{"Email": "example@test.com", "Name": "Test User"}] # Reset manual
                    st.success("All recipient data cleared.")
                    st.rerun()
        else:
            st.info("No recipient data loaded yet. Upload a file or add data manually.")

        st.markdown("---")
        st.subheader("Attachments")
        if 'attachments' not in st.session_state:
            st.session_state.attachments = [] # List of dicts: {"name": "file.pdf", "data": b"bytes..."}

        uploaded_attachments_list = st.file_uploader(
            "Add Attachments to your Email",
            accept_multiple_files=True,
            key="file_uploader_attachments_widget" # Unique key for the widget
        )

        # Process newly uploaded attachments
        if uploaded_attachments_list:
            current_attachment_names = {att['name'] for att in st.session_state.attachments}
            newly_added_count = 0
            for uploaded_file in uploaded_attachments_list:
                if uploaded_file.name not in current_attachment_names:
                    st.session_state.attachments.append(
                        {"name": uploaded_file.name, "data": uploaded_file.getvalue()}
                    )
                    current_attachment_names.add(uploaded_file.name)
                    newly_added_count += 1
            if newly_added_count > 0:
                st.success(fAdded {newly_added_count} new attachment(s).")
            # Clear the uploader widget's internal state by rerunning if files were processed.
            # This helps avoid re-processing if other parts of the UI cause a rerun.
            # However, st.file_uploader manages its list, so we primarily manage st.session_state.attachments
            # Re-assigning uploaded_attachments_list to None or rerunning might be too aggressive here.
            # The logic above ensures we only add new files from the uploader's current list.

        if st.session_state.attachments:
            st.write(f"{len(st.session_state.attachments)} attachment(s) currently added:")

            # Create columns for attachments list: one for name, one for remove button
            cols_def = [0.8, 0.2] # 80% for name, 20% for button

            for i, att in enumerate(st.session_state.attachments):
                col1, col2 = st.columns(cols_def)
                with col1:
                    st.caption(f"- {att['name']} ({len(att['data'])/1024:.1f} KB)")
                with col2:
                    if st.button(f"Remove", key=f"remove_att_{i}"):
                        st.session_state.attachments.pop(i)
                        st.rerun() # Rerun to update the list immediately

            if st.button("Clear All Attachments", key="clear_all_attachments_button"):
                st.session_state.attachments = []
                st.rerun()
        else:
            st.caption("No attachments added yet.")

        st.markdown("---")
        st.subheader("Email Templates")

        col_template_1, col_template_2 = st.columns(2)

        with col_template_1:
            # Save Template
            if st.session_state.get('email_subject') or st.session_state.get('email_body'):
                template_data = {
                    "subject": st.session_state.get('email_subject', ""),
                    "body": st.session_state.get('email_body', "")
                }
                try:
                    json_template = json.dumps(template_data, indent=2)
                    st.download_button(
                        label="💾 Save Current Email as Template",
                        data=json_template,
                        file_name="email_template.json",
                        mime="application/json",
                        key="download_template_button"
                    )
                except Exception as e:
                    st.error(f"Error preparing template for download: {e}")
            else:
                st.button("💾 Save Current Email as Template", disabled=True, key="download_template_button_disabled")


        with col_template_2:
            # Load Template
            uploaded_template_file = st.file_uploader(
                "📂 Load Email Template (.json)",
                type=['json'],
                key="upload_template_uploader"
            )
            if uploaded_template_file is not None:
                try:
                    template_content = json.load(uploaded_template_file)
                    if isinstance(template_content, dict) and "subject" in template_content and "body" in template_content:
                        st.session_state.email_subject = template_content["subject"]
                        st.session_state.email_body = template_content["body"]
                        st.success("Email template loaded successfully!")
                        # Clear the uploader by rerunning or setting its value to None if possible
                        # For file_uploader, usually just processing it is enough, subsequent reruns won't re-process unless file changes
                        st.rerun()
                    else:
                        st.error("Invalid template file format. Expected JSON with 'subject' and 'body' keys.")
                except json.JSONDecodeError:
                    st.error("Error decoding JSON. Make sure the template file is a valid JSON.")
                except Exception as e:
                    st.error(f"Error loading template: {e}")


    # 3. Email Composition Section
    with st.expander("✍️ Step 3: Compose Your Email", expanded=True): # Expanded by default
        st.subheader("Email Content")

        if 'email_subject' not in st.session_state:
            st.session_state.email_subject = ""
        if 'email_body' not in st.session_state:
            st.session_state.email_body = ""

        st.session_state.email_subject = st.text_input(
            "Subject",
            value=st.session_state.email_subject,
            placeholder="E.g., Happy Birthday, {Name}!"
        )

        st.session_state.email_body = st.text_area(
            "Email Body (HTML is supported)",
            value=st.session_state.email_body,
            height=300, # Increased height for potentially longer HTML
            placeholder="<h1>Hello {Name},</h1>\n<p>This is an <b>HTML</b> email. You can use HTML tags for formatting.</p>\n<p>For example, a link: <a href='https://www.example.com'>Visit Example.com</a></p>\n<p>Best regards,<br>Your Company</p>"
        )
        st.caption("Use placeholders like `{ColumnName}` (e.g., `{Name}`, `{Email}`). Write HTML directly for rich formatting.")

        # AI Subject Line Suggestion (Placeholder)
        if st.button("💡 Suggest Subject Line (AI - basic)"):
            # Basic suggestion, actual AI would be more complex
            if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
                first_row = st.session_state.recipient_df.iloc[0]
                name_placeholder = "{Name}" if "Name" in first_row else ("{" + st.session_state.recipient_df.columns[0] + "}" if len(st.session_state.recipient_df.columns) > 0 else "")
                if name_placeholder:
                    st.session_state.email_subject = f"A Special Message For You, {name_placeholder}!"
                else:
                    st.session_state.email_subject = "An Important Update For You!"
                st.rerun() # To update the subject input field
            else:
                st.warning("Please load recipient data to help generate a subject line.")

        st.markdown("---")
        st.subheader("Email Preview")

        if st.session_state.recipient_df is not None and not st.session_state.recipient_df.empty:
            preview_data_source = st.selectbox(
                "Select recipient for preview:",
                ["First Recipient"] + (list(st.session_state.recipient_df.index[:10]) if len(st.session_state.recipient_df) > 1 else []), # Show more if available
                format_func=lambda x: "First Recipient" if x == "First Recipient" else f"Recipient at Index {x}"
            )

            preview_recipient_data = None
            if preview_data_source == "First Recipient":
                preview_recipient_data = st.session_state.recipient_df.iloc[0]
            elif isinstance(preview_data_source, int): # Index from the list
                 if preview_data_source < len(st.session_state.recipient_df):
                    preview_recipient_data = st.session_state.recipient_df.iloc[preview_data_source]

            if preview_recipient_data is not None:
                try:
                    # Replace placeholders
                    preview_subject = st.session_state.email_subject
                    preview_body = st.session_state.email_body
                    for col_name in preview_recipient_data.index:
                        placeholder = f"{{{col_name}}}"
                        # Ensure data is string for replacement, handle NaN/None
                        replacement_value = str(preview_recipient_data[col_name]) if pd.notna(preview_recipient_data[col_name]) else ""
                        preview_subject = preview_subject.replace(placeholder, replacement_value)
                        preview_body = preview_body.replace(placeholder, replacement_value)

                    st.markdown(f"**To:** `{preview_recipient_data.get('Email', 'N/A - Email column missing or empty')}`")
                    st.markdown(f"**Subject:** {preview_subject}")
                    st.markdown("**Body:**")
                    st.markdown(f"<div style='border: 1px solid #ccc; padding: 10px; border-radius: 5px;'>{preview_body.replace('\n', '<br>')}</div>", unsafe_allow_html=True)
                except Exception as e:
                    st.error(f"Error generating preview: {e}")
                    st.write("Preview Data:", preview_recipient_data.to_dict())
            else:
                st.info("Could not load selected recipient data for preview.")
        else:
            st.info("Load recipient data to see a preview.")

        st.markdown("---")
        st.subheader("Attachments")
        if 'attachments' not in st.session_state:
            st.session_state.attachments = [] # List of dicts: {"name": "file.pdf", "data": b"bytes..."}

        uploaded_attachments_list = st.file_uploader(
            "Add Attachments to your Email",
            accept_multiple_files=True,
            key="file_uploader_attachments_widget" # Unique key for the widget
        )

        if uploaded_attachments_list:
            # This logic aims to add only new files from the uploader to prevent duplication on reruns
            # It also helps if the user uploads the same file again, it won't be re-added if already present by name.
            current_attachment_names = {att['name'] for att in st.session_state.attachments}
            newly_added_files_processed_this_run = False # Flag to see if we need to update display

            for uploaded_file in uploaded_attachments_list:
                if uploaded_file.name not in current_attachment_names:
                    st.session_state.attachments.append(
                        {"name": uploaded_file.name, "data": uploaded_file.getvalue()}
                    )
                    current_attachment_names.add(uploaded_file.name) # Keep track of names added in this session/batch
                    newly_added_files_processed_this_run = True

            # if newly_added_files_processed_this_run:
                # st.experimental_rerun() # Rerun to clear the uploader and reflect additions
                # Potentially too aggressive, let's see. The uploader might persist its list.
                # A more robust way is to manage st.session_state.attachments carefully.
                # The file_uploader widget itself doesn't automatically clear after processing.
                # We are adding to st.session_state.attachments, so the list below will be accurate.
                pass


        if st.session_state.attachments:
            st.write(f"{len(st.session_state.attachments)} attachment(s) currently added:")

            # Create columns for attachments list: one for name, one for remove button
            cols_def = [0.8, 0.2] # 80% for name, 20% for button

            # Iterate backwards if removing to avoid index errors, or create a new list
            attachments_to_remove_indices = []
            for i, att in enumerate(st.session_state.attachments):
                col1, col2 = st.columns(cols_def)
                with col1:
                    st.caption(f"- {att['name']} ({len(att['data'])/1024:.1f} KB)")
                with col2:
                    if st.button(f"Remove", key=f"remove_att_{att['name']}_{i}"): # More unique key
                        attachments_to_remove_indices.append(i)

            if attachments_to_remove_indices:
                for index in sorted(attachments_to_remove_indices, reverse=True):
                    st.session_state.attachments.pop(index)
                st.rerun() # Rerun to update the list immediately

            if st.button("Clear All Attachments", key="clear_all_attachments_button"):
                st.session_state.attachments = []
                st.rerun()
        else:
            st.caption("No attachments added yet.")


import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
import time # For adding small delays

# ... (other parts of the run_sender_app function) ...

    # 4. Sending Section
    with st.expander("🚀 Step 4: Send Emails", expanded=True): # Expanded by default
        st.subheader("Ready to Send?")

        if 'send_log' not in st.session_state:
            st.session_state.send_log = []

        col1, col2 = st.columns([1,3])

        with col1:
            send_button_disabled = not (
                st.session_state.config.get('sender_email') and
                st.session_state.config.get('email_password') and
                st.session_state.config.get('smtp_server') and
                st.session_state.config.get('smtp_port') and
                st.session_state.recipient_df is not None and
                not st.session_state.recipient_df.empty and
                st.session_state.email_subject and
                st.session_state.email_body and
                'Email' in st.session_state.recipient_df.columns
            )

            if send_button_disabled:
                if not (st.session_state.config.get('sender_email') and
                        st.session_state.config.get('email_password') and
                        st.session_state.config.get('smtp_server') and
                        st.session_state.config.get('smtp_port')):
                    st.warning("Sender configuration is incomplete.")
                if st.session_state.recipient_df is None or st.session_state.recipient_df.empty:
                    st.warning("No recipient data loaded.")
                elif 'Email' not in st.session_state.recipient_df.columns:
                     st.warning("Recipient data must have an 'Email' column.")
                if not (st.session_state.email_subject and st.session_state.email_body):
                    st.warning("Email subject or body is empty.")


            if st.button("🚀 Send All Emails", disabled=send_button_disabled, type="primary"):
                st.session_state.send_log = ["Starting email sending process..."]

                config = st.session_state.config
                df = st.session_state.recipient_df
                subject_template = st.session_state.email_subject
                body_template = st.session_state.email_body

                if 'Email' not in df.columns:
                    st.error("Critical: 'Email' column not found in recipient data.")
                    st.session_state.send_log.append("Error: 'Email' column not found.")
                    # No rerun here, let the log show
                else:
                    total_emails = len(df)
                    sent_count = 0
                    failed_count = 0

                    progress_bar = st.progress(0)
                    status_text = st.empty()

                    try:
                        server = None # Initialize server variable
                        if config['smtp_security'] == "SSL":
                            server = smtplib.SMTP_SSL(config['smtp_server'], config['smtp_port'])
                        else: # TLS or None
                            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'])
                            if config['smtp_security'] == "TLS":
                                server.starttls()

                        server.login(config['sender_email'], config['email_password'])
                        st.session_state.send_log.append(f"Logged in to SMTP server {config['smtp_server']}.")

                        for i, row in df.iterrows():
                            recipient_email = row.get('Email')
                            if not recipient_email or pd.isna(recipient_email):
                                log_msg = f"Skipping row {i+1}: Missing email address."
                                st.session_state.send_log.append(log_msg)
                                failed_count += 1
                                progress_bar.progress((i + 1) / total_emails)
                                status_text.text(f"Progress: {i+1}/{total_emails} (Skipped: Missing Email)")
                                continue

                            try:
                                msg = MIMEMultipart()
                                msg['From'] = config['sender_email']
                                msg['To'] = recipient_email

                                current_subject = subject_template
                                current_body = body_template
                                for col_name in df.columns:
                                    placeholder = f"{{{col_name}}}"
                                    replacement_value = str(row.get(col_name, '')) if pd.notna(row.get(col_name)) else ""
                                    current_subject = current_subject.replace(placeholder, replacement_value)
                                    current_body = current_body.replace(placeholder, replacement_value)

                                msg['Subject'] = current_subject
                                msg.attach(MIMEText(current_body, 'html')) # Changed to 'html'

                                # Add attachments from session state
                                if 'attachments' in st.session_state and st.session_state.attachments:
                                    for attachment_data in st.session_state.attachments:
                                        try:
                                            ctype, encoding = mimetypes.guess_type(attachment_data["name"])
                                            if ctype is None or encoding is not None:
                                                ctype = 'application/octet-stream' # Default MIME type

                                            maintype, subtype = ctype.split('/', 1)
                                            part = MIMEBase(maintype, subtype)
                                            part.set_payload(attachment_data["data"])
                                            encoders.encode_base64(part)
                                            part.add_header('Content-Disposition',
                                                            f'attachment; filename="{attachment_data["name"]}"')
                                            msg.attach(part)
                                            st.session_state.send_log.append(f"Attached {attachment_data['name']} to email for {recipient_email}")
                                        except Exception as e_attach:
                                            st.session_state.send_log.append(f"Error attaching {attachment_data.get('name', 'unknown file')} to email for {recipient_email}: {e_attach}")

                                server.sendmail(config['sender_email'], recipient_email, msg.as_string())
                                log_msg = f"Successfully sent email to {recipient_email} (Row {i+1})"
                                st.session_state.send_log.append(log_msg)
                                sent_count +=1
                                time.sleep(0.1) # Small delay to avoid overwhelming the server

                            except Exception as e_send:
                                log_msg = f"Failed to send email to {recipient_email} (Row {i+1}): {e_send}"
                                st.session_state.send_log.append(log_msg)
                                failed_count += 1

                            progress_bar.progress((i + 1) / total_emails)
                            status_text.text(f"Progress: {i+1}/{total_emails} (Sent: {sent_count}, Failed: {failed_count})")

                        if server:
                            server.quit()
                        st.session_state.send_log.append("SMTP server connection closed.")

                    except smtplib.SMTPAuthenticationError:
                        st.error("SMTP Authentication Error: Check your email address and password, or app-specific password settings (e.g., for Gmail). Also ensure 'less secure app access' is enabled if required by your provider.")
                        st.session_state.send_log.append("Error: SMTP Authentication Failed.")
                    except smtplib.SMTPConnectError:
                        st.error(f"SMTP Connection Error: Could not connect to server {config['smtp_server']}:{config['smtp_port']}. Check server address and port.")
                        st.session_state.send_log.append("Error: SMTP Connection Failed.")
                    except Exception as e_smtp:
                        st.error(f"An SMTP error occurred: {e_smtp}")
                        st.session_state.send_log.append(f"Error: SMTP problem - {e_smtp}")
                    finally:
                        if server: # Ensure server is closed if an error occurs mid-process
                            try:
                                server.quit()
                            except: # Ignore errors on quit if already disconnected
                                pass
                        final_summary = f"Email sending process finished. Total: {total_emails}, Sent: {sent_count}, Failed/Skipped: {failed_count}."
                        st.session_state.send_log.append(final_summary)
                        status_text.text(final_summary)
                        st.balloons() # Fun little success indicator
                        # No st.rerun() here, we want the log to persist.

        with col2:
            st.subheader("Sending Progress & Log")
            log_display = st.text_area(
                "Log:",
                value="\n".join(st.session_state.send_log),
                height=300,
                key="send_log_display",
                disabled=True
            )

def show_landing_page():
    st.header("Welcome to the Mass Email Sender Deluxe!")
    # Using a more generic and appealing image if possible, or keeping the current one.
    # For stability, I'll keep the existing one. A custom uploaded one would be better for a real app.
    st.image("https://img.freepik.com/free-vector/arroba-email-symbol-particle-style_78370-3010.jpg", width=200, caption="Streamline Your Email Campaigns")

    st.markdown("""
    Empower your communication strategy with the **Mass Email Sender Deluxe**. This application is designed for businesses and organizations to send personalized mass emails to customers, staff, or any contact list with ease and efficiency.
    """)

    st.subheader("✨ Key Features")
    st.markdown("""
    *   **Intuitive Configuration:** Quickly set up your sending email account using common providers (Gmail, Outlook, Yahoo) or custom SMTP settings.
    *   **Versatile Data Handling:**
        *   Upload recipient data seamlessly from CSV or Excel files.
        *   Manually add or edit recipient information directly within the app using a user-friendly table editor.
    *   **Dynamic Personalization:** Craft unique messages for each recipient. Use placeholders (e.g., `{Name}`, `{OrderID}`, `{MembershipLevel}`) in your email subject and body that dynamically pull data from your recipient list.
    *   **Live Preview:** See exactly how your email will look for a specific recipient before sending the entire batch.
    *   **Real-time Sending Progress:** Monitor the status of your email campaign with a detailed log, including successful sends and any issues encountered.
    *   **Basic AI Assistance:** Get a little help with a simple AI-powered subject line suggestion. (More advanced AI features planned!)
    """)

    st.subheader("🚀 How to Use the Sender App")
    st.markdown("""
    1.  **Navigate to "🚀 Send Emails"** using the sidebar menu.
    2.  **Step 1: Configure Sender Details:**
        *   Enter your email address and password (or App Password for services like Gmail).
        *   Select your email provider for pre-filled SMTP settings or enter custom SMTP details.
    3.  **Step 2: Add Recipient Data:**
        *   Choose "Upload File" to load a CSV or Excel file.
        *   Or, select "Manual Entry" to use the built-in table editor to add or modify contacts.
        *   Ensure your data includes an "Email" column for recipients. Other columns can be used for personalization.
    4.  **Step 3: Compose Your Email:**
        *   Write a compelling **Subject** and **Email Body**.
        *   Insert placeholders (e.g., `{Name}`) corresponding to your data column headers to personalize messages.
        *   Use the **Preview** section to check how the email will appear for a selected recipient.
    5.  **Step 4: Send Emails:**
        *   Once everything is set up, click the **"🚀 Send All Emails"** button.
        *   Monitor the **Sending Progress & Log** for real-time updates.
    """)

    st.subheader("💻 Download & Run Locally")
    st.markdown("""
    This is a Streamlit web application. To run it on your own computer:
    1.  **Python:** Ensure you have Python (version 3.7 or newer) installed.
    2.  **Install Libraries:** Open your terminal or command prompt and install the necessary Python packages:
        ```bash
        pip install streamlit pandas openpyxl
        ```
    3.  **Get the Code:** Save the application code as a Python file (e.g., `email_app.py`).
    4.  **Run the App:** Navigate to the directory where you saved the file and execute:
        ```bash
        streamlit run email_app.py
        ```
    Your web browser should automatically open to the application.
    """)

    st.info("💡 **Tip:** For Gmail, you might need to generate an 'App Password' for SMTP access if 2-Step Verification is enabled. Also, ensure 'Less Secure App Access' is appropriately set if not using App Passwords (though App Passwords are more secure).")

    st.markdown("---")
    st.subheader("🔮 Exciting AI Features on the Horizon!")
    st.markdown("""
    We're working on integrating more advanced AI capabilities to make your email campaigns even smarter:
    *   **Advanced AI Subject Lines:** Suggestions optimized for engagement and open rates.
    *   **AI Content Generation & Refinement:** Get help drafting compelling and grammatically correct email copy.
    *   **Smarter Personalization Engine:** AI-driven recommendations for tailoring messages based on recipient attributes and behavior (requires more data).
    *   **Optimal Send Time Analysis:** Insights into when to send emails for maximum impact (requires integration with analytics).
    """)

    st.markdown("---")
    st.markdown("We hope this tool significantly streamlines your communication efforts! Feedback is always welcome.")

if __name__ == "__main__":
    main()
