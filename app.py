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
import os
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

        # Initialize Google Sheets related session state variables if they don't exist
        if 'google_creds_json_content' not in st.session_state:
            st.session_state.google_creds_json_content = None
        if 'google_auth_flow_completed' not in st.session_state:
            st.session_state.google_auth_flow_completed = False
        if 'google_credentials' not in st.session_state: # Stores the actual OAuth2 credentials object
            st.session_state.google_credentials = None
        if 'google_auth_flow_obj' not in st.session_state: # Stores the google_auth_oauthlib.flow.Flow object
            st.session_state.google_auth_flow_obj = None


        data_input_method = st.radio(
            "Choose data input method:",
            ("Upload File", "Manual Entry", "Google Sheets"), # Added Google Sheets
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

        if data_input_method == "Google Sheets":
            st.subheader("Import Data from Google Sheets")
            st.markdown("""
            To use this feature, you need to set up credentials in the Google Cloud Platform (GCP)
            and provide the `credentials.json` file. Here’s a summary of the steps:

            **1. Google Cloud Platform Project Setup:**
            *   Go to the [Google Cloud Console](https://console.cloud.google.com/).
            *   Create a new project (or select an existing one).
            *   Give your project a name (e.g., "Streamlit Sheets Integration") and note the Project ID.

            **2. Enable APIs:**
            *   In your GCP project, go to "APIs & Services" > "Library".
            *   Search for and enable the **"Google Sheets API"**.
            *   Search for and enable the **"Google Drive API"** (this might be needed for future enhancements like a file picker; for reading specific sheet URLs, Sheets API alone is often enough, but enabling it is good practice).

            **3. Create OAuth 2.0 Credentials:**
            *   Go to "APIs & Services" > "Credentials".
            *   Click "+ CREATE CREDENTIALS" and select "OAuth client ID".
            *   If prompted, configure the "OAuth consent screen":
                *   Choose "User Type" (likely "External" if you're using a personal Gmail, or "Internal" if you have a Google Workspace org).
                *   Fill in the app name (e.g., "Streamlit Mass Mailer"), user support email, and developer contact information. Click "SAVE AND CONTINUE".
                *   **Scopes:** You don't need to add scopes here on the consent screen page itself if your application requests them dynamically, but be aware of the scopes your app will request (e.g., `https://www.googleapis.com/auth/spreadsheets.readonly`). Click "SAVE AND CONTINUE".
                *   **Test Users (for External apps in testing mode):** Add your Google account email address as a test user. Click "SAVE AND CONTINUE".
            *   Back on the "Create OAuth client ID" screen:
                *   Select "Application type": **"Web application"**.
                *   Give it a name (e.g., "Streamlit Sheets Client").
                *   **Authorized redirect URIs:** This is crucial.
                    *   For local development with Streamlit (default port 8501), add: `http://localhost:8501`
                    *   *If you deploy your Streamlit app, you MUST add the deployed app's full URL as a redirect URI as well.*
                    *   The app will attempt to guide you on the exact redirect URI it expects once you start the authentication process.
                *   Click "CREATE".

            **4. Download `credentials.json`:**
            *   After creation, a dialog will show your Client ID and Client Secret. **Download the JSON file** by clicking the download icon next to your newly created OAuth 2.0 Client ID in the credentials list. Rename this file to `credentials.json` if it's different.
            *   **Keep this file secure!** It contains your client secret.

            **5. Upload `credentials.json` below.**

            **Required Python Libraries:**
            Make sure you have these libraries installed in your Python environment:
            ```bash
            pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib pandas
            ```
            (Pandas is likely already installed if you've used other features of this app).
            """)
            uploaded_google_creds = st.file_uploader(
                "Upload your `credentials.json` file from GCP",
                type=['json'],
                key="google_creds_uploader",
                help="Download this from your Google Cloud Platform project's OAuth 2.0 Client ID settings."
            )

            if uploaded_google_creds is not None:
                try:
                    # Read the content and store it as a string in session_state
                    # The actual parsing into a dict will happen when initiating the flow
                    st.session_state.google_creds_json_content = uploaded_google_creds.getvalue().decode('utf-8')
                    st.success("`credentials.json` uploaded successfully.")
                except Exception as e:
                    st.error(f"Error reading credentials file: {e}")
                    st.session_state.google_creds_json_content = None

            if st.session_state.google_creds_json_content:
                # Display some info from the credentials if needed (e.g., project_id, client_id) - be careful not to expose client_secret
                try:
                    creds_dict = json.loads(st.session_state.google_creds_json_content)
                    client_id = creds_dict.get("web", {}).get("client_id") or creds_dict.get("installed", {}).get("client_id")
                    if client_id:
                        st.caption(f"Credentials loaded. Client ID: {client_id[:10]}...{client_id[-4:] if len(client_id) > 14 else ''}")
                except json.JSONDecodeError:
                    st.error("Uploaded credentials file is not valid JSON.")
                    st.session_state.google_creds_json_content = None # Invalidate if not JSON
                except Exception: # Catch any other error if structure is unexpected
                    st.warning("Could not parse Client ID from credentials file, but file content is stored.")


            # This button will trigger Part 1 of OAuth flow
            if st.button(
                "🔒 Authenticate with Google",
                key="google_auth_start_button_actual",  # Changed key to avoid conflict with any previous placeholder
                disabled=not st.session_state.google_creds_json_content or st.session_state.google_auth_flow_completed,
                help="You need to upload your credentials.json first."
            ):
                st.session_state.google_auth_start_button_clicked = True # Flag that the button was clicked
                # Clear previous auth attempt state if any
                st.session_state.google_auth_show_redirect_url_input = False
                st.session_state.google_auth_flow_obj = None
                st.session_state.google_auth_oauth_state = None


            if st.session_state.google_auth_flow_completed:
                st.success("✅ Successfully authenticated with Google.")
                # Attempt to display user email if available (optional, needs more scopes typically)
                # if st.session_state.google_credentials and hasattr(st.session_state.google_credentials, 'id_token') and st.session_state.google_credentials.id_token:
                #     try:
                #         id_info = id_token.verify_oauth2_token(st.session_state.google_credentials.id_token, google.auth.transport.requests.Request(), st.session_state.google_credentials.client_id)
                #         st.caption(f"Authenticated as: {id_info.get('email')}")
                #     except Exception as e:
                #         st.caption("Authenticated (could not retrieve email).")

                if st.button("🔄 Clear Google Authentication", key="google_logout_button_active"):
                    # Clear all Google Sheets related session state
                    keys_to_clear = [
                        'google_creds_json_content', 'google_credentials',
                        'google_auth_flow_completed', 'google_auth_flow_obj',
                        'google_auth_oauth_state', 'google_auth_show_redirect_url_input',
                        'google_redirect_uri', 'redirect_url_input_gauth',
                        'google_sheet_url_id', 'google_sheet_tab_name', 'google_sheet_range',
                        'google_sheet_load_in_progress'
                    ]
                    for key in keys_to_clear:
                        if key in st.session_state:
                            del st.session_state[key]

                    st.info("Google Authentication and sheet settings cleared.")
                    st.rerun()

            # Part 1: Initiate OAuth Flow (triggered by the flag set from button click)
            if st.session_state.get("google_auth_start_button_clicked") and not st.session_state.google_auth_flow_completed:
                try:
                    client_config = json.loads(st.session_state.google_creds_json_content)
                    if "web" not in client_config and "installed" not in client_config:
                        st.error("Invalid credentials.json format: Missing 'web' or 'installed' key.")
                    else:
                        # For local dev, redirect_uri MUST be in GCP console.
                        # Streamlit's default port is 8501.
                        redirect_uri = "http://localhost:8501"
                        st.session_state.google_redirect_uri = redirect_uri # Store for later verification/use potentially

                        flow = Flow.from_client_config(
                            client_config,
                            scopes=['https://www.googleapis.com/auth/spreadsheets.readonly'],
                            redirect_uri=redirect_uri
                        )

                        authorization_url, state = flow.authorization_url(
                            access_type='offline',
                            prompt='consent'
                        )

                        st.session_state.google_auth_flow_obj = flow
                        st.session_state.google_auth_oauth_state = state

                        st.markdown(f"""
                        **Step 1: Authorize Access**
                        <a href="{authorization_url}" target="_blank" style="display:inline-block;padding:0.5em 1em;background-color:#4CAF50;color:white;text-align:center;text-decoration:none;border-radius:4px;">Click here to authorize with Google</a>
                        """, unsafe_allow_html=True)
                        st.info("After authorizing, Google will redirect you. Your browser might show a 'This site can’t be reached' or similar error if you're running locally – this is expected. Copy the ENTIRE URL from your browser's address bar (it will start with `http://localhost:8501/?state=...`).")
                        st.session_state.google_auth_show_redirect_url_input = True

                except json.JSONDecodeError:
                    st.error("Could not parse `credentials.json`. Ensure it's valid JSON.")
                except Exception as e:
                    st.error(f"Could not initiate Google authentication: {str(e)}")
                finally:
                    st.session_state.google_auth_start_button_clicked = False # Reset flag

            # Input for the redirect URL (shown after clicking authenticate)
            if st.session_state.get("google_auth_show_redirect_url_input") and not st.session_state.google_auth_flow_completed:
                st.markdown("**Step 2: Paste Authorization Response URL**")
                redirect_url_from_user = st.text_area( # text_area for long URLs
                    "Paste the full URL from Google after authorization here:",
                    key="redirect_url_input_gauth", # Changed key
                    height=100,
                    help="Example: http://localhost:8501/?state=...&code=...&scope=..."
                )
                # Part 2 trigger button
                if st.button(
                    "🔗 Complete Authentication",
                    key="google_auth_complete_button_actual",
                    disabled=not redirect_url_from_user
                ):
                    if not st.session_state.google_auth_flow_obj:
                        st.error("Authentication flow not initiated or session expired. Please try authenticating again.")
                    else:
                        try:
                            # CSRF protection: Ensure the state parameter matches.
                            # The redirect_url_from_user needs to be parsed to extract the state.
                            # A robust way is to use urllib.parse.urlparse and parse_qs.
                            # For simplicity here, we'll assume user pastes the full URL and `fetch_token` handles it.
                            # However, `google-auth-oauthlib` expects `authorization_response` kwarg for fetch_token.
                            # The `state` is verified by the library if passed correctly to authorization_url and present in response.

                            # Reconstruct the flow object if not in session state (though it should be)
                            # flow = st.session_state.google_auth_flow_obj

                            # The redirect_uri used when creating the flow initially must be passed again if it was part of the client_config.
                            # However, from_client_config with redirect_uri set in the flow object should handle it.

                            # Ensure the state matches (manual check for extra safety, though library might do it)
                            parsed_redirect_url = urlparse(redirect_url_from_user)
                            query_params = parse_qs(parsed_redirect_url.query)
                            returned_state = query_params.get('state', [None])[0]

                            if returned_state != st.session_state.get('google_auth_oauth_state'):
                                st.error("OAuth state mismatch (CSRF protection). Please try authenticating again.")
                            else:
                                st.session_state.google_auth_flow_obj.fetch_token(authorization_response=redirect_url_from_user)
                                st.session_state.google_credentials = st.session_state.google_auth_flow_obj.credentials
                                st.session_state.google_auth_flow_completed = True

                                # Clear temporary auth flow variables
                                st.session_state.google_auth_show_redirect_url_input = False
                                st.session_state.google_auth_flow_obj = None
                                st.session_state.google_auth_oauth_state = None
                                st.session_state.redirect_url_input_gauth = "" # Clear the input text_area

                                st.success("✅ Successfully authenticated with Google!")
                                st.info("You can now proceed to load data from Google Sheets.")
                                st.rerun()

                        except Exception as e:
                            st.error(f"Error completing authentication: {str(e)}")
                            st.session_state.google_auth_flow_completed = False
                            st.session_state.google_credentials = None

            # --- UI for Sheet Selection (if authenticated) ---
            if st.session_state.google_auth_flow_completed and st.session_state.google_credentials:
                st.markdown("---")
                st.subheader("📄 Select Google Sheet and Range")

                sheet_url_or_id = st.text_input(
                    "Google Sheet URL or ID:",
                    key="google_sheet_url_id",
                    placeholder="e.g., https://docs.google.com/spreadsheets/d/your_sheet_id/edit or just your_sheet_id"
                )
                sheet_name = st.text_input(
                    "Sheet Name (Tab Name):",
                    value="Sheet1", # Common default
                    key="google_sheet_tab_name",
                    placeholder="e.g., Sheet1, Contacts Q1"
                )
                sheet_range = st.text_input(
                    "Data Range (optional, e.g., A1:D50 or MyNamedRange):",
                    key="google_sheet_range",
                    placeholder="Leave empty to attempt reading the entire sheet or used range."
                )

                if st.button("📥 Load Data from Google Sheet", key="load_g_sheet_data_button"):
                    # Logic for this button will be in the next step
                    st.info("Loading data from Google Sheet... (Implementation pending)")
                    if not sheet_url_or_id.strip():
                        st.warning("Please provide the Google Sheet URL or ID.")
                    elif not sheet_name.strip():
                        st.warning("Please provide the Sheet Name (Tab Name).")
                    else:
                        try:
                            st.session_state.google_sheet_load_in_progress = True # Flag to show spinner/message
                            st.rerun() # Rerun to show message immediately
                        except Exception as e_initial_load_setup: # Should not happen often
                             st.error(f"Error preparing to load sheet: {e_initial_load_setup}")

            # Actual data loading logic, triggered by the flag and rerun
            if st.session_state.get("google_sheet_load_in_progress"):
                with st.spinner("Fetching data from Google Sheet... Please wait."):
                    try:
                        creds = st.session_state.google_credentials
                        if not creds or not creds.valid:
                            if creds and creds.expired and creds.refresh_token:
                                st.info("Google credentials expired, attempting to refresh...")
                                creds.refresh(Request()) # google.auth.transport.requests.Request
                                st.session_state.google_credentials = creds # Store refreshed credentials
                                st.success("Credentials refreshed.")
                            else:
                                st.error("Google authentication is invalid or expired. Please re-authenticate.")
                                st.session_state.google_auth_flow_completed = False # Force re-auth
                                st.session_state.google_sheet_load_in_progress = False
                                st.rerun()

                        service = build('sheets', 'v4', credentials=creds)

                        # Extract Sheet ID from URL or use directly if ID is provided
                        sheet_id_input = st.session_state.get('google_sheet_url_id', "").strip()
                        # Basic regex to extract ID from a typical sheets URL
                        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", sheet_id_input)
                        if match:
                            spreadsheet_id = match.group(1)
                        else:
                            spreadsheet_id = sheet_id_input # Assume it's an ID

                        # Construct range name
                        sheet_name_input = st.session_state.get('google_sheet_tab_name',"Sheet1").strip()
                        range_input = st.session_state.get('google_sheet_range',"").strip()

                        if not range_input: # If range is empty, try to get all data from the sheet
                            range_to_fetch = sheet_name_input
                        else:
                            # Ensure sheet name is part of range if not already
                            if '!' not in range_input:
                                range_to_fetch = f"{sheet_name_input}!{range_input}"
                            else:
                                range_to_fetch = range_input

                        st.caption(f"Attempting to fetch: Spreadsheet ID='{spreadsheet_id}', Range='{range_to_fetch}'")

                        result = service.spreadsheets().values().get(
                            spreadsheetId=spreadsheet_id, range=range_to_fetch
                        ).execute()

                        values = result.get('values', [])

                        if not values:
                            st.warning("No data found in the specified sheet/range, or the sheet is empty.")
                            st.session_state.recipient_df = pd.DataFrame() # Empty DataFrame
                        else:
                            # Assuming the first row is headers
                            headers = values[0]
                            data_rows = values[1:]
                            df = pd.DataFrame(data_rows, columns=headers)

                            st.session_state.recipient_df = df
                            st.session_state.manual_data = df.to_dict('records') # For manual editing consistency
                            st.success(f"Successfully loaded {len(df)} rows from Google Sheet '{sheet_name_input}'.")
                            if 'Email' not in df.columns:
                                st.warning("The loaded data does not contain an 'Email' column, which is required for sending emails.")

                    except Exception as e:
                        st.error(f"An error occurred while fetching data from Google Sheets: {str(e)}")
                        # More specific error handling could be added here for common API errors
                        # e.g., if e.resp.status == 403 (PermissionDenied), etc.
                    finally:
                        st.session_state.google_sheet_load_in_progress = False # Reset flag
                        # No st.rerun() here, let the result display. Another rerun will happen if user interacts.


        # --- Contact List Management UI ---
        st.markdown("---")
        st.subheader("🗂️ Contact List Management")

        CONTACT_LIST_DIR = "contact_lists"
        # Ensure contact list directory exists
        if not os.path.exists(CONTACT_LIST_DIR):
            try:
                os.makedirs(CONTACT_LIST_DIR)
            except OSError as e:
                st.error(f"Could not create directory for contact lists: {CONTACT_LIST_DIR}. Saved lists may not work. Error: {e}")
                # If directory can't be made, this feature will be largely non-functional.

        def get_saved_lists():
            if not os.path.exists(CONTACT_LIST_DIR) or not os.path.isdir(CONTACT_LIST_DIR):
                return []
            try:
                files = [f for f in os.listdir(CONTACT_LIST_DIR) if f.endswith(".csv")]
                return sorted([os.path.splitext(f)[0] for f in files])
            except Exception as e:
                # st.error(f"Error reading saved contact lists: {e}") # Can be noisy
                return []

        saved_lists = get_saved_lists()

        list_name_col, save_btn_col = st.columns([3,1])
        with list_name_col:
            save_list_name = st.text_input(
                "Enter name to save current list:",
                key="contact_list_name_input",
                placeholder="E.g., Newsletter Q1"
            )
        with save_btn_col:
            st.write("") # Spacer for button alignment
            st.write("") # Spacer for button alignment
            if st.button("💾 Save List", key="save_contact_list_btn", use_container_width=True):
                if not save_list_name.strip():
                    st.warning("Please enter a name for the contact list.")
                elif st.session_state.recipient_df is None or st.session_state.recipient_df.empty:
                    st.warning("No recipient data to save.")
                else:
                    # Improved filename sanitization
                    temp_name = save_list_name.strip()
                    # Remove characters not allowed in typical filenames, replace multiple spaces/underscores with single underscore
                    safe_list_name = re.sub(r'[^\w\s-]', '', temp_name)
                    safe_list_name = re.sub(r'\s+', '_', safe_list_name).strip('_')

                    if not safe_list_name:
                        st.error("Invalid list name. Please use letters, numbers, spaces, underscores, or hyphens. Name cannot be empty after sanitization.")
                    else:
                        filepath = os.path.join(CONTACT_LIST_DIR, safe_list_name + ".csv")

                        # Overwrite confirmation logic
                        if os.path.exists(filepath) and f"overwrite_confirmed_{safe_list_name}" not in st.session_state:
                            st.session_state[f"confirm_overwrite_{safe_list_name}"] = True
                            st.rerun() # Rerun to show confirmation buttons

                        if st.session_state.get(f"confirm_overwrite_{safe_list_name}"):
                            st.warning(f"List '{safe_list_name}' already exists.")
                            col_ow_1, col_ow_2 = st.columns(2)
                            with col_ow_1:
                                if st.button(f"✅ Yes, Overwrite '{safe_list_name}'", key=f"overwrite_yes_{safe_list_name}"):
                                    st.session_state[f"overwrite_confirmed_{safe_list_name}"] = True
                                    del st.session_state[f"confirm_overwrite_{safe_list_name}"]
                                    st.rerun()
                            with col_ow_2:
                                if st.button(f"❌ No, Cancel Overwrite", key=f"overwrite_no_{safe_list_name}"):
                                    del st.session_state[f"confirm_overwrite_{safe_list_name}"]
                                    st.info(f"Save operation for '{safe_list_name}' cancelled.")
                                    st.rerun()
                        else:
                            # Proceed with save if no confirmation needed or if confirmed
                            try:
                                st.session_state.recipient_df.to_csv(filepath, index=False)
                                st.success(f"Contact list '{safe_list_name}' saved successfully!")
                                st.session_state.contact_list_name_input = ""
                                if f"overwrite_confirmed_{safe_list_name}" in st.session_state:
                                    del st.session_state[f"overwrite_confirmed_{safe_list_name}"]
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error saving contact list '{safe_list_name}': {e}")

        if not saved_lists:
            st.caption("No saved contact lists found.")
        else:
            st.markdown("---") # Separator
            load_select_col, load_btn_col, del_btn_col = st.columns([2,1,1])
            with load_select_col:
                selected_list_to_action = st.selectbox(
                    "Select a saved list to load or delete:",
                    options=[""] + saved_lists,
                    key="select_saved_list_dropdown",
                    index=0
                )
            with load_btn_col:
                st.write("") # Spacer
                st.write("") # Spacer
                if st.button("📂 Load List", key="load_contact_list_btn", disabled=not selected_list_to_action, use_container_width=True):
                    filepath = os.path.join(CONTACT_LIST_DIR, selected_list_to_action + ".csv")
                    try:
                        loaded_df = pd.read_csv(filepath)
                        st.session_state.recipient_df = loaded_df
                        st.session_state.manual_data = loaded_df.to_dict('records')
                        st.success(f"List '{selected_list_to_action}' loaded!")
                        st.rerun()
                    except FileNotFoundError:
                        st.error(f"List '{selected_list_to_action}' not found.")
                    except pd.errors.EmptyDataError:
                        st.error(f"List '{selected_list_to_action}' is empty or not valid CSV.")
                    except Exception as e:
                        st.error(f"Error loading list '{selected_list_to_action}': {e}")

            with del_btn_col:
                st.write("") # Spacer
                st.write("") # Spacer
                if st.button("🗑️ Delete List", key="delete_contact_list_btn", disabled=not selected_list_to_action, use_container_width=True):
                    if selected_list_to_action: # Ensure a list is selected
                        st.session_state[f"confirm_delete_{selected_list_to_action}"] = True
                        st.rerun() # Rerun to show confirmation

                if selected_list_to_action and st.session_state.get(f"confirm_delete_{selected_list_to_action}"):
                    st.error(f"Are you sure you want to delete the list '{selected_list_to_action}'? This action cannot be undone.")
                    col_del_1, col_del_2 = st.columns(2)
                    with col_del_1:
                        if st.button(f"✅ Yes, Delete '{selected_list_to_action}'", key=f"delete_yes_{selected_list_to_action}"):
                            filepath = os.path.join(CONTACT_LIST_DIR, selected_list_to_action + ".csv")
                            try:
                                os.remove(filepath)
                                st.success(f"Contact list '{selected_list_to_action}' deleted!")
                                del st.session_state[f"confirm_delete_{selected_list_to_action}"]
                                # Reset selectbox to avoid trying to delete again on auto-rerun
                                st.session_state.select_saved_list_dropdown = ""
                                st.rerun()
                            except FileNotFoundError:
                                st.error(f"List '{selected_list_to_action}' not found for deletion.")
                                del st.session_state[f"confirm_delete_{selected_list_to_action}"]
                                st.rerun()
                            except Exception as e:
                                st.error(f"Error deleting list '{selected_list_to_action}': {e}")
                                del st.session_state[f"confirm_delete_{selected_list_to_action}"]
                                st.rerun()
                    with col_del_2:
                        if st.button(f"❌ No, Keep List", key=f"delete_no_{selected_list_to_action}"):
                            del st.session_state[f"confirm_delete_{selected_list_to_action}"]
                            st.info(f"Deletion of '{selected_list_to_action}' cancelled.")
                            st.rerun()

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
from google_auth_oauthlib.flow import Flow # Added for Google OAuth
from urllib.parse import urlparse, parse_qs # Added for state verification
from googleapiclient.discovery import build # Added for Sheets API
from google.auth.transport.requests import Request # For token refresh
# from google.oauth2.credentials import Credentials # Might need later for building service

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
