# --- Required Libraries ---
import streamlit as st
import pandas as pd
from io import BytesIO
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import datetime
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font

# --- Streamlit Page Configuration ---
st.set_page_config(page_title="Calling Tree Login", layout="wide")

# --- Login Credentials (Hardcoded) ---
USERNAME = "Backoffice"
PASSWORD = "xticallingtree"

# --- Session State Initialization ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# --- Login Page Function ---
def login_page():
    st.title("🔐 Calling Tree Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username == USERNAME and password == PASSWORD:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("❌ Incorrect username or password.")

# --- Function to Load Data from SharePoint ---
def get_Calling_Tree():
    site_url = "https://xclusivetrading-my.sharepoint.com/personal/m_mahee_xclusivetradinginc_com/"
    relative_url = "/personal/m_mahee_xclusivetradinginc_com/Documents/Desktop/Python/Calling Tree/Jun-25/Calling tree.xlsx"
    username = "developers@xclusivetradinginc.com"
    password = "devxti@2025"

    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
        file = ctx.web.get_file_by_server_relative_url(relative_url)
        file_buffer = BytesIO()
        file.download(file_buffer)
        ctx.execute_query()
        file_buffer.seek(0)
        df = pd.read_excel(file_buffer, sheet_name="Main", engine='openpyxl')
        return df
    else:
        st.error("❌ SharePoint authentication failed.")
        return pd.DataFrame()

# --- Main App Logic ---
if st.session_state.logged_in:
    calling_tree_df = get_Calling_Tree()

    if not calling_tree_df.empty:
        # Top bar: Logout (left), Download (center)
        col1, col2, col3 = st.columns([1, 2, 1])

        with col1:
            if st.button("🔓 Logout"):
                st.session_state.logged_in = False
                st.rerun()

        with col2:
            file_month = datetime.datetime.now().strftime("%b-%y")
            filename = f"Calling Tree - {file_month}.xlsx"

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                calling_tree_df.to_excel(writer, sheet_name="Main", index=False)
                workbook = writer.book
                worksheet = writer.sheets["Main"]

                header_fill = PatternFill(start_color="0C769E", end_color="0C769E", fill_type="solid")
                body_fill = PatternFill(start_color="CAEDFB", end_color="CAEDFB", fill_type="solid")
                border = Border(
                    left=Side(style="thin"),
                    right=Side(style="thin"),
                    top=Side(style="thin"),
                    bottom=Side(style="thin")
                )
                align_center = Alignment(horizontal="center", vertical="center")
                bold_white = Font(bold=True, color="FFFFFF")

                # Header styling
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = bold_white
                    cell.alignment = align_center
                    cell.border = border

                # Data styling
                for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, max_col=worksheet.max_column):
                    for cell in row:
                        cell.fill = body_fill
                        cell.alignment = align_center
                        cell.border = border

            output.seek(0)
            st.download_button(
                label="⬇️ Download Calling Tree",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        # Display the editable DataFrame
        st.data_editor(
            calling_tree_df,
            use_container_width=True,
            hide_index=True,
            height=720
        )
    else:
        st.warning("No data loaded from the SharePoint file.")
else:
    login_page()
