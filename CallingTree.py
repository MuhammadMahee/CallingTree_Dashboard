from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import pandas as pd
from io import BytesIO
import streamlit as st
#pip install Office365-REST-Python-Client pandas openpyxl streamlit
st.set_page_config(page_title="Calling Tree DataFrame", page_icon=":phone:", layout="wide")
# @st.cache_data
 
def get_Calling_Tree():
    # SharePoint info
    site_url = "https://xclusivetrading-my.sharepoint.com/personal/m_mahee_xclusivetradinginc_com/"
    relative_url = "/personal/m_mahee_xclusivetradinginc_com/Documents/Desktop/Python/Calling Tree/Jun-25/Calling tree.xlsx"
    username = "developers@xclusivetradinginc.com"
    password = "devxti@2025"
   
    # Authenticate
    ctx_auth = AuthenticationContext(site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(site_url, ctx_auth)
 
        # Prepare a BytesIO object to hold the downloaded file
        file = ctx.web.get_file_by_server_relative_url(relative_url)
        file_buffer = BytesIO()
        file.download(file_buffer)
        ctx.execute_query()
 
        # Move to beginning of buffer before reading
        file_buffer.seek(0)
 
        # Load Excel into DataFrame
        df = pd.read_excel(file_buffer, sheet_name="Main", engine='openpyxl')
       
    else:
        print("Authentication failed.")
   
    return df
 
 
print("Calling Tree DataFrame:")
calling_tree_df = get_Calling_Tree()    
# Render DataFrame inside a scrollable container
with st.container():

    
    st.data_editor(calling_tree_df, use_container_width=True, hide_index=True, height=720)
    
