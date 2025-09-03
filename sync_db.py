import streamlit as st
import pandas as pd
import mysql.connector
import os
from io import BytesIO


# MySQL Connection
def get_connection():
    return mysql.connector.connect(
        host='localhost',
        user='Mahee',
        password='Muhammadmahee123.',
        database='callingtree'
    )

# Truncate Table
def truncate_table():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("TRUNCATE TABLE calling_trees")
    conn.commit()
    conn.close()
    st.success("Table truncated successfully.")

# Insert a single file
def insert_file(name, file_data):
    conn = get_connection()
    cursor = conn.cursor()
    query = "INSERT INTO calling_trees (name, data_blob) VALUES (%s, %s)"
    cursor.execute(query, (name, file_data))
    conn.commit()
    conn.close()

# Insert all files from 'Calling_Trees' folder
def insert_from_folder():
    truncate_table()
    folder_path = os.path.join(os.getcwd(), "Calling_Trees")
    if not os.path.exists(folder_path):
        st.error(f"'Calling_Trees' folder not found at: {folder_path}")
        return

    files = os.listdir(folder_path)
    if not files:
        st.warning("No files found in the 'Calling_Trees' folder.")
        return

    inserted_count = 0
    for file_name in files:
        file_path = os.path.join(folder_path, file_name)
        if os.path.isfile(file_path):
            with open(file_path, "rb") as f:
                file_data = f.read()
                insert_file(file_name, file_data)
                inserted_count += 1

    st.success(f"Inserted {inserted_count} files into the database.")

# Show Table (without blob)
def show_table():
    conn = get_connection()
    df = pd.read_sql("SELECT id, name, last_updated FROM calling_trees", conn)
    conn.close()
    st.write(df)

def select_calling_tree(date="Aug_25"):
    conn = get_connection()
    
    # MySQL uses %s as the parameter placeholder
    query = "SELECT data_blob FROM calling_trees WHERE name LIKE %s"
    like_pattern = f"%{date}%"
    
    df_blob = pd.read_sql(query, conn, params=[like_pattern])
    conn.close()
    
    if df_blob.empty:
        st.warning("No data found for the given date.")
        return

    # Assuming one result
    blob_data = df_blob.iloc[0]["data_blob"]

    # Convert binary blob (Excel file) to DataFrame
    excel_file = BytesIO(blob_data)
    df_excel = pd.read_excel(excel_file)

    return df_excel

def update_data_button():
    if st.button("Sync Database"):
        insert_from_folder()

