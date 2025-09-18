import mysql.connector
import pandas as pd
import os
import glob
import time
import msvcrt  # only works in Windows console

# ------------------ Paths ------------------
DOORCODE_FILE = r"C:\Users\MuhammadMahee\OneDrive - Xclusive Trading Inc\Door Codes.xlsx"
NTID_FILE = r"C:\Users\MuhammadMahee\OneDrive - Xclusive Trading Inc\NTID's.xlsx"
CALLING_TREE_FOLDER = r"C:\Users\MuhammadMahee\OneDrive - Xclusive Trading Inc\Desktop\Python\Calling Tree Dashboard\CallingTree_Dashboard\Calling_Trees"

# ------------------ MySQL connection ------------------
def get_connection():
    return mysql.connector.connect(
        host='192.168.18.214',
        user='Mahee',
        password='Muhammadmahee123.',
        database='callingtree'
    )

# ------------------ Doorcodes ------------------
def create_doorcodes_table():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS doorcodes;")
    cursor.execute("""
        CREATE TABLE doorcodes (
            Door_code BIGINT PRIMARY KEY,
            Market VARCHAR(255),
            Store VARCHAR(255),
            DM VARCHAR(255),
            RSM VARCHAR(255),
            MD VARCHAR(255)
        );
    """)
    conn.commit()
    conn.close()
    print("doorcodes table is ready ✅")

def update_doorcodes():
    if not os.path.exists(DOORCODE_FILE):
        print(f"⚠️ File not found: {DOORCODE_FILE}")
        return

    try:
        df = pd.read_excel(DOORCODE_FILE)
    except PermissionError:
        print(f"⚠️ Permission denied: {DOORCODE_FILE} is open. Skipping...")
        return

    df['RSM'] = df['RSM'].replace('-', None)

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM doorcodes")
    conn.commit()

    for _, row in df.iterrows():
        cursor.execute(
            "INSERT INTO doorcodes (Door_code, Market, Store, DM, RSM, MD) VALUES (%s, %s, %s, %s, %s, %s)",
            (int(row['Door_code']), row['Market'], row['Store'], row['DM'], row['RSM'], row['MD'])
        )

    conn.commit()
    conn.close()
    print(f"Updated {len(df)} rows in doorcodes ✅")

# ------------------ NTIDs ------------------
def create_ntids_table():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DROP TABLE IF EXISTS ntids;")
    cursor.execute("""
        CREATE TABLE ntids (
            NTID VARCHAR(20) PRIMARY KEY,
            Name VARCHAR(255),  
            DESIGNATION VARCHAR(50)
        );
    """)
    conn.commit()
    conn.close()
    print("ntids table is ready ✅")

def update_ntids():
    if not os.path.exists(NTID_FILE):
        print(f"⚠️ File not found: {NTID_FILE}")
        return

    try:
        df = pd.read_excel(NTID_FILE)
    except PermissionError:
        print(f"⚠️ Permission denied: {NTID_FILE} is open. Skipping...")
        return

    # Fill missing NTID or DESIGNATION with placeholder
    df['NTIDs'] = df['NTIDs'].fillna('-').str.strip()
    df['DESIGNATION'] = df['DESIGNATION'].fillna('-').str.strip()

    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("DELETE FROM ntids")
    conn.commit()

    for _, row in df.iterrows():
        cursor.execute(
            "INSERT INTO ntids (NTID, Name, DESIGNATION) VALUES (%s, %s, %s)",
            (row['NTIDs'], row['DM/TM/RM/Ops'], row['DESIGNATION'])
        )

    conn.commit()
    conn.close()
    print(f"Updated {len(df)} rows in ntids ✅")

# ------------------ Calling Trees ------------------
def update_all_calling_trees_db_only():
    files = glob.glob(os.path.join(CALLING_TREE_FOLDER, "*.xlsx"))
    if not files:
        print("No Excel files found in folder.")
        return

    conn = get_connection()
    cursor = conn.cursor()

    for local_file in files:
        file_name = os.path.basename(local_file)

        try:
            with open(local_file, "rb") as f:
                local_data = f.read()
        except PermissionError:
            print(f"⚠️ Permission denied: {file_name} is open. Skipping...")
            continue

        cursor.execute("SELECT LENGTH(data_blob) as size FROM calling_trees WHERE name = %s", (file_name,))
        row = cursor.fetchone()

        if row:
            db_size = row[0]
            if len(local_data) != db_size:
                cursor.execute(
                    "UPDATE calling_trees SET data_blob = %s, last_updated = NOW() WHERE name = %s",
                    (local_data, file_name)
                )
                conn.commit()
                print(f"Updated {file_name} in database ✅")
            else:
                print(f"{file_name} is already up-to-date in database.")
        else:
            cursor.execute(
                "INSERT INTO calling_trees (name, data_blob, last_updated) VALUES (%s, %s, NOW())",
                (file_name, local_data)
            )
            conn.commit()
            print(f"Inserted {file_name} into database ✅")

    conn.close()

# ------------------ Manual Input ------------------
def wait_for_enter_or_esc(timeout=60):
    print(f"\nPress Enter to update immediately or ESC to quit. Waiting {timeout} seconds...")
    start_time = time.time()
    while True:
        if msvcrt.kbhit():
            key = msvcrt.getch()
            if key == b'\r':  # Enter key
                return "enter"
            elif key == b'\x1b':  # ESC key
                return "esc"
        if time.time() - start_time >= timeout:
            return "timeout"
        time.sleep(0.1)

# ------------------ Main Loop ------------------
def main_loop():
    create_doorcodes_table()
    create_ntids_table()
    last_calling_tree_update = 0
    CALLING_TREE_INTERVAL = 60  # 1 minute

    while True:
        now = time.time()
        print(f"\nRunning Doorcodes update at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        update_doorcodes()
        print(f"\nRunning NTIDs update at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        update_ntids()

        # Run Calling Trees update once per hour
        if now - last_calling_tree_update >= CALLING_TREE_INTERVAL:
            print(f"\nRunning Calling Trees update at {time.strftime('%Y-%m-%d %H:%M:%S')}")
            update_all_calling_trees_db_only()
            last_calling_tree_update = now

        # Wait for Enter, ESC, or timeout (60 sec)
        result = wait_for_enter_or_esc(timeout=60)
        if result == "enter":
            print("\nRunning all updates immediately ✅")
            update_doorcodes()
            update_ntids()
            update_all_calling_trees_db_only()
            last_calling_tree_update = time.time()
        elif result == "esc":
            print("\nProgram stopped by user ✅")
            break
        elif result == "timeout":
            print("\n60 seconds passed. Running next scheduled update ✅")
            continue

if __name__ == "__main__":
    main_loop()
