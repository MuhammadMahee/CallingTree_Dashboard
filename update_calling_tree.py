import mysql.connector
import os
import glob
import time
import msvcrt

# Folder where your Calling Tree Excel files are stored
LOCAL_FOLDER = r"C:\Users\MuhammadMahee\OneDrive - Xclusive Trading Inc\Desktop\Python\Calling Tree Dashboard\CallingTree_Dashboard\Calling_Trees"

def get_connection():
    return mysql.connector.connect(
        host='192.168.18.214',
        user='Mahee',
        password='Muhammadmahee123.',
        database='callingtree'
    )

def update_all_calling_trees_db_only():
    files = glob.glob(os.path.join(LOCAL_FOLDER, "*.xlsx"))
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
            print(f"⚠️ Permission denied: {file_name} is open in another program. Skipping...")
            continue  # skip this file and move to the next

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

def wait_for_enter_or_esc(timeout=3600):
    """Wait up to `timeout` seconds for Enter to run immediately or ESC to quit."""
    print("\nPress Enter to run update immediately or ESC to quit. Waiting up to 1 hour...")
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

def main_loop():
    while True:
        print(f"\nRunning update at {time.strftime('%Y-%m-%d %H:%M:%S')}")
        update_all_calling_trees_db_only()

        while True:
            result = wait_for_enter_or_esc(timeout=60)  # wait up to 1 hour
            if result == "enter":
                print("\nRunning update immediately ✅")
                update_all_calling_trees_db_only()
            elif result == "esc":
                print("\nProgram stopped by user ✅")
                return
            elif result == "timeout":
                print("\n1 hour passed. Running scheduled update ✅")
                break

if __name__ == "__main__":
    main_loop()
