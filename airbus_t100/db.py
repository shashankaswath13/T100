# import pandas as pd
# import sqlite3
# import os

# def excel_to_sqlite(excel_path, db_path=None):
#     """
#     Convert an Excel file into a SQLite database.
#     Each sheet will be stored as a separate table.
    
#     :param excel_path: Path to the Excel file
#     :param db_path: Path for the output SQLite DB file (default: same name as Excel)
#     """
#     # Default DB file name if not provided
#     if db_path is None:
#         db_path = os.path.splitext(excel_path)[0] + ".db"
    
#     # Load Excel file
#     xls = pd.ExcelFile(excel_path)
    
#     # Create SQLite connection
#     conn = sqlite3.connect(db_path)
    
#     # Loop through all sheets and save as tables
#     for sheet_name in xls.sheet_names:
#         df = pd.read_excel(xls, sheet_name=sheet_name)
#         # Replace spaces in sheet names to make valid table names
#         table_name = sheet_name.replace(" ", "_")
#         df.to_sql(table_name, conn, if_exists="replace", index=False)
#         print(f"✔ Saved sheet '{sheet_name}' as table '{table_name}'")
    
#     conn.close()
#     print(f"\n✅ Excel file converted successfully to SQLite DB: {db_path}")


# # Example usage
# excel_to_sqlite(r"C:\Users\shashank.aswath\Documents\airbus_t100\Hole_codes_and_fasteners.xlsx", r"C:\Users\shashank.aswath\Documents\airbus_t100\details.db")



import sqlite3
import pandas as pd

# Path to your DB file
# db_path = r"C:\Users\shashank.aswath\Downloads\temp\temp\Hole codes and fasteners.db"
db_path = r"C:\Users\shashank.aswath\Documents\airbus_t100\details.db"
# Connect to the database
conn = sqlite3.connect(db_path)

# Get all table names
tables = pd.read_sql("SELECT name FROM sqlite_master WHERE type='table';", conn)
print("Tables in database:\n", tables, "\n")

# Loop through each table and print all its data
for table_name in tables['name']:
    print(f"\n=== Table: {table_name} ===")
    df = pd.read_sql(f"SELECT * FROM {table_name};", conn)
    print(df)

conn.close()



# import sqlite3

# def update_fastener1(db_path, hole_code, collar3):
#     try:
#         conn = sqlite3.connect(db_path)
#         cur = conn.cursor()

#         # ✅ Check if the HOLE CODE exists
#         cur.execute('SELECT COUNT(*) FROM Sheet1 WHERE "Hole code1" = ?', (hole_code,))
#         exists = cur.fetchone()[0]

#         if exists:
#             # ✅ Update Fastener 1
#             cur.execute('UPDATE Sheet1 SET "Fastener 1" = ? WHERE "Hole code1" = ?',
#                         (collar3, hole_code))
#             print(f"✅ Updated Fastener 1 for {hole_code} to {collar3}")
#         else:
#             # ✅ Insert new row if needed
#             cur.execute('INSERT INTO Sheet1 ("Hole code1", "Collar3") VALUES (?, ?)',
#                         (hole_code, collar3))
#             print(f"✅ Inserted new row: {hole_code} -> Fastener 1 = {collar3}")

#         conn.commit()
#         conn.close()
#     except Exception as e:
#         print(f"❌ Error: {e}")


# # Example usage
# db_path = r"C:\Users\shashank.aswath\Documents\airbus_t100\details.db"
# hole_code = "ABS1707BP1V3A"
# collar3 = "EN6115B3E"

# update_fastener1(db_path, hole_code, collar3)


