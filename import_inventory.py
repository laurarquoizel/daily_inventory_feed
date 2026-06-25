import os
from ftplib import FTP
import pandas as pd
from supabase import create_client

# Connect to FTP
ftp = FTP(os.environ["FTP_HOST"])
ftp.login(
    os.environ["FTP_USER"],
    os.environ["FTP_PASS"]
)

# Find newest Excel file
files = ftp.nlst()

excel_files = [
    f for f in files
    if f.lower().endswith((".xlsx", ".xls"))
]

if not excel_files:
    raise Exception("No Excel files found.")

latest_file = sorted(excel_files)[-1]

print(f"Downloading {latest_file}")

with open("inventory.xlsx", "wb") as f:
    ftp.retrbinary(f"RETR {latest_file}", f.write)

ftp.quit()

# Read Excel
df = pd.read_excel("inventory.xlsx")
df.columns = [
    c.strip().lower().replace(" ", "_")
    for c in df.columns
]

records = []

for _, row in df.iterrows():
    records.append({
        "sku": str(row.get("sku","")),
        "status": str(row.get("status","")),
        "currently_available": row.get("currently_available"),
        "future_available": row.get("future_available"),
        "future_date_available": None if pd.isna(row.get("future_date_available")) else str(row.get("future_date_available")).split(" ")[0],
        "wholesale_price": row.get("wholesale_price"),
        "imap": row.get("imap"),
        "promotional_price": row.get("promotional_price"),
        "upc": str(row.get("upc","")),
        "shipped_via": str(row.get("shipped_via","")),
        "dim_weight": row.get("dim_weight"),
        "future_status": str(row.get("future_status","")),
        "search_text": f"{row.get('sku','')} {row.get('upc','')}".lower()
    })

supabase = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SERVICE_KEY"]
)

supabase.table("daily_inventory").delete().neq("sku","").execute()
supabase.table("daily_inventory").insert(records).execute()

print(f"Imported {len(records)} records.")
