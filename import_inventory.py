import os
from ftplib import FTP
import pandas as pd
from supabase import create_client


FTP_FOLDER = "Inventory"


def clean_value(value):
    if pd.isna(value):
        return None
    return value


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_number(value):
    if pd.isna(value) or value == "":
        return None
    return float(value)


def clean_int(value):
    if pd.isna(value) or value == "":
        return 0
    return int(value)


def clean_date(value):
    if pd.isna(value) or value == "":
        return None
    return str(value).split(" ")[0]


# Connect to FTP
ftp = FTP(os.environ["FTP_HOST"])
ftp.login(os.environ["FTP_USER"], os.environ["FTP_PASS"])

# Move into Inventory folder
ftp.cwd(FTP_FOLDER)

# List files
files = ftp.nlst()

print("====================")
print("Current FTP folder:")
print(ftp.pwd())
print("====================")

print("Files found:")
for f in files:
    print(repr(f))
print("====================")

# Only use Quoizel daily inventory Excel files
excel_files = [
    f for f in files
    if f.lower().endswith((".xlsx", ".xls"))
    and "quoizeldailyinventoryfeed" in f.lower()
]

if not excel_files:
    raise Exception("No QuoizelDailyInventoryFeed Excel files found in FTP Inventory folder.")

# Pick the newest alphabetically
latest_file = sorted(excel_files)[-1]

print(f"Downloading: {latest_file}")

with open("inventory.xlsx", "wb") as f:
    ftp.retrbinary(f"RETR {latest_file}", f.write)

ftp.quit()


# Read Excel
df = pd.read_excel("inventory.xlsx")

# Normalize column names
df.columns = [
    c.strip().lower().replace(" ", "_")
    for c in df.columns
]

print("Columns found:")
for c in df.columns:
    print(c)


records = []

for _, row in df.iterrows():
    sku = clean_text(row.get("sku"))
    upc = clean_text(row.get("upc"))

    if not sku:
        continue

    records.append({
        "sku": sku,
        "status": clean_text(row.get("status")),
        "currently_available": clean_int(row.get("currently_available")),
        "future_available": clean_int(row.get("future_available")),
        "future_date_available": clean_date(row.get("future_date_available")),
        "wholesale_price": clean_number(row.get("wholesale_price")),
        "imap": clean_number(row.get("imap")),
        "promotional_price": clean_number(row.get("promotional_price")),
        "upc": upc,
        "shipped_via": clean_text(row.get("shipped_via")),
        "dim_weight": clean_number(row.get("dim_weight")),
        "future_status": clean_text(row.get("future_status")),
        "search_text": f"{sku} {upc}".lower()
    })


print(f"Prepared {len(records)} records for import.")


# Connect to Supabase
supabase = create_client(
    os.environ["SUPABASE_URL"],
    os.environ["SUPABASE_SERVICE_KEY"]
)

# Clear existing table
supabase.table("daily_inventory").delete().neq("sku", "").execute()

# Insert in batches
batch_size = 500

for i in range(0, len(records), batch_size):
    batch = records[i:i + batch_size]
    supabase.table("daily_inventory").insert(batch).execute()
    print(f"Inserted records {i + 1} to {i + len(batch)}")

print("Inventory import complete.")
