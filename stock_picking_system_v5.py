# version: v05 
# Date: 10-Nov-2025
# Author: Blenson & Claude.ai
# The zero order lines removed from the generated out file.
# Sl No maintained as per the New order.

# clear temp / previously loaded excel files in case code has previously run
!rm *.xlsx


import pandas as pd
import numpy as np
from datetime import datetime
from google.colab import files
import io

# Upload the Master-stock Excel file
print("Please upload your Master-stock Excel file...")
uploaded = files.upload()

# Get the uploaded file name
filename = list(uploaded.keys())[0]
print(f"\nProcessing file: {filename}")

# Read all sheets from the Excel file
excel_file = pd.ExcelFile(io.BytesIO(uploaded[filename]))
all_sheets = {}

# Load all sheets
for sheet_name in excel_file.sheet_names:
    all_sheets[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name)
    print(f"Loaded sheet: {sheet_name}")

# Verify required sheets exist
required_sheets = ["Stock-In-Hand", "New-Order", "Out-stock", "Not-Available"]
for sheet in required_sheets:
    if sheet not in all_sheets:
        print(f"ERROR: Required sheet '{sheet}' not found!")
        raise ValueError(f"Missing required sheet: {sheet}")

# Get references to working sheets
stock_in_hand = all_sheets["Stock-In-Hand"].copy()
new_order = all_sheets["New-Order"].copy()
out_stock = all_sheets["Out-stock"].copy()
not_available = all_sheets["Not-Available"].copy()

print(f"\nStock-In-Hand rows: {len(stock_in_hand)}")
print(f"New-Order rows: {len(new_order)}")
print(f"Out-stock existing rows: {len(out_stock)}")
print(f"Not-Available existing rows: {len(not_available)}")

# Initialize lists to store new records
new_out_stock_records = []
new_not_available_records = []

# Current date for recording
current_date = datetime.now().strftime("%Y.%m.%d")

# Process each order in New-Order sheet
print("\n" + "="*80)
print("PROCESSING PICKING LIST")
print("="*80)

for order_idx, order_row in new_order.iterrows():
    part_number = order_row["Part Number"]
    req_qty = order_row["Req-Qty"]
    reference = order_row["REFERENCE"]
    d_no = order_row["D/NO"]
    order_date = order_row["Date"]
    mail_ref = order_row["Mail Reference"]
    
    print(f"\n[Order {order_idx + 1}] Part: {part_number} | Required Qty: {req_qty}")
    
    # Skip if required quantity is 0 or NaN
    if pd.isna(req_qty) or req_qty <= 0:
        print(f"  ⚠ Skipping - Invalid required quantity")
        continue
    
    total_picked = 0
    remaining_qty = req_qty
    
    # Find all matching rows in Stock-In-Hand
    matching_indices = stock_in_hand[stock_in_hand["Part Number"] == part_number].index.tolist()
    
    if not matching_indices:
        print(f"  ✗ No stock found for this part number")
        # Record in Not-Available with full quantity
        new_not_available_records.append({
            "Sl No": order_row["Sl No"],  # Use Sl No from New-Order sheet
            "Part Number": part_number,
            "Part Description": order_row["Part Description"],
            "Req-Qty": req_qty,
            "REFERENCE": reference,
            "D/NO": d_no,
            "Date": order_date,
            "Mail Reference": mail_ref,
            "NA-Qty": req_qty
        })
        continue
    
    # Process each matching stock entry
    for stock_idx in matching_indices:
        if remaining_qty <= 0:
            break
        
        stock_row = stock_in_hand.loc[stock_idx]
        available_qty = stock_row["Qty"]
        
        # Skip if stock quantity is 0 or NaN
        if pd.isna(available_qty) or available_qty <= 0:
            continue
        
        # Determine quantity to pick
        if available_qty >= remaining_qty:
            # Case 1 & 2: Stock is sufficient
            pick_qty = remaining_qty
            stock_in_hand.loc[stock_idx, "Qty"] = available_qty - pick_qty
            print(f"  ✓ Picked {pick_qty} from stock (remaining in stock: {available_qty - pick_qty})")
        else:
            # Case 3: Stock is less than required
            pick_qty = available_qty
            stock_in_hand.loc[stock_idx, "Qty"] = 0
            print(f"  ⚠ Partially picked {pick_qty} from stock (stock depleted)")
        
        total_picked += pick_qty
        remaining_qty -= pick_qty
        
        # Create Out-stock record
        out_stock_record = {
            "Sl No": order_row["Sl No"],  # Use Sl No from New-Order sheet
            "Part Number": stock_row["Part Number"],
            "Part Description": stock_row["Part Description"],
            "Qty": pick_qty,
            "Unit Weight(gm)": stock_row["Unit Weight(gm)"],
            "HS Code": stock_row["HS Code"],
            "Origin": stock_row["Origin"],
            "Location": stock_row["Location"],
            "KME Case Ref": stock_row["KME Case Ref"],
            "Ref": stock_row["Ref"],
            "BO lookup": stock_row["BO lookup"],
            "SO Number": stock_row["SO Number"],
            "SRV Remarks": stock_row["SRV Remarks"],
            "Date (added) yyy.mm.dd": stock_row["Date (added) yyy.mm.dd"],
            "REFERENCE": reference,
            "D/NO": d_no,
            "Batch/ Case": "",  # To be filled manually
            "Date": order_date,
            "Mail Reference": mail_ref,
            "srv print": stock_row.get("srv print", "")
        }
        new_out_stock_records.append(out_stock_record)
    
    # Case 4: Check if there's still unmet demand
    if remaining_qty > 0:
        print(f"  ✗ Not available quantity: {remaining_qty}")
        new_not_available_records.append({
            "Sl No": order_row["Sl No"],  # Use Sl No from New-Order sheet
            "Part Number": part_number,
            "Part Description": order_row["Part Description"],
            "Req-Qty": req_qty,
            "REFERENCE": reference,
            "D/NO": d_no,
            "Date": order_date,
            "Mail Reference": mail_ref,
            "NA-Qty": remaining_qty
        })
    
    print(f"  → Total picked: {total_picked}/{req_qty}")

# Remove rows with zero quantity from Stock-In-Hand
print("\n" + "="*80)
print("CLEANING UP STOCK-IN-HAND")
print("="*80)
initial_stock_count = len(stock_in_hand)
stock_in_hand = stock_in_hand[stock_in_hand["Qty"] > 0].reset_index(drop=True)
removed_count = initial_stock_count - len(stock_in_hand)
print(f"Removed {removed_count} rows with zero quantity")
print(f"Remaining stock rows: {len(stock_in_hand)}")

# Append new records to Out-stock
if new_out_stock_records:
    out_stock_new = pd.DataFrame(new_out_stock_records)
    out_stock = pd.concat([out_stock, out_stock_new], ignore_index=True)
    print(f"\n✓ Added {len(new_out_stock_records)} records to Out-stock sheet")

# Append new records to Not-Available
if new_not_available_records:
    not_available_new = pd.DataFrame(new_not_available_records)
    not_available = pd.concat([not_available, not_available_new], ignore_index=True)
    print(f"✓ Added {len(new_not_available_records)} records to Not-Available sheet")

# Clear New-Order sheet
new_order = pd.DataFrame(columns=new_order.columns)
print(f"✓ Cleared New-Order sheet")

# Update the all_sheets dictionary
all_sheets["Stock-In-Hand"] = stock_in_hand
all_sheets["New-Order"] = new_order
all_sheets["Out-stock"] = out_stock
all_sheets["Not-Available"] = not_available

# Create output filename with timestamp
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_filename = f"Master-stock_{timestamp}.xlsx"

# Write all sheets back to Excel
print(f"\n" + "="*80)
print("SAVING UPDATED FILE")
print("="*80)
with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
    for sheet_name, df in all_sheets.items():
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f"✓ Saved sheet: {sheet_name}")

print(f"\n✓ File saved as: {output_filename}")
print(f"\n" + "="*80)
print("SUMMARY")
print("="*80)
print(f"Orders processed: {len(new_order_records) if 'new_order_records' in locals() else 'N/A'}")
print(f"Picked items recorded: {len(new_out_stock_records)}")
print(f"Not available items: {len(new_not_available_records)}")
print(f"Stock-In-Hand remaining: {len(stock_in_hand)} rows")
print("="*80)

# Download the file
files.download(output_filename)
print(f"\n✓ Download started for: {output_filename}")
