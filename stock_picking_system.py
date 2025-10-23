import pandas as pd
import numpy as np
from google.colab import files
from datetime import datetime
import io

def upload_master_stock():
    """Upload the Master Stock Excel file"""
    print("Please upload your Master-Stock Excel file...")
    uploaded = files.upload()
    filename = list(uploaded.keys())[0]
    return filename

def load_excel_file(filename):
    """Load all sheets from the Excel file"""
    print(f"\nLoading file: {filename}")
    xls = pd.ExcelFile(filename)
    all_sheets = {}
    
    for sheet_name in xls.sheet_names:
        all_sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        print(f"  - Loaded sheet: '{sheet_name}' with {len(all_sheets[sheet_name])} rows")
    
    return all_sheets

def process_picking(all_sheets):
    """Process the picking list based on New-Order sheet"""
    
    # Get the required sheets
    stock_in_hand = all_sheets['Stock-In-Hand'].copy()
    new_order = all_sheets['New-Order'].copy()
    out_stock = all_sheets['Out-stock'].copy()
    not_available = all_sheets['Not-Available'].copy()
    
    print("\n" + "="*80)
    print("STARTING PICKING PROCESS")
    print("="*80)
    
    # Process each order
    for order_idx, order_row in new_order.iterrows():
        part_number = order_row['Part Number']
        req_qty = order_row['Req-Qty']
        reference = order_row['REFERENCE']
        d_no = order_row['D/NO']
        date = order_row['Date']
        mail_ref = order_row['Mail Reference']
        
        print(f"\n{'='*80}")
        print(f"Processing Order #{order_idx + 1}")
        print(f"Part Number: {part_number} | Required Qty: {req_qty}")
        print(f"Reference: {reference} | D/NO: {d_no}")
        print(f"{'='*80}")
        
        total_picked = 0
        remaining_qty = req_qty
        
        # Find all matching part numbers in stock
        matching_indices = stock_in_hand[stock_in_hand['Part Number'] == part_number].index.tolist()
        
        if not matching_indices:
            print(f"  ⚠ No stock found for Part Number: {part_number}")
        
        # Process each matching stock line
        for stock_idx in matching_indices:
            if remaining_qty <= 0:
                break
            
            stock_qty = stock_in_hand.loc[stock_idx, 'Qty']
            
            print(f"\n  Stock Line Found:")
            print(f"    Index: {stock_idx} | Available Qty: {stock_qty}")
            
            if stock_qty == remaining_qty:
                # Case 1: Exact match
                print(f"    ✓ Case 1: Exact match - Picking {stock_qty} units")
                
                # Create out-stock record
                out_stock_record = stock_in_hand.loc[stock_idx].copy()
                out_stock_record['REFERENCE'] = reference
                out_stock_record['D/NO'] = d_no
                out_stock_record['Date'] = date
                out_stock_record['Mail Reference'] = mail_ref
                out_stock_record['Batch/ Case'] = ''
                
                # Append to out-stock
                out_stock = pd.concat([out_stock, pd.DataFrame([out_stock_record])], ignore_index=True)
                
                # Remove from stock (set Qty to 0 to avoid deletion issues)
                stock_in_hand.loc[stock_idx, 'Qty'] = 0
                
                total_picked += stock_qty
                remaining_qty = 0
                
            elif stock_qty > remaining_qty:
                # Case 2: Stock quantity is greater
                print(f"    ✓ Case 2: Partial pick - Picking {remaining_qty} units (leaving {stock_qty - remaining_qty})")
                
                # Create out-stock record with picked quantity
                out_stock_record = stock_in_hand.loc[stock_idx].copy()
                out_stock_record['Qty'] = remaining_qty
                out_stock_record['REFERENCE'] = reference
                out_stock_record['D/NO'] = d_no
                out_stock_record['Date'] = date
                out_stock_record['Mail Reference'] = mail_ref
                out_stock_record['Batch/ Case'] = ''
                
                # Append to out-stock
                out_stock = pd.concat([out_stock, pd.DataFrame([out_stock_record])], ignore_index=True)
                
                # Subtract from stock
                stock_in_hand.loc[stock_idx, 'Qty'] -= remaining_qty
                
                total_picked += remaining_qty
                remaining_qty = 0
                
            else:
                # Case 3: Stock quantity is less
                print(f"    ✓ Case 3: Insufficient stock - Picking all {stock_qty} units (still need {remaining_qty - stock_qty})")
                
                # Create out-stock record
                out_stock_record = stock_in_hand.loc[stock_idx].copy()
                out_stock_record['REFERENCE'] = reference
                out_stock_record['D/NO'] = d_no
                out_stock_record['Date'] = date
                out_stock_record['Mail Reference'] = mail_ref
                out_stock_record['Batch/ Case'] = ''
                
                # Append to out-stock
                out_stock = pd.concat([out_stock, pd.DataFrame([out_stock_record])], ignore_index=True)
                
                # Set stock to 0
                stock_in_hand.loc[stock_idx, 'Qty'] = 0
                
                total_picked += stock_qty
                remaining_qty -= stock_qty
        
        # Case 4: Not available quantity
        if remaining_qty > 0:
            print(f"\n  ⚠ NOT AVAILABLE: {remaining_qty} units could not be picked")
            
            not_available_record = {
                'Sl No': len(not_available) + 1,
                'Part Number': part_number,
                'Part Description': order_row['Part Description'],
                'Req-Qty': req_qty,
                'REFERENCE': reference,
                'D/NO': d_no,
                'Date': date,
                'Mail Reference': mail_ref,
                'NA-Qty': remaining_qty
            }
            
            not_available = pd.concat([not_available, pd.DataFrame([not_available_record])], ignore_index=True)
        
        print(f"\n  SUMMARY:")
        print(f"    Total Picked: {total_picked} units")
        print(f"    Not Available: {remaining_qty} units")
    
    # Remove lines with zero quantity from Stock-In-Hand
    stock_in_hand = stock_in_hand[stock_in_hand['Qty'] > 0].reset_index(drop=True)
    
    # Update Sl No for all sheets
    if len(stock_in_hand) > 0:
        stock_in_hand['Sl No'] = range(1, len(stock_in_hand) + 1)
    if len(out_stock) > 0:
        out_stock['Sl No'] = range(1, len(out_stock) + 1)
    if len(not_available) > 0:
        not_available['Sl No'] = range(1, len(not_available) + 1)
    
    # Clear New-Order sheet
    new_order = pd.DataFrame(columns=new_order.columns)
    
    # Update the sheets in all_sheets dictionary
    all_sheets['Stock-In-Hand'] = stock_in_hand
    all_sheets['New-Order'] = new_order
    all_sheets['Out-stock'] = out_stock
    all_sheets['Not-Available'] = not_available
    
    print("\n" + "="*80)
    print("PICKING PROCESS COMPLETED")
    print("="*80)
    print(f"Updated Stock-In-Hand: {len(stock_in_hand)} lines")
    print(f"Updated Out-stock: {len(out_stock)} lines")
    print(f"Updated Not-Available: {len(not_available)} lines")
    print(f"New-Order sheet cleared")
    
    return all_sheets

def save_excel_file(all_sheets):
    """Save all sheets to Excel file with timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"Master-Stock_{timestamp}.xlsx"
    
    print(f"\n{'='*80}")
    print(f"Saving updated file: {output_filename}")
    print(f"{'='*80}")
    
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        for sheet_name, df in all_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"  - Saved sheet: '{sheet_name}'")
    
    print(f"\nFile saved successfully!")
    print(f"Downloading: {output_filename}")
    files.download(output_filename)
    
    return output_filename

def main():
    """Main execution function"""
    print("="*80)
    print("STOCK PICKING SYSTEM - MASTER STOCK MANAGEMENT")
    print("="*80)
    
    try:
        # Step 1: Upload file
        filename = upload_master_stock()
        
        # Step 2: Load all sheets
        all_sheets = load_excel_file(filename)
        
        # Step 3: Verify required sheets exist
        required_sheets = ['Stock-In-Hand', 'New-Order', 'Out-stock', 'Not-Available']
        missing_sheets = [s for s in required_sheets if s not in all_sheets]
        
        if missing_sheets:
            print(f"\n❌ ERROR: Missing required sheets: {missing_sheets}")
            return
        
        # Step 4: Process picking
        all_sheets = process_picking(all_sheets)
        
        # Step 5: Save and download
        output_file = save_excel_file(all_sheets)
        
        print("\n" + "="*80)
        print("✓ PROCESS COMPLETED SUCCESSFULLY!")
        print("="*80)
        
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()

# Run the main function
if __name__ == "__main__":
    main()
