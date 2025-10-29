I've created a comprehensive Python script for Google Colab that handles your stock picking system. Here's what the code does:
Key Features:
1. File Upload & Loading

Accepts Excel file upload through Google Colab
Loads all sheets and preserves unused sheets

2. Smart Picking Algorithm

Case 1: Exact match → Records in Out-stock, removes from Stock-In-Hand
Case 2: Stock > Required → Records picked qty in Out-stock, subtracts from Stock-In-Hand
Case 3: Stock < Required → Picks all available, continues loop for more stock
Case 4: Still insufficient → Records shortage in Not-Available sheet

3. Data Recording

Transfers all relevant data (HS Code, Origin, Location, etc.)
Adds REFERENCE, D/NO, Date, Mail Reference from New-Order
Leaves Batch/Case empty for manual entry
Tracks total picked vs required quantities

4. Clean Operations

Removes zero-quantity lines from Stock-In-Hand
Updates Sl No for all sheets
Clears New-Order sheet after processing
Preserves all other sheets untouched

5. Download

Saves with timestamp: Master-Stock_YYYYMMDD_HHMMSS.xlsx
Automatically downloads the updated file

To Use:

Copy the code to Google Colab
Run the cell
Upload your Master-Stock Excel file when prompted
The script will process and download the updated file

The code includes detailed console output showing exactly what's being picked, any shortages, and a complete summary of the operation!
