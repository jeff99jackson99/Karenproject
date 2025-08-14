import pandas as pd
import numpy as np

def check_column_structure():
    """Check the exact column structure to find where each required column is located."""
    
    file_path = '../2025-0731 Production Summary FINAL.xlsx'
    
    print("🔍 **Exact Column Structure Check...**")
    
    try:
        # Read with no headers to see the raw structure
        df = pd.read_excel(file_path, sheet_name='Data', header=None)
        print(f"📊 **Raw shape:** {df.shape}")
        
        # Check Row 12 (the header row) for all columns
        print("\n🔍 **Row 12 (Header Row) - All Columns:**")
        header_row = df.iloc[12]
        for i, val in enumerate(header_row):
            if pd.notna(val) and str(val).strip() != '':
                print(f"  Position {i}: {val}")
        
        # Check specific columns mentioned in Karen 2.0
        print("\n🔍 **Karen 2.0 Required Columns - Exact Locations:**")
        karen_requirements = {
            'B': 'Insurer Code',
            'C': 'Product Type Code', 
            'D': 'Coverage Code',
            'E': 'Dealer Number',
            'F': 'Dealer Name',
            'H': 'Contract Number',
            'L': 'Contract Sale Date',
            'J': 'Transaction Date',
            'M': 'Transaction Type',
            'U': 'Customer Last Name',
            'Z': 'Contract Term',
            'AA': 'Cancellation Factor',
            'AB': 'Cancellation Reason',
            'AE': 'Cancellation Date'
        }
        
        for col_letter, expected_name in karen_requirements.items():
            found = False
            for i, val in enumerate(header_row):
                if pd.notna(val):
                    val_str = str(val).upper()
                    expected_upper = expected_name.upper()
                    
                    # Check for exact or close matches
                    if expected_upper in val_str or val_str in expected_upper:
                        print(f"  ✅ **{col_letter} ({expected_name}):** {val} at Position {i}")
                        found = True
                        break
            
            if not found:
                print(f"  ❌ **{col_letter} ({expected_name}):** NOT FOUND")
        
        # Check Admin columns specifically
        print("\n🔍 **Admin Columns - Exact Locations:**")
        admin_columns = []
        for i, val in enumerate(header_row):
            if pd.notna(val):
                val_str = str(val).upper()
                if 'ADMIN' in val_str and 'AMOUNT' in val_str:
                    admin_columns.append((i, val))
        
        for pos, val in admin_columns:
            print(f"  Position {pos}: {val}")
        
        # Check transaction types in the data
        print("\n🔍 **Transaction Types in Data Rows:**")
        # Look at rows 13+ (actual data)
        for row_idx in range(13, min(20, len(df))):
            transaction_col = df.iloc[row_idx, 9]  # Column 9
            if pd.notna(transaction_col):
                print(f"  Row {row_idx}, Col 9: {transaction_col}")
        
        # Check if there are multiple transaction type columns
        print("\n🔍 **Looking for Transaction Type columns in headers:**")
        for i, val in enumerate(header_row):
            if pd.notna(val):
                val_str = str(val).upper()
                if 'TRANSACTION' in val_str and 'TYPE' in val_str:
                    print(f"  Position {i}: {val}")
                    
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_column_structure()
