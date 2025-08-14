import pandas as pd
import numpy as np

def examine_excel_structure():
    """Examine the actual Excel file structure to understand how to read it properly."""
    
    # Read the actual file from the parent directory
    file_path = '../2025-0731 Production Summary FINAL.xlsx'
    
    print("🔍 **Examining actual Excel file structure...**")
    
    try:
        # First, let's see what sheets exist
        excel_file = pd.ExcelFile(file_path)
        print(f"📊 **Sheets found:** {excel_file.sheet_names}")
        
        # Focus on the Data sheet
        if 'Data' in excel_file.sheet_names:
            print(f"\n📋 **Examining Data sheet...**")
            
            # Read with no headers to see the raw structure
            df = pd.read_excel(file_path, sheet_name='Data', header=None)
            print(f"  Raw shape: {df.shape}")
            
            # Look for Admin columns by examining more rows
            print("\n🔍 **Looking for Admin columns in more rows...**")
            for row_idx in range(20):  # Check first 20 rows
                row_data = df.iloc[row_idx, :50]  # Check first 50 columns
                admin_count = 0
                admin_positions = []
                
                for i, val in enumerate(row_data):
                    if pd.notna(val):
                        val_str = str(val).upper()
                        if 'ADMIN' in val_str and 'AMOUNT' in val_str:
                            admin_count += 1
                            admin_positions.append((i, val))
                
                if admin_count >= 3:
                    print(f"  ✅ **Row {row_idx} contains {admin_count} Admin columns:**")
                    for pos, val in admin_positions:
                        print(f"     Position {pos}: {val}")
                    print(f"     Sample row data: {row_data[:20].tolist()}")
                    break
            
            # Look for transaction types in more detail
            print("\n🔍 **Looking for transaction types in detail...**")
            transaction_col = None
            for col_idx in range(20):  # Check first 20 columns
                col_data = df.iloc[10:30, col_idx].dropna()  # Check rows 10-30
                if len(col_data) > 0:
                    str_values = [str(v).upper() for v in col_data]
                    nb_count = sum(1 for v in str_values if v in ['NB', 'NEW', 'NEW BUSINESS'])
                    c_count = sum(1 for v in str_values if v in ['C', 'CANCEL', 'CANCELLATION'])
                    r_count = sum(1 for v in str_values if v in ['R', 'REINSTATE', 'REINSTATEMENT'])
                    
                    if nb_count > 0 or c_count > 0 or r_count > 0:
                        print(f"  ✅ **Column {col_idx} contains transaction types:**")
                        print(f"     NB: {nb_count}, C: {c_count}, R: {r_count}")
                        print(f"     Sample values: {str_values[:10]}")
                        if not transaction_col:
                            transaction_col = col_idx
            
            # Look for numeric columns that might be Admin amounts
            print("\n🔍 **Looking for numeric columns that might be Admin amounts...**")
            numeric_columns = []
            for col_idx in range(100):  # Check first 100 columns
                try:
                    col_data = df.iloc[15:50, col_idx].dropna()  # Check rows 15-50
                    if len(col_data) > 0:
                        numeric_vals = pd.to_numeric(col_data, errors='coerce')
                        non_null_numeric = numeric_vals.dropna()
                        if len(non_null_numeric) > 0:
                            non_zero = (non_null_numeric != 0).sum()
                            if non_zero > 0:
                                numeric_columns.append((col_idx, non_zero, len(non_null_numeric)))
                except:
                    continue
            
            # Sort by number of non-zero values
            numeric_columns.sort(key=lambda x: x[1], reverse=True)
            
            print(f"🔍 **Found {len(numeric_columns)} numeric columns with non-zero values:**")
            for col_idx, non_zero, total in numeric_columns[:20]:
                print(f"  Position {col_idx}: {non_zero}/{total} non-zero values")
                
        else:
            print("❌ Data sheet not found!")
            
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    examine_excel_structure()
