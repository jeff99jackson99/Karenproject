import pandas as pd
import numpy as np

def deep_examine_excel():
    """Deep examination of the Excel file to find the correct Admin columns."""
    
    file_path = '../2025-0731 Production Summary FINAL.xlsx'
    
    print("🔍 **Deep Examination of Excel File Structure...**")
    
    try:
        # Read with no headers to see the raw structure
        df = pd.read_excel(file_path, sheet_name='Data', header=None)
        print(f"📊 **Raw shape:** {df.shape}")
        
        # Look for Admin columns in ALL rows
        print("\n🔍 **Searching ALL rows for Admin columns...**")
        for row_idx in range(50):  # Check first 50 rows
            row_data = df.iloc[row_idx, :100]  # Check first 100 columns
            admin_columns = []
            
            for i, val in enumerate(row_data):
                if pd.notna(val):
                    val_str = str(val).upper()
                    if 'ADMIN' in val_str and 'AMOUNT' in val_str:
                        admin_columns.append((i, val))
            
            if len(admin_columns) >= 3:
                print(f"  ✅ **Row {row_idx} contains {len(admin_columns)} Admin columns:**")
                for pos, val in admin_columns:
                    print(f"     Position {pos}: {val}")
                
                # Check if this row has the specific Admin columns we need
                admin_texts = [val.upper() for _, val in admin_columns]
                has_admin_3 = any('ADMIN 3' in text for text in admin_texts)
                has_admin_4 = any('ADMIN 4' in text for text in admin_texts)
                has_admin_6 = any('ADMIN 6' in text for text in admin_texts)
                has_admin_7 = any('ADMIN 7' in text for text in admin_texts)
                has_admin_8 = any('ADMIN 8' in text for text in admin_texts)
                has_admin_9 = any('ADMIN 9' in text for text in admin_texts)
                has_admin_10 = any('ADMIN 10' in text for text in admin_texts)
                
                print(f"     Karen 2.0 Match Check:")
                print(f"       ADMIN 3: {has_admin_3}")
                print(f"       ADMIN 4: {has_admin_4}")
                print(f"       ADMIN 6: {has_admin_6}")
                print(f"       ADMIN 7: {has_admin_7}")
                print(f"       ADMIN 8: {has_admin_8}")
                print(f"       ADMIN 9: {has_admin_9}")
                print(f"       ADMIN 10: {has_admin_10}")
                
                if has_admin_3 and has_admin_4 and has_admin_6 and has_admin_7 and has_admin_8 and has_admin_9 and has_admin_10:
                    print(f"     🎯 **PERFECT MATCH FOUND!** Row {row_idx} has all required Admin columns")
                    break
        
        # Look for transaction types in more detail
        print("\n🔍 **Examining transaction types in detail...**")
        for col_idx in range(20):
            col_data = df.iloc[10:50, col_idx].dropna()
            if len(col_data) > 0:
                str_values = [str(v).upper() for v in col_data]
                nb_count = sum(1 for v in str_values if v in ['NB', 'NEW', 'NEW BUSINESS'])
                c_count = sum(1 for v in str_values if v in ['C', 'CANCEL', 'CANCELLATION'])
                r_count = sum(1 for v in str_values if v in ['R', 'REINSTATE', 'REINSTATEMENT'])
                
                if nb_count > 0 or c_count > 0 or r_count > 0:
                    print(f"  ✅ **Column {col_idx} contains transaction types:**")
                    print(f"     NB: {nb_count}, C: {c_count}, R: {r_count}")
                    print(f"     Sample values: {str_values[:10]}")
        
        # Look for the specific columns mentioned in Karen 2.0
        print("\n🔍 **Looking for Karen 2.0 specific columns...**")
        karen_columns = {
            'B': ['INSURER', 'INSURER CODE'],
            'C': ['PRODUCT TYPE', 'PRODUCT TYPE CODE'],
            'D': ['COVERAGE', 'COVERAGE CODE'],
            'E': ['DEALER NUMBER'],
            'F': ['DEALER NAME'],
            'H': ['CONTRACT NUMBER'],
            'L': ['CONTRACT SALE DATE', 'SALE DATE'],
            'J': ['TRANSACTION DATE'],
            'M': ['TRANSACTION TYPE'],
            'U': ['CUSTOMER LAST NAME', 'LAST NAME'],
            'Z': ['CONTRACT TERM', 'TERM'],
            'AA': ['CANCELLATION FACTOR', 'DAYS IN FORCE'],
            'AB': ['CANCELLATION REASON'],
            'AE': ['CANCELLATION DATE']
        }
        
        for col_letter, search_terms in karen_columns.items():
            found = False
            for row_idx in range(20):
                row_data = df.iloc[row_idx, :50]
                for i, val in enumerate(row_data):
                    if pd.notna(val):
                        val_str = str(val).upper()
                        if any(term in val_str for term in search_terms):
                            print(f"  ✅ **{col_letter} found:** {val} at Row {row_idx}, Col {i}")
                            found = True
                            break
                if found:
                    break
            
            if not found:
                print(f"  ❌ **{col_letter} NOT FOUND**")
                
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    deep_examine_excel()
