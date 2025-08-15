#!/usr/bin/env python3
"""
Check exact column positions for Instructions 3.0 accuracy
"""

import pandas as pd

def check_column_positions():
    """Check exact column positions for Instructions 3.0."""
    
    file_path = '../2025-0731 Production Summary FINAL.xlsx'
    
    print("🔍 **CHECKING COLUMN POSITIONS FOR INSTRUCTIONS 3.0**")
    print("=" * 70)
    
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name='Data', header=None)
        
        # Row 12 (index 12) contains the actual column headers
        header_row = df.iloc[12]
        
        print(f"📊 **Total columns:** {len(header_row)}")
        print(f"📊 **Data rows:** {len(df) - 13}")
        
        # Find Transaction Type column (should be Column J)
        print(f"\n🎯 **TRANSACTION TYPE COLUMN (Column J):**")
        for i, header in enumerate(header_row):
            if 'TRANSACTION' in str(header).upper() and 'TYPE' in str(header).upper():
                print(f"  Position {i} (Column {chr(65 + i)}): {header}")
                # Check if this is actually Column J (position 9)
                if i == 9:
                    print(f"  ✅ CORRECT: This is Column J (position 9)")
                else:
                    print(f"  ❌ INCORRECT: Expected Column J (position 9), found at position {i}")
        
        # Find Admin columns and their positions
        print(f"\n🎯 **ADMIN COLUMNS AND POSITIONS:**")
        admin_cols = []
        for i, header in enumerate(header_row):
            if 'ADMIN' in str(header).upper():
                admin_cols.append((i, header))
        
        for pos, col_name in admin_cols:
            col_letter = chr(65 + pos)
            print(f"  Position {pos} (Column {col_letter}): {col_name}")
        
        # Check specific Admin 6, 7, 8 positions
        print(f"\n🎯 **ADMIN 6, 7, 8 SPECIFIC CHECK:**")
        admin_6_pos = None
        admin_7_pos = None
        admin_8_pos = None
        
        for pos, col_name in admin_cols:
            if 'ADMIN 6' in str(col_name).upper() and 'AMOUNT' in str(col_name).upper():
                admin_6_pos = pos
                print(f"  Admin 6 Amount: Position {pos} (Column {chr(65 + pos)})")
            elif 'ADMIN 7' in str(col_name).upper() and 'AMOUNT' in str(col_name).upper():
                admin_7_pos = pos
                print(f"  Admin 7 Amount: Position {pos} (Column {chr(65 + pos)})")
            elif 'ADMIN 8' in str(col_name).upper() and 'AMOUNT' in str(col_name).upper():
                admin_8_pos = pos
                print(f"  Admin 8 Amount: Position {pos} (Column {chr(65 + pos)})")
        
        # Check if positions match Instructions 3.0
        print(f"\n🎯 **INSTRUCTIONS 3.0 POSITION VERIFICATION:**")
        print(f"  Instructions say Admin 6 Amount should be in AU (position 46)")
        print(f"  Instructions say Admin 7 Amount should be in AW (position 48)")
        print(f"  Instructions say Admin 8 Amount should be in AY (position 50)")
        
        if admin_6_pos == 46:
            print(f"  ✅ Admin 6 Amount: CORRECT position {admin_6_pos} (Column AU)")
        else:
            print(f"  ❌ Admin 6 Amount: WRONG position {admin_6_pos}, expected 46 (Column AU)")
            
        if admin_7_pos == 48:
            print(f"  ✅ Admin 7 Amount: CORRECT position {admin_7_pos} (Column AW)")
        else:
            print(f"  ❌ Admin 7 Amount: WRONG position {admin_7_pos}, expected 48 (Column AW)")
            
        if admin_8_pos == 50:
            print(f"  ✅ Admin 8 Amount: CORRECT position {admin_8_pos} (Column AY)")
        else:
            print(f"  ❌ Admin 8 Amount: WRONG position {admin_8_pos}, expected 50 (Column AY)")
        
        # Check other key columns
        print(f"\n🎯 **OTHER KEY COLUMNS:**")
        key_columns = ['Insurer Code', 'Product Type Code', 'Coverage Code', 'Dealer Number', 
                      'Dealer Name', 'Contract Number', 'Contract Sale Date', 'Customer Last Name',
                      'Term Months', 'Cancellation Date', 'Number of Days in Force', 
                      'Cancellation Factor', 'Cancellation Reason']
        
        for key_col in key_columns:
            for i, header in enumerate(header_row):
                if key_col.upper() in str(header).upper():
                    print(f"  {key_col}: Position {i} (Column {chr(65 + i)})")
                    break
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_column_positions()
