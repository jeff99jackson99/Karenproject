#!/usr/bin/env python3
"""
Verify Admin 6, 7, 8 columns in the actual Excel file
to ensure we're getting the correct data for Instructions 3.0
"""

import pandas as pd
import numpy as np

def verify_admin_columns():
    """Examine the actual Admin 6, 7, 8 columns in the Excel file."""
    
    # Path to the actual Excel file
    file_path = '../2025-0731 Production Summary FINAL.xlsx'
    
    print("🔍 **VERIFYING ADMIN 6, 7, 8 COLUMNS IN ACTUAL EXCEL FILE**")
    print("=" * 70)
    
    try:
        # Read the Excel file
        excel_data = pd.ExcelFile(file_path)
        print(f"📊 **Sheets found:** {excel_data.sheet_names}")
        
        if 'Data' in excel_data.sheet_names:
            print(f"\n📋 **Examining Data sheet...**")
            
            # Read with no headers to see the raw structure
            df = pd.read_excel(file_path, sheet_name='Data', header=None)
            print(f"📏 **Raw data shape:** {df.shape}")
            
            # Row 12 (index 12) contains the actual column headers
            header_row = df.iloc[12]
            print(f"\n🔍 **Header row (Row 12):**")
            for i, header in enumerate(header_row):
                if 'ADMIN' in str(header).upper():
                    print(f"  Position {i}: {header}")
            
            # Look specifically for Admin 6, 7, 8 columns
            print(f"\n🎯 **SPECIFIC SEARCH FOR ADMIN 6, 7, 8:**")
            
            admin_6_cols = []
            admin_7_cols = []
            admin_8_cols = []
            
            for i, header in enumerate(header_row):
                header_str = str(header).upper()
                if 'ADMIN 6' in header_str and 'AMOUNT' in header_str:
                    admin_6_cols.append((i, header))
                elif 'ADMIN 7' in header_str and 'AMOUNT' in header_str:
                    admin_7_cols.append((i, header))
                elif 'ADMIN 8' in header_str and 'AMOUNT' in header_str:
                    admin_8_cols.append((i, header))
            
            print(f"🔍 **Admin 6 Amount columns found:** {admin_6_cols}")
            print(f"🔍 **Admin 7 Amount columns found:** {admin_7_cols}")
            print(f"🔍 **Admin 8 Amount columns found:** {admin_8_cols}")
            
            # Now examine the actual data in these columns
            if admin_6_cols:
                print(f"\n📊 **EXAMINING ADMIN 6 AMOUNT DATA:**")
                for pos, col_name in admin_6_cols:
                    print(f"\n  Column: {col_name} (Position {pos})")
                    
                    # Get data from this column (skip header row)
                    col_data = df.iloc[13:, pos]
                    
                    # Show sample values
                    non_null_data = col_data.dropna()
                    print(f"    Total rows: {len(col_data)}")
                    print(f"    Non-null rows: {len(non_null_data)}")
                    print(f"    Null/None rows: {len(col_data) - len(non_null_data)}")
                    
                    if len(non_null_data) > 0:
                        print(f"    Sample values: {non_null_data.head(10).tolist()}")
                        
                        # Check for numeric values
                        numeric_data = pd.to_numeric(non_null_data, errors='coerce').dropna()
                        print(f"    Numeric values: {len(numeric_data)}")
                        if len(numeric_data) > 0:
                            print(f"    Numeric range: {numeric_data.min()} to {numeric_data.max()}")
                            print(f"    Negative values: {(numeric_data < 0).sum()}")
                            print(f"    Zero values: {(numeric_data == 0).sum()}")
                            print(f"    Positive values: {(numeric_data > 0).sum()}")
            
            if admin_7_cols:
                print(f"\n📊 **EXAMINING ADMIN 7 AMOUNT DATA:**")
                for pos, col_name in admin_7_cols:
                    print(f"\n  Column: {col_name} (Position {pos})")
                    
                    # Get data from this column (skip header row)
                    col_data = df.iloc[13:, pos]
                    
                    # Show sample values
                    non_null_data = col_data.dropna()
                    print(f"    Total rows: {len(col_data)}")
                    print(f"    Non-null rows: {len(non_null_data)}")
                    print(f"    Null/None rows: {len(col_data) - len(non_null_data)}")
                    
                    if len(non_null_data) > 0:
                        print(f"    Sample values: {non_null_data.head(10).tolist()}")
                        
                        # Check for numeric values
                        numeric_data = pd.to_numeric(non_null_data, errors='coerce').dropna()
                        print(f"    Numeric values: {len(numeric_data)}")
                        if len(numeric_data) > 0:
                            print(f"    Numeric range: {numeric_data.min()} to {numeric_data.max()}")
                            print(f"    Negative values: {(numeric_data < 0).sum()}")
                            print(f"    Zero values: {(numeric_data == 0).sum()}")
                            print(f"    Positive values: {(numeric_data > 0).sum()}")
            
            if admin_8_cols:
                print(f"\n📊 **EXAMINING ADMIN 8 AMOUNT DATA:**")
                for pos, col_name in admin_8_cols:
                    print(f"\n  Column: {col_name} (Position {pos})")
                    
                    # Get data from this column (skip header row)
                    col_data = df.iloc[13:, pos]
                    
                    # Show sample values
                    non_null_data = col_data.dropna()
                    print(f"    Total rows: {len(col_data)}")
                    print(f"    Non-null rows: {len(non_null_data)}")
                    print(f"    Null/None rows: {len(col_data) - len(non_null_data)}")
                    
                    if len(non_null_data) > 0:
                        print(f"    Sample values: {non_null_data.head(10).tolist()}")
                        
                        # Check for numeric values
                        numeric_data = pd.to_numeric(non_null_data, errors='coerce').dropna()
                        print(f"    Numeric values: {len(numeric_data)}")
                        if len(numeric_data) > 0:
                            print(f"    Numeric range: {numeric_data.min()} to {numeric_data.max()}")
                            print(f"    Negative values: {(numeric_data < 0).sum()}")
                            print(f"    Zero values: {(numeric_data == 0).sum()}")
                            print(f"    Positive values: {(numeric_data > 0).sum()}")
            
            # Also check for any other Admin columns that might have data
            print(f"\n🔍 **CHECKING ALL ADMIN COLUMNS FOR DATA:**")
            all_admin_cols = []
            for i, header in enumerate(header_row):
                header_str = str(header).upper()
                if 'ADMIN' in header_str and 'AMOUNT' in header_str:
                    all_admin_cols.append((i, header))
            
            for pos, col_name in all_admin_cols:
                print(f"\n  {col_name} (Position {pos}):")
                col_data = df.iloc[13:, pos]
                non_null_data = col_data.dropna()
                print(f"    Non-null: {len(non_null_data)}/{len(col_data)}")
                
                if len(non_null_data) > 0:
                    numeric_data = pd.to_numeric(non_null_data, errors='coerce').dropna()
                    if len(numeric_data) > 0:
                        print(f"    Numeric: {len(numeric_data)}, Range: {numeric_data.min():.2f} to {numeric_data.max():.2f}")
                        print(f"    Negative: {(numeric_data < 0).sum()}, Zero: {(numeric_data == 0).sum()}, Positive: {(numeric_data > 0).sum()}")
                    else:
                        print(f"    No numeric values found")
                else:
                    print(f"    All values are null/empty")
            
        else:
            print("❌ Data sheet not found!")
            
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    verify_admin_columns()
