#!/usr/bin/env python3
"""
Debug script to check NCB column values in the Excel file
"""

import pandas as pd
import numpy as np
import sys

def debug_ncb_values():
    """Debug NCB column values in the Excel file."""
    print("🔍 DEBUGGING NCB COLUMN VALUES")
    print("=" * 50)
    
    try:
        # Read the Excel file
        file_path = "../2025-0731 Production Summary FINAL.xlsx"
        print(f"📁 Reading file: {file_path}")
        
        # Read without header first to see the structure
        df_raw = pd.read_excel(file_path, sheet_name='Data', header=None)
        print(f"📊 Raw data shape: {df_raw.shape}")
        
        # Check row 12 (index 12) for headers
        header_row = df_raw.iloc[12]
        print(f"🔍 Header row (Row 12): {header_row[:30].tolist()}")
        
        # Look for Admin columns
        admin_cols = []
        admin_indices = []
        for i, col in enumerate(header_row):
            if 'ADMIN' in str(col) and 'Amount' in str(col):
                admin_cols.append(col)
                admin_indices.append(i)
                print(f"✅ Found Admin column: {col} at index {i}")
        
        print(f"\n📊 Found {len(admin_cols)} Admin columns")
        
        if len(admin_cols) >= 7:
            # Get sample data from Admin columns
            print(f"\n🔍 Sample data from Admin columns:")
            
            for i, (col_name, col_idx) in enumerate(zip(admin_cols, admin_indices)):
                # Get first 10 non-null values
                values = df_raw.iloc[13:, col_idx].dropna().head(10)
                print(f"\n  {col_name} (index {col_idx}):")
                print(f"    Sample values: {values.tolist()}")
                
                # Check data types
                try:
                    numeric_vals = pd.to_numeric(values, errors='coerce')
                    non_null_numeric = numeric_vals.dropna()
                    if len(non_null_numeric) > 0:
                        print(f"    Numeric range: {non_null_numeric.min():.2f} to {non_null_numeric.max():.2f}")
                        print(f"    Non-zero count: {(non_null_numeric != 0).sum()}")
                    else:
                        print(f"    No valid numeric values")
                except:
                    print(f"    Could not convert to numeric")
            
            # Test NCB sum calculation on first few rows
            print(f"\n🔍 Testing NCB sum calculation on first 5 rows:")
            
            # Get the required Admin columns (3, 4, 6, 7, 8, 9, 10)
            required_admin = ['ADMIN 3 Amount', 'ADMIN 4 Amount', 'ADMIN 6 Amount', 
                            'ADMIN 7 Amount', 'ADMIN 8 Amount', 'ADMIN 9 Amount', 'ADMIN 10 Amount']
            
            required_indices = []
            for admin_name in required_admin:
                if admin_name in admin_cols:
                    idx = admin_cols.index(admin_name)
                    required_indices.append(admin_indices[idx])
                else:
                    print(f"❌ Missing required Admin column: {admin_name}")
            
            if len(required_indices) == 7:
                print(f"✅ All 7 required Admin columns found")
                
                # Calculate NCB sum for first 5 rows
                for row_idx in range(13, 18):  # Rows 13-17 (first 5 data rows)
                    ncb_sum = 0
                    ncb_values = []
                    
                    for col_idx in required_indices:
                        value = df_raw.iloc[row_idx, col_idx]
                        try:
                            numeric_val = pd.to_numeric(value, errors='coerce')
                            if not pd.isna(numeric_val):
                                ncb_sum += numeric_val
                                ncb_values.append(numeric_val)
                            else:
                                ncb_values.append(0)
                        except:
                            ncb_values.append(0)
                    
                    print(f"  Row {row_idx}: NCB values = {ncb_values}, Sum = {ncb_sum:.2f}")
                    
            else:
                print(f"❌ Only found {len(required_indices)} of 7 required Admin columns")
                
        else:
            print(f"❌ Not enough Admin columns found: {len(admin_cols)} < 7")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_ncb_values()
