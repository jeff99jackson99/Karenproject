#!/usr/bin/env python3
"""
Debug script to check required column finding
"""

import pandas as pd
import sys

def debug_required_columns():
    """Debug required column finding."""
    print("🔍 DEBUGGING REQUIRED COLUMN FINDING")
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
        
        # Look for required columns
        required_columns = {
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
            'Z': 'Term Months',
            'AA': 'Cancellation Factor',
            'AB': 'Cancellation Reason',
            'AE': 'Cancellation Date'
        }
        
        print(f"\n🔍 Looking for required columns:")
        found_columns = {}
        
        for col_letter, col_name in required_columns.items():
            # Look for exact match first
            if col_name in header_row.values:
                col_idx = list(header_row.values).index(col_name)
                found_columns[col_letter] = col_name
                print(f"✅ {col_letter}: {col_name} found at index {col_idx}")
            else:
                # Look for partial match
                found = False
                for i, col in enumerate(header_row):
                    if col_name.lower() in str(col).lower():
                        found_columns[col_letter] = col
                        print(f"✅ {col_letter}: {col} (partial match for {col_name}) at index {i}")
                        found = True
                        break
                
                if not found:
                    print(f"❌ {col_letter}: {col_name} NOT FOUND")
        
        print(f"\n📊 Found {len(found_columns)} of {len(required_columns)} required columns")
        
        # Check if we have enough columns
        if len(found_columns) >= 10:
            print("✅ Sufficient required columns found")
        else:
            print(f"❌ Insufficient required columns: {len(found_columns)} < 10")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_required_columns()
