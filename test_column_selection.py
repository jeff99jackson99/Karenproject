#!/usr/bin/env python3
"""
Test script to verify column selection for output
"""

import pandas as pd
import sys

def test_column_selection():
    """Test column selection for output."""
    print("🧪 TESTING COLUMN SELECTION FOR OUTPUT")
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
        
        # Use row 12 as column names
        df = df_raw.copy()
        df.columns = header_row
        # Remove rows 0-12 (they're not data) and reset index
        df = df.iloc[13:].reset_index(drop=True)
        print(f"📏 Data shape after header fix: {df.shape}")
        
        # Mock the column finding functions
        ncb_columns = {
            'AO': 'ADMIN 3 Amount',
            'AQ': 'ADMIN 4 Amount',
            'AU': 'ADMIN 6 Amount',
            'AW': 'ADMIN 7 Amount',
            'AY': 'ADMIN 8 Amount',
            'BA': 'ADMIN 9 Amount',
            'BC': 'ADMIN 10 Amount'
        }
        
        required_cols = {
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
        
        print(f"✅ NCB columns: {len(ncb_columns)}")
        print(f"✅ Required columns: {len(required_cols)}")
        
        # Test column selection for NB output (17 columns)
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY'),
            ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        print(f"\n🔍 NB output columns ({len(nb_output_cols)}):")
        for i, col in enumerate(nb_output_cols):
            print(f"  {i+1:2d}: {col}")
        
        # Check for None values
        none_count = sum(1 for col in nb_output_cols if col is None)
        if none_count > 0:
            print(f"❌ {none_count} None values in NB column selection")
        else:
            print("✅ No None values in NB column selection")
        
        # Check for duplicates
        nb_duplicates = [col for col in set(nb_output_cols) if nb_output_cols.count(col) > 1]
        if nb_duplicates:
            print(f"❌ Duplicate columns in NB selection: {nb_duplicates}")
        else:
            print("✅ No duplicate columns in NB selection")
        
        # Filter to available columns
        available_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        print(f"✅ Available columns for NB output: {len(available_cols)}")
        
        if len(available_cols) >= 17:
            print("✅ Sufficient columns for NB output")
        else:
            print(f"❌ Insufficient columns: {len(available_cols)} < 17")
            missing_cols = [col for col in nb_output_cols if col is not None and col not in df.columns]
            print(f"❌ Missing columns: {missing_cols}")
        
        # Test C output columns (21 columns)
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), required_cols.get('Z'), required_cols.get('AA'),
            required_cols.get('AB'), required_cols.get('AE'), ncb_columns.get('AO'),
            ncb_columns.get('AQ'), ncb_columns.get('AU'), ncb_columns.get('AW'),
            ncb_columns.get('AY'), ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        print(f"\n🔍 C output columns ({len(c_output_cols)}):")
        for i, col in enumerate(c_output_cols):
            print(f"  {i+1:2d}: {col}")
        
        # Check for None values
        none_count = sum(1 for col in c_output_cols if col is None)
        if none_count > 0:
            print(f"❌ {none_count} None values in C column selection")
        else:
            print("✅ No None values in C column selection")
        
        # Check for duplicates
        c_duplicates = [col for col in set(c_output_cols) if c_output_cols.count(col) > 1]
        if c_duplicates:
            print(f"❌ Duplicate columns in C selection: {c_duplicates}")
        else:
            print("✅ No duplicate columns in C selection")
        
        # Filter to available columns
        available_cols = [col for col in c_output_cols if col is not None and col in df.columns]
        print(f"✅ Available columns for C output: {len(available_cols)}")
        
        if len(available_cols) >= 21:
            print("✅ Sufficient columns for C output")
        else:
            print(f"❌ Insufficient columns: {len(available_cols)} < 21")
            missing_cols = [col for col in c_output_cols if col is not None and col not in df.columns]
            print(f"❌ Missing columns: {missing_cols}")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_column_selection()
