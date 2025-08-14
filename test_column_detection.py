#!/usr/bin/env python3
"""
Test script to verify the improved transaction column detection
"""

import pandas as pd

def test_column_detection():
    """Test the improved transaction column detection logic."""
    print("🧪 TESTING IMPROVED TRANSACTION COLUMN DETECTION")
    print("=" * 60)
    
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
        
        # Test the improved transaction column detection logic
        transaction_col = None
        
        print(f"\n🔍 **Step 1: Looking for columns with actual transaction type values (NB, C, R)**")
        # 🔍 IMPROVED TRANSACTION COLUMN DETECTION - PRIORITIZE ACTUAL TRANSACTION DATA
        # First, look for columns that contain the actual transaction type values (NB, C, R)
        # This is the most reliable method since we know these values exist
        for i, col in enumerate(df.columns):
            try:
                col_data = df.iloc[:, i]  # Use integer index to avoid column name issues
                if hasattr(col_data, 'dtype') and col_data.dtype == 'object':
                    # Check if this column contains the actual transaction type values
                    sample_vals = col_data.dropna().head(50).tolist()
                    nb_count = sum(1 for v in sample_vals if str(v).upper() == 'NB')
                    c_count = sum(1 for v in sample_vals if str(v).upper() == 'C')
                    r_count = sum(1 for v in sample_vals if str(v).upper() == 'R')
                    
                    # If we find significant counts of actual transaction types, this is our column
                    if nb_count > 100 or c_count > 100 or r_count > 1:  # NB and C should have many records
                        transaction_col = col
                        print(f"✅ **Found transaction column by actual values:** {col}")
                        print(f"  NB count: {nb_count}, C count: {c_count}, R count: {r_count}")
                        print(f"  Sample values: {sample_vals[:10]}")
                        break
            except Exception as e:
                print(f"⚠️ Error processing column {col}: {e}")
                continue
        
        if not transaction_col:
            print(f"🔍 **Step 2: Looking for columns with names indicating transaction types**")
            # If not found by actual values, look for columns with names indicating transaction types
            for i, col in enumerate(df.columns):
                try:
                    col_name = str(col).lower()
                    if 'transaction' in col_name and 'type' in col_name:
                        col_data = df.iloc[:, i]  # Use integer index to avoid column name issues
                        if hasattr(col_data, 'dtype') and col_data.dtype == 'object':
                            sample_vals = col_data.dropna().head(20).tolist()
                            if any(str(v).upper() in ['NB', 'C', 'R'] for v in sample_vals):
                                transaction_col = col
                                print(f"✅ **Found transaction column by name:** {col} with values: {sample_vals[:10]}")
                                break
                except Exception as e:
                    print(f"⚠️ Error processing column {col}: {e}")
                    continue
        
        if not transaction_col:
            print(f"🔍 **Step 3: Looking for columns with NB, C, R value counts**")
            # If still not found, try alternative search
            for col in df.columns:
                col_data = df[col]
                if hasattr(col_data, 'dtype') and col_data.dtype == 'object':
                    sample_vals = col_data.dropna().head(20).tolist()
                    # Look for any column with NB, C, R values
                    nb_count = sum(1 for v in sample_vals if str(v).upper() == 'NB')
                    c_count = sum(1 for v in sample_vals if str(v).upper() == 'C')
                    r_count = sum(1 for v in sample_vals if str(v).upper() == 'R')
                    
                    if nb_count > 0 or c_count > 0 or r_count > 0:
                        transaction_col = col
                        print(f"✅ **Found transaction column by value count:** {col}")
                        print(f"  NB count: {nb_count}, C count: {c_count}, R count: {r_count}")
                        print(f"  Sample values: {sample_vals[:10]}")
                        break
        
        if not transaction_col:
            print("❌ **No transaction column found!**")
            print("🔍 **Debugging info:**")
            print(f"  Looking for columns containing 'NB', 'C', 'R' values")
            print(f"  Checked {len(df.columns)} columns")
            
            # Show all columns for debugging
            print(f"\n📊 **All columns in the data:**")
            for i, col in enumerate(df.columns):
                print(f"  {i}: {col}")
            return
        
        print(f"\n🎯 **TRANSACTION COLUMN DETECTION SUCCESS!**")
        print(f"  Found column: {transaction_col}")
        
        # Show the transaction type distribution
        if transaction_col in df.columns:
            transaction_data = df[transaction_col]
            if hasattr(transaction_data, 'iloc') and hasattr(transaction_data, 'shape') and len(transaction_data.shape) > 1:
                # It's a DataFrame with multiple columns
                transaction_series = transaction_data.iloc[:, 0]
            else:
                # It's already a Series
                transaction_series = transaction_data
            
            print(f"\n📊 **Transaction type distribution:**")
            unique_types = transaction_series.dropna().unique()
            print(f"  Unique transaction types: {unique_types}")
            
            # Count each type
            for t_type in unique_types:
                if pd.notna(t_type):
                    count = (transaction_series == t_type).sum()
                    print(f"    {t_type}: {count}")
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_column_detection()
