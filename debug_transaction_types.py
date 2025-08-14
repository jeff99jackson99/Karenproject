#!/usr/bin/env python3
"""
Debug script to check transaction types in the Excel file
"""

import pandas as pd
import sys

def debug_transaction_types():
    """Debug transaction types in the Excel file."""
    print("🔍 DEBUGGING TRANSACTION TYPES")
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
        print(f"🔍 Header row (Row 12): {header_row[:20].tolist()}")
        
        # Look for transaction type column
        transaction_col_idx = None
        for i, col in enumerate(header_row):
            if 'Transaction Type' in str(col):
                transaction_col_idx = i
                break
        
        if transaction_col_idx is not None:
            print(f"✅ Found Transaction Type column at index: {transaction_col_idx}")
            
            # Get transaction types from data rows (starting from row 13)
            transaction_types = df_raw.iloc[13:, transaction_col_idx].dropna().astype(str)
            print(f"📊 Transaction types found: {len(transaction_types)}")
            
            # Show unique values
            unique_types = transaction_types.unique()
            print(f"🔍 Unique transaction types: {unique_types}")
            
            # Count each type
            for trans_type in unique_types:
                count = (transaction_types == trans_type).sum()
                print(f"  {trans_type}: {count}")
                
        else:
            print("❌ Transaction Type column not found")
            
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_transaction_types()
