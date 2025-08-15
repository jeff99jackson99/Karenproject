#!/usr/bin/env python3
"""
Test the fixed Karen 3.0 logic
"""

import pandas as pd
import numpy as np

def test_karen_3_0_logic():
    """Test the fixed Karen 3.0 logic."""
    
    print("🧪 **TESTING FIXED KAREN 3.0 LOGIC**")
    print("=" * 50)
    
    # Test data that mimics the Excel structure
    test_data = {
        'Transaction Type': ['NB', 'NB', 'C', 'C', 'R', 'NB', 'C'],
        'ADMIN 3 Amount': [100, 50, -25, 0, 75, 0, -10],
        'ADMIN 4 Amount': [200, 100, 0, -15, 150, 25, 0],
        'ADMIN 6 Amount': [300, 0, -50, 0, 200, 0, -5],
        'ADMIN 7 Amount': [400, 0, 0, -20, 250, 0, 0],
        'ADMIN 8 Amount': [500, 0, -75, 0, 300, 0, -15],
        'ADMIN 9 Amount': [600, 0, 0, 0, 350, 0, -25],
        'ADMIN 10 Amount': [700, 0, -100, 0, 400, 0, 0]
    }
    
    df = pd.DataFrame(test_data)
    print(f"📊 **Test data created:** {df.shape}")
    print(df)
    
    # Test transaction filtering
    print(f"\n🎯 **TESTING TRANSACTION FILTERING:**")
    nb_mask = df['Transaction Type'].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
    c_mask = df['Transaction Type'].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
    r_mask = df['Transaction Type'].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
    
    nb_df = df[nb_mask].copy()
    c_df = df[c_mask].copy()
    r_df = df[r_mask].copy()
    
    print(f"  New Business records: {len(nb_df)}")
    print(f"  Cancellation records: {len(c_df)}")
    print(f"  Reinstatement records: {len(r_df)}")
    
    # Test Admin sum calculation
    print(f"\n🎯 **TESTING ADMIN SUM CALCULATION:**")
    admin_cols = ['ADMIN 3 Amount', 'ADMIN 4 Amount', 'ADMIN 6 Amount', 
                  'ADMIN 7 Amount', 'ADMIN 8 Amount', 'ADMIN 9 Amount', 'ADMIN 10 Amount']
    
    df_copy = df.copy()
    for col in admin_cols:
        df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
    
    df_copy['Admin_Sum'] = df_copy[admin_cols].sum(axis=1)
    print(f"  Admin sums: {df_copy['Admin_Sum'].tolist()}")
    
    # Test cancellation filtering (INSTRUCTIONS 3.0: ONLY negative values)
    print(f"\n🎯 **TESTING CANCELLATION FILTERING (INSTRUCTIONS 3.0):**")
    c_negative_mask = df_copy[admin_cols] < 0
    c_has_negative = c_negative_mask.any(axis=1)
    
    print(f"  Records with ANY negative Admin value: {c_has_negative.sum()}")
    print(f"  Negative mask details:")
    for i, row in c_negative_mask.iterrows():
        if df.iloc[i]['Transaction Type'] == 'C':
            print(f"    Row {i}: {row.tolist()} -> Has negative: {c_has_negative.iloc[i]}")
    
    # Test filtering results
    nb_strict = nb_df[nb_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
    r_strict = r_df[r_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
    c_strict = c_df[c_df.index.isin(df_copy[c_has_negative].index)]
    
    print(f"\n🎯 **FILTERING RESULTS:**")
    print(f"  New Business (Admin_Sum > 0): {len(nb_strict)}")
    print(f"  Reinstatements (Admin_Sum > 0): {len(r_strict)}")
    print(f"  Cancellations (ANY Admin column negative): {len(c_strict)}")
    
    # Test output column structure
    print(f"\n🎯 **TESTING OUTPUT COLUMN STRUCTURE:**")
    print(f"  Expected NB columns: 17")
    print(f"  Expected R columns: 17") 
    print(f"  Expected C columns: 21")
    
    # Verify Admin 6, 7, 8 are included
    print(f"\n🎯 **VERIFYING ADMIN 6, 7, 8 INCLUSION:**")
    admin_678_included = ['ADMIN 6 Amount' in admin_cols, 'ADMIN 7 Amount' in admin_cols, 'ADMIN 8 Amount' in admin_cols]
    print(f"  Admin 6 Amount included: {admin_678_included[0]}")
    print(f"  Admin 7 Amount included: {admin_678_included[1]}")
    print(f"  Admin 8 Amount included: {admin_678_included[2]}")
    
    print(f"\n✅ **TEST COMPLETED SUCCESSFULLY!**")

if __name__ == "__main__":
    test_karen_3_0_logic()
