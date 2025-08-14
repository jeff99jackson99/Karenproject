#!/usr/bin/env python3
import pandas as pd

# Read the Data sheet
df = pd.read_excel('../2025-0731 Production Summary FINAL.xlsx', sheet_name='Data', header=0)

print("Data shape after header:", df.shape)

print("\nLooking for additional Admin columns that could be NCB columns...")
# Look for columns that might contain Admin amounts
for i, col in enumerate(df.columns):
    try:
        # Skip the header row and look at actual data
        values = df.iloc[1:, i].dropna()
        if len(values) > 0:
            # Check if this column contains numeric values
            numeric_vals = pd.to_numeric(values, errors='coerce')
            if numeric_vals.notna().sum() > 0:
                non_zero = (numeric_vals != 0).sum()
                if non_zero > 0 and non_zero < len(values) * 0.9:  # Not all zeros, not all non-zero
                    # Look for Admin columns specifically
                    if 'ADMIN' in str(col) or 'RESERVE' in str(col):
                        print(f"Column {i} ({col}): {non_zero} non-zero values out of {len(values)}")
                        print(f"  Sample values: {values.head(5).tolist()}")
                        if non_zero > 10:  # Only show columns with significant data
                            print(f"  This looks like a potential NCB column!")
                            print()
    except:
        pass

print("\nLooking for specific Admin columns that might be missing...")
# Check specific Admin columns we haven't mapped yet
admin_positions = [35, 36, 37, 38, 39, 41, 43, 44, 45, 47, 48, 49, 50, 51, 53, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 67, 69, 70, 71, 73, 74]
for pos in admin_positions:
    if pos < len(df.columns):
        col = df.columns[pos]
        # Look at more rows to find actual values
        values = df.iloc[1:100, pos].dropna()
        if len(values) > 0:
            numeric_vals = pd.to_numeric(values, errors='coerce')
            non_zero = (numeric_vals.notna() & (numeric_vals != 0)).sum()
            if non_zero > 0:
                print(f"Column {pos} ({col}): {non_zero} non-zero values out of {len(values)}")
                print(f"  Sample: {values.head(3).tolist()}")
                if 'ADMIN' in str(col) and 'Amount' in str(col):
                    print(f"  *** This is an ADMIN Amount column! ***")
                print()
