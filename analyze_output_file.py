#!/usr/bin/env python3
"""
Analyze Karen 3.0 Cancellations Output File
"""

import pandas as pd
import os

def analyze_output_file():
    print("🔍 Analyzing Karen_3_0_Cancellations.xlsx output file...")
    
    # Check if file exists
    file_path = "Karen_3_0_Cancellations.xlsx"
    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return
    
    try:
        # Read the Excel file
        excel_data = pd.ExcelFile(file_path)
        print(f"📊 Sheets found: {excel_data.sheet_names}")
        
        # Read the first sheet (assuming it's the cancellations data)
        sheet_name = excel_data.sheet_names[0]
        print(f"📋 Reading sheet: {sheet_name}")
        
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        print(f"📏 Data shape: {df.shape}")
        print(f"📋 Columns: {list(df.columns)}")
        
        # Look for Admin columns
        admin_cols = []
        for col in df.columns:
            if 'ADMIN' in str(col).upper():
                admin_cols.append(col)
        
        print(f"\n🔍 Found {len(admin_cols)} Admin columns:")
        for col in admin_cols:
            print(f"  - {col}")
        
        # Check data in Admin columns
        print(f"\n🔍 Admin column data analysis:")
        for col in admin_cols:
            col_data = df[col]
            non_null_count = col_data.notna().sum()
            zero_count = (col_data == 0).sum()
            negative_count = (col_data < 0).sum()
            positive_count = (col_data > 0).sum()
            
            print(f"  {col}:")
            print(f"    Total rows: {len(col_data)}")
            print(f"    Non-null: {non_null_count}")
            print(f"    Zeros: {zero_count}")
            print(f"    Negative: {negative_count}")
            print(f"    Positive: {positive_count}")
            
            if non_null_count > 0:
                sample_values = col_data.dropna().head(3).tolist()
                print(f"    Sample values: {sample_values}")
            print()
        
        # Check for MIN columns (which might be the output format)
        min_cols = []
        for col in df.columns:
            if 'MIN' in str(col).upper():
                min_cols.append(col)
        
        if min_cols:
            print(f"🔍 Found {len(min_cols)} MIN columns:")
            for col in min_cols:
                print(f"  - {col}")
            
            print(f"\n🔍 MIN column data analysis:")
            for col in min_cols:
                col_data = df[col]
                non_null_count = col_data.notna().sum()
                zero_count = (col_data == 0).sum()
                negative_count = (col_data < 0).sum()
                positive_count = (col_data > 0).sum()
                
                print(f"  {col}:")
                print(f"    Total rows: {len(col_data)}")
                print(f"    Non-null: {non_null_count}")
                print(f"    Zeros: {zero_count}")
                print(f"    Negative: {negative_count}")
                print(f"    Positive: {positive_count}")
                
                if non_null_count > 0:
                    sample_values = col_data.dropna().head(3).tolist()
                    print(f"    Sample values: {sample_values}")
                print()
        
        # Show first few rows to understand the structure
        print(f"🔍 First 3 rows of data:")
        print(df.head(3).to_string())
        
        # Summary
        print(f"\n📊 Summary:")
        print(f"  - Total rows: {len(df)}")
        print(f"  - Total columns: {len(df.columns)}")
        print(f"  - Admin columns: {len(admin_cols)}")
        print(f"  - MIN columns: {len(min_cols)}")
        
        # Identify the issue
        print(f"\n🔍 Issue Analysis:")
        if min_cols:
            # Check if MIN 6, 7, 8 are all zeros
            min_6_col = None
            min_7_col = None
            min_8_col = None
            
            for col in min_cols:
                if 'MIN 6' in str(col).upper():
                    min_6_col = col
                elif 'MIN 7' in str(col).upper():
                    min_7_col = col
                elif 'MIN 8' in str(col).upper():
                    min_8_col = col
            
            if min_6_col and df[min_6_col].sum() == 0:
                print(f"  ❌ {min_6_col}: All values are 0")
            if min_7_col and df[min_7_col].sum() == 0:
                print(f"  ❌ {min_7_col}: All values are 0")
            if min_8_col and df[min_8_col].sum() == 0:
                print(f"  ❌ {min_8_col}: All values are 0")
        
        print(f"\n💡 This analysis should help identify why Admin 6, 7, 8 columns are showing as 0's")
        
    except Exception as e:
        print(f"❌ Error analyzing file: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    analyze_output_file()
