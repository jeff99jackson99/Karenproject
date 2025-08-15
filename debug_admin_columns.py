#!/usr/bin/env python3
"""
Debug Admin Columns - Check what Admin columns actually exist in the Excel file
"""

import pandas as pd
import streamlit as st

def debug_admin_columns():
    st.title("🔍 Debug Admin Columns")
    st.write("This script will help identify why Admin 6, 7, 8 columns are showing as 0's")
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            excel_data = pd.ExcelFile(uploaded_file)
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            # Check if Data sheet exists
            if 'Data' in excel_data.sheet_names:
                st.write("📋 **Processing Data Sheet**")
                
                # Read the Data sheet - row 12 (index 12) contains the actual headers
                df = pd.read_excel(uploaded_file, sheet_name='Data', header=None)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Row 12 (index 12) contains the actual column headers
                header_row = df.iloc[12]
                st.write(f"🔍 **Header row (Row 12):** {header_row[:20].tolist()}")
                
                # Create DataFrame with proper column names
                data_rows = df.iloc[13:].copy()
                new_df = data_rows.copy()
                new_df.columns = header_row
                
                st.write("🔍 **Looking for Admin columns...**")
                
                # Find all columns containing "ADMIN"
                admin_cols = []
                for col in new_df.columns:
                    if 'ADMIN' in str(col).upper():
                        admin_cols.append(col)
                
                st.write(f"🔍 **Found {len(admin_cols)} Admin columns:**")
                for col in admin_cols:
                    st.write(f"  - {col}")
                
                # Check specific Admin columns that should exist
                required_admin_cols = [
                    'ADMIN 3 Amount', 'ADMIN 4 Amount', 'ADMIN 6 Amount', 
                    'ADMIN 7 Amount', 'ADMIN 8 Amount', 'ADMIN 9 Amount', 'ADMIN 10 Amount'
                ]
                
                st.write("🔍 **Checking required Admin columns:**")
                found_cols = {}
                for req_col in required_admin_cols:
                    if req_col in new_df.columns:
                        found_cols[req_col] = True
                        # Check data in this column
                        col_data = new_df[req_col]
                        non_null_count = col_data.notna().sum()
                        if non_null_count > 0:
                            sample_values = col_data.dropna().head(5).tolist()
                            st.write(f"  ✅ {req_col}: {non_null_count} non-null values, Sample: {sample_values}")
                        else:
                            st.write(f"  ⚠️ {req_col}: Found but all values are null/0")
                    else:
                        found_cols[req_col] = False
                        st.write(f"  ❌ {req_col}: NOT FOUND")
                
                # Check for similar column names
                st.write("🔍 **Looking for similar column names that might be Admin columns:**")
                for col in new_df.columns:
                    col_str = str(col).upper()
                    if any(admin_num in col_str for admin_num in ['ADMIN 6', 'ADMIN 7', 'ADMIN 8']):
                        st.write(f"  🔍 Found similar: {col}")
                        # Check if it's an amount column
                        if 'AMOUNT' in col_str:
                            col_data = new_df[col]
                            non_null_count = col_data.notna().sum()
                            if non_null_count > 0:
                                sample_values = col_data.dropna().head(3).tolist()
                                st.write(f"    Amount column with {non_null_count} non-null values, Sample: {sample_values}")
                            else:
                                st.write(f"    Amount column but all values are null/0")
                
                # Show first few rows of data to see what's actually there
                st.write("🔍 **First 5 rows of data (Admin columns only):**")
                admin_data = new_df[admin_cols].head(5)
                st.dataframe(admin_data)
                
                # Summary
                st.write("📊 **Summary:**")
                found_count = sum(found_cols.values())
                st.write(f"  - Required Admin columns found: {found_count}/{len(required_admin_cols)}")
                
                if found_count < len(required_admin_cols):
                    st.error("❌ **Missing Admin columns detected!**")
                    st.write("This explains why some columns are showing as 0's in your output.")
                else:
                    st.success("✅ **All required Admin columns found!**")
                    st.write("The issue might be in the data processing or column mapping.")
                
            else:
                st.error(f"❌ 'Data' sheet not found. Available sheets: {excel_data.sheet_names}")
                
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    debug_admin_columns()
