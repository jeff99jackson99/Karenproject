#!/usr/bin/env python3
"""
Karen NCB Data Processor - Iteration 2.0
Expected Output: 2k-2500 rows in specific order with proper column mapping
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen NCB 2.0", page_icon="🚀", layout="wide")

def detect_column_structure_v2(excel_file):
    """Detect column structure from Col Ref sheet for version 2.0 - using working logic from first version."""
    try:
        # First, let's see what sheets are available
        excel_data = pd.ExcelFile(excel_file)
        st.write(f"🔍 **Available sheets:** {excel_data.sheet_names}")
        
        # Look for the Col Ref sheet (try different possible names)
        col_ref_sheet = None
        possible_names = ['Col Ref', 'ColRef', 'Column Reference', 'ColumnRef', 'Reference', 'Col_Ref', 'xref', 'Xref']
        
        for sheet_name in excel_data.sheet_names:
            if any(name.lower() in sheet_name.lower() for name in possible_names):
                col_ref_sheet = sheet_name
                break
        
        if not col_ref_sheet:
            st.error(f"❌ Could not find Col Ref sheet. Available sheets: {excel_data.sheet_names}")
            return {}, {}, {}
        
        st.write(f"✅ **Using reference sheet:** {col_ref_sheet}")
        
        # Read the reference sheet
        col_ref_df = pd.read_excel(excel_file, sheet_name=col_ref_sheet)
        
        st.write(f"🔍 **Reference sheet loaded:** {col_ref_df.shape}")
        st.write(f"🔍 **Reference sheet first few rows:**")
        st.dataframe(col_ref_df.head(10))
        
        # Use the EXACT working logic from the first version
        column_mapping = {}
        admin_columns = {}
        
        # Look for the row that contains column descriptions
        for idx, row in col_ref_df.iterrows():
            if 'Admin' in str(row.iloc[2]) or 'NCB' in str(row.iloc[7]):
                # This is the row with column descriptions
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and str(value).strip():
                        desc = str(value).strip()
                        column_mapping[col_idx] = desc
                        
                        # Use the EXACT working logic from first version
                        if 'Admin' in desc or 'NCB' in desc:
                            if 'Amount' in desc or 'Fee' in desc:
                                # This is an Admin amount column
                                if 'Agent' in desc:
                                    admin_columns['Admin 3'] = col_idx
                                elif 'Dealer' in desc:
                                    admin_columns['Admin 4'] = col_idx
                                elif 'Offset' in desc:
                                    if 'Agent' in desc:
                                        admin_columns['Admin 9'] = col_idx
                                    elif 'Dealer' in desc:
                                        admin_columns['Admin 10'] = col_idx
                break
        
        # If we don't have enough admin columns, try to find them by content analysis (working logic from first version)
        if len(admin_columns) < 4:
            st.warning("⚠️ Could not identify all Admin columns from mapping, trying content-based detection...")
            
            # Look for columns that might contain Admin amounts by analyzing their content
            potential_admin_cols = []
            
            # Read the main data sheet to analyze column content
            data_sheet = None
            for sheet in excel_data.sheet_names:
                if sheet != col_ref_sheet and not any(keyword in sheet.lower() for keyword in ['summary', 'xref', 'ref', 'col']):
                    data_sheet = sheet
                    break
            
            if data_sheet:
                df = pd.read_excel(excel_file, sheet_name=data_sheet)
                
                for col in df.columns:
                    try:
                        # Skip datetime columns
                        if isinstance(col, pd.Timestamp) or 'datetime' in str(type(col)).lower():
                            continue
                            
                        # Try to convert to numeric
                        numeric_values = pd.to_numeric(df[col], errors='coerce')
                        if not numeric_values.isna().all() and numeric_values.dtype in ['int64', 'float64']:
                            # Check if this column has meaningful non-zero values
                            non_zero_count = (numeric_values > 0).sum()
                            total_count = len(numeric_values)
                            
                            # Only consider columns with a reasonable number of non-zero values
                            if non_zero_count > 0 and non_zero_count < total_count * 0.9:  # Not all zeros, not all non-zero
                                potential_admin_cols.append({
                                    'column': col,
                                    'non_zero_count': non_zero_count,
                                    'total_count': total_count,
                                    'non_zero_ratio': non_zero_count / total_count
                                })
                    except:
                        pass
                
                # Sort by non-zero ratio to find the most likely Admin columns
                potential_admin_cols.sort(key=lambda x: x['non_zero_ratio'], reverse=True)
                
                st.write(f"🔍 **Potential Admin columns found:** {len(potential_admin_cols)}")
                for i, col_info in enumerate(potential_admin_cols[:10]):  # Show top 10
                    st.write(f"  {i+1}. {col_info['column']}: {col_info['non_zero_count']}/{col_info['total_count']} non-zero ({col_info['non_zero_ratio']:.1%})")
                
                # Select the top 4 columns as Admin columns
                if len(potential_admin_cols) >= 4:
                    admin_columns = {
                        'Admin 3': potential_admin_cols[0]['column'],
                        'Admin 4': potential_admin_cols[1]['column'], 
                        'Admin 9': potential_admin_cols[2]['column'],
                        'Admin 10': potential_admin_cols[3]['column']
                    }
                    
                    st.write(f"✅ **Admin columns assigned by content analysis:**")
                    for admin_type, col_name in admin_columns.items():
                        st.write(f"  {admin_type}: {col_name}")
                else:
                    st.error(f"❌ Not enough potential Admin columns found. Need 4, found {len(potential_admin_cols)}")
        
        if admin_columns:
            st.write("🔍 **Column Structure Detected:**")
            st.write(f"  Total columns mapped: {len(column_mapping)}")
            st.write(f"  Admin columns found: {len(admin_columns)}")
            
            st.write("🔍 **Admin column mappings:**")
            for admin_type, col_idx in admin_columns.items():
                st.write(f"  {admin_type}: Column {col_idx}")
                if col_idx in column_mapping:
                    st.write(f"    Description: {column_mapping[col_idx]}")
            
            return column_mapping, admin_columns, {}
        else:
            st.warning("⚠️ No Admin column mapping found in reference sheet")
            return {}, {}, {}
        
    except Exception as e:
        st.error(f"❌ Error reading reference sheet: {e}")
        import traceback
        st.code(traceback.format_exc())
        return {}, {}, {}

def find_transaction_column(df):
    """Find the transaction type column."""
    transaction_col = None
    
    # Look for transaction type values in columns
    for col in df.columns:
        if df[col].dtype == 'object':
            sample_vals = df[col].dropna().head(10).astype(str).str.upper()
            if any(val in ['NB', 'C', 'R', 'NEW BUSINESS', 'CANCELLATION', 'REINSTATEMENT'] for val in sample_vals):
                transaction_col = col
                break
    
    # Fallback to name-based search
    if not transaction_col:
        for col in df.columns:
            if any(keyword in str(col).upper() for keyword in ['TRANSACTION', 'TYPE', 'TRANS']):
                transaction_col = col
                break
    
    return transaction_col

def process_data_v2(df, column_mapping, admin_columns, amount_columns):
    """Process data according to version 2.0 requirements - using working logic from first version with new column mapping."""
    
    # Find transaction column
    transaction_col = find_transaction_column(df)
    if not transaction_col:
        st.error("❌ Could not find transaction type column")
        return None
    
    st.write(f"✅ **Transaction column found:** {transaction_col}")
    
    # Separate data by transaction type
    nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
    c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
    r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
    
    st.write(f"📊 **Data breakdown:**")
    st.write(f"  New Business: {nb_mask.sum()}")
    st.write(f"  Cancellations: {c_mask.sum()}")
    st.write(f"  Reinstatements: {r_mask.sum()}")
    
    # Process New Business data with the WORKING filtering logic from first version
    nb_df = df[nb_mask].copy()
    
    # Apply the EXACT filtering logic that was working before (giving ~1200 records)
    if len(admin_columns) >= 4:
        # Get the admin column names using the working logic from first version
        admin_cols = [
            admin_columns.get('Admin 3'),
            admin_columns.get('Admin 4'), 
            admin_columns.get('Admin 9'),
            admin_columns.get('Admin 10')
        ]
        
        # Remove None values
        admin_cols = [col for col in admin_cols if col is not None]
        
        if len(admin_cols) < 4:
            st.error(f"❌ Need exactly 4 Admin columns, found {len(admin_cols)}")
            return None
        
        st.write(f"✅ **Processing with Admin columns:** {admin_cols}")
        
        # Convert admin columns to numeric (working logic from first version)
        for col in admin_cols:
            if col in df.columns:
                nb_df[col] = pd.to_numeric(nb_df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ Column {col} not found in data")
                return None
        
        # Calculate sum of admin amounts (working logic from first version)
        nb_df['Admin_Sum'] = nb_df[admin_cols].sum(axis=1)
        
        # Apply the EXACT user requirement: ALL 4 Admin amounts > 0 AND sum > 0
        # This is the logic that was working and giving ~1200 records
        nb_filtered = nb_df[
            (nb_df['Admin_Sum'] > 0) &
            (nb_df[admin_cols[0]] > 0) &
            (nb_df[admin_cols[1]] > 0) &
            (nb_df[admin_cols[2]] > 0) &
            (nb_df[admin_cols[3]] > 0)
        ]
        
        st.write(f"✅ **New Business filtered:** {len(nb_filtered)} records (using working logic from first version)")
        st.write(f"  Expected: ~1200 records")
        st.write(f"  Actual: {len(nb_filtered)} records")
        
    else:
        nb_filtered = nb_df
        st.write(f"⚠️ **Using unfiltered New Business data:** {len(nb_filtered)} records")
        st.write(f"  Reason: Only found {len(admin_columns)} Admin columns, need 4")
    
    # Process Cancellation data (negative/empty/0 values expected)
    c_df = df[c_mask].copy()
    
    # Apply filtering for cancellations (sum != 0) - this was working
    if len(admin_columns) >= 4:
        admin_cols = [
            admin_columns.get('Admin 3'),
            admin_columns.get('Admin 4'), 
            admin_columns.get('Admin 9'),
            admin_columns.get('Admin 10')
        ]
        admin_cols = [col for col in admin_cols if col is not None]
        
        if len(admin_cols) >= 4:
            for col in admin_cols:
                c_df[col] = pd.to_numeric(c_df[col], errors='coerce').fillna(0)
            c_df['Admin_Sum'] = c_df[admin_cols].sum(axis=1)
            c_filtered = c_df[c_df['Admin_Sum'] != 0]
            st.write(f"✅ **Cancellations filtered:** {len(c_filtered)} records")
        else:
            c_filtered = c_df
            st.write(f"⚠️ **Using unfiltered Cancellation data:** {len(c_filtered)} records")
    else:
        c_filtered = c_df
        st.write(f"⚠️ **Using unfiltered Cancellation data:** {len(c_filtered)} records")
    
    # Process Reinstatement data (positive/empty/0 values expected)
    r_df = df[r_mask].copy()
    
    # Apply filtering for reinstatements (sum != 0) - this was working
    if len(admin_columns) >= 4:
        admin_cols = [
            admin_columns.get('Admin 3'),
            admin_columns.get('Admin 4'), 
            admin_columns.get('Admin 9'),
            admin_columns.get('Admin 10')
        ]
        admin_cols = [col for col in admin_cols if col is not None]
        
        if len(admin_cols) >= 4:
            for col in admin_cols:
                r_df[col] = pd.to_numeric(r_df[col], errors='coerce').fillna(0)
            r_df['Admin_Sum'] = r_df[admin_cols].sum(axis=1)
            r_filtered = r_df[r_df['Admin_Sum'] != 0]
            st.write(f"✅ **Reinstatements filtered:** {len(r_filtered)} records")
        else:
            r_filtered = r_df
            st.write(f"⚠️ **Using unfiltered Reinstatement data:** {len(r_filtered)} records")
    else:
        r_filtered = r_df
        st.write(f"⚠️ **Using unfiltered Reinstatement data:** {len(r_filtered)} records")
    
    # Create output dataframes with the NEW specific column mapping from instructions
    # Data Set 1 - New Business (NB)
    nb_output = pd.DataFrame()
    if len(nb_filtered) > 0:
        nb_output['Insurer_Code'] = nb_filtered.iloc[:, 1]  # Column B
        nb_output['Product_Type_Code'] = nb_filtered.iloc[:, 2]  # Column C
        nb_output['Coverage_Code'] = nb_filtered.iloc[:, 3]  # Column D
        nb_output['Dealer_Number'] = nb_filtered.iloc[:, 4]  # Column E
        nb_output['Dealer_Name'] = nb_filtered.iloc[:, 5]  # Column F
        nb_output['Contract_Number'] = nb_filtered.iloc[:, 7]  # Column H
        nb_output['Contract_Sale_Date'] = nb_filtered.iloc[:, 11]  # Column L
        nb_output['Transaction_Date'] = nb_filtered.iloc[:, 9]  # Column J
        nb_output['Transaction_Type'] = nb_filtered.iloc[:, 12]  # Column M
        nb_output['Customer_Last_Name'] = nb_filtered.iloc[:, 20]  # Column U
        
        # Admin Amount columns
        nb_output['Admin_3_Amount_Agent_NCB_Fee'] = nb_filtered.iloc[:, 40]  # Column AO
        nb_output['Admin_4_Amount_Dealer_NCB_Fee'] = nb_filtered.iloc[:, 41]  # Column AQ
        nb_output['Admin_6_Amount_Agent_NCB_Offset'] = nb_filtered.iloc[:, 46]  # Column AU
        nb_output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = nb_filtered.iloc[:, 47]  # Column AW
        nb_output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = nb_filtered.iloc[:, 48]  # Column AY
        nb_output['Admin_9_Amount_Agent_NCB_Offset'] = nb_filtered.iloc[:, 52]  # Column BA
        nb_output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = nb_filtered.iloc[:, 53]  # Column BC
        
        nb_output['Transaction_Type'] = 'NB'
        nb_output['Row_Type'] = 'New Business'
    
    # Data Set 2 - Reinstatements (R)
    r_output = pd.DataFrame()
    if len(r_filtered) > 0:
        r_output['Insurer_Code'] = r_filtered.iloc[:, 1]  # Column B
        r_output['Product_Type_Code'] = r_filtered.iloc[:, 2]  # Column C
        r_output['Coverage_Code'] = r_filtered.iloc[:, 3]  # Column D
        r_output['Dealer_Number'] = r_filtered.iloc[:, 4]  # Column E
        r_output['Dealer_Name'] = r_filtered.iloc[:, 5]  # Column F
        r_output['Contract_Number'] = r_filtered.iloc[:, 7]  # Column H
        r_output['Contract_Sale_Date'] = r_filtered.iloc[:, 11]  # Column L
        r_output['Transaction_Date'] = r_filtered.iloc[:, 9]  # Column J
        r_output['Transaction_Type'] = r_filtered.iloc[:, 12]  # Column M
        r_output['Customer_Last_Name'] = r_filtered.iloc[:, 20]  # Column U
        
        # Admin Amount columns
        r_output['Admin_3_Amount_Agent_NCB_Fee'] = r_filtered.iloc[:, 40]  # Column AO
        r_output['Admin_4_Amount_Dealer_NCB_Fee'] = r_filtered.iloc[:, 41]  # Column AQ
        r_output['Admin_6_Amount_Agent_NCB_Offset'] = r_filtered.iloc[:, 46]  # Column AU
        r_output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = r_filtered.iloc[:, 47]  # Column AW
        r_output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = r_filtered.iloc[:, 48]  # Column AY
        r_output['Admin_9_Amount_Agent_NCB_Offset'] = r_filtered.iloc[:, 52]  # Column BA
        r_output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = r_filtered.iloc[:, 53]  # Column BC
        
        r_output['Transaction_Type'] = 'R'
        r_output['Row_Type'] = 'Reinstatement'
    
    # Data Set 3 - Cancellations (C)
    c_output = pd.DataFrame()
    if len(c_filtered) > 0:
        c_output['Insurer_Code'] = c_filtered.iloc[:, 1]  # Column B
        c_output['Product_Type_Code'] = c_filtered.iloc[:, 2]  # Column C
        c_output['Coverage_Code'] = c_filtered.iloc[:, 3]  # Column D
        c_output['Dealer_Number'] = c_filtered.iloc[:, 4]  # Column E
        c_output['Dealer_Name'] = c_filtered.iloc[:, 5]  # Column F
        c_output['Contract_Number'] = c_filtered.iloc[:, 7]  # Column H
        c_output['Contract_Sale_Date'] = c_filtered.iloc[:, 11]  # Column L
        c_output['Transaction_Date'] = c_filtered.iloc[:, 9]  # Column J
        c_output['Transaction_Type'] = c_filtered.iloc[:, 12]  # Column M
        c_output['Customer_Last_Name'] = c_filtered.iloc[:, 20]  # Column U
        
        # Additional columns for cancellations
        c_output['Contract_Term'] = c_filtered.iloc[:, 25]  # Column Z
        c_output['Cancellation_Date'] = c_filtered.iloc[:, 30]  # Column AE
        c_output['Cancellation_Reason'] = c_filtered.iloc[:, 27]  # Column AB
        c_output['Cancellation_Factor'] = c_filtered.iloc[:, 26]  # Column AA
        
        # Admin Amount columns
        c_output['Admin_3_Amount_Agent_NCB_Fee'] = c_filtered.iloc[:, 40]  # Column AO
        c_output['Admin_4_Amount_Dealer_NCB_Fee'] = c_filtered.iloc[:, 41]  # Column AQ
        c_output['Admin_6_Amount_Agent_NCB_Offset'] = c_filtered.iloc[:, 46]  # Column AU
        c_output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = c_filtered.iloc[:, 47]  # Column AW
        c_output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = c_filtered.iloc[:, 48]  # Column AY
        c_output['Admin_9_Amount_Agent_NCB_Offset'] = c_filtered.iloc[:, 52]  # Column BA
        c_output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = c_filtered.iloc[:, 53]  # Column BC
        
        c_output['Transaction_Type'] = 'C'
        c_output['Row_Type'] = 'Cancellation'
    
    # Create combined output for the main download
    combined_output = pd.concat([nb_output, c_output, r_output], ignore_index=True)
    
    st.write(f"✅ **Output created:** {len(combined_output)} total records")
    st.write(f"  New Business: {len(nb_output)} (target: ~1200)")
    st.write(f"  Cancellations: {len(c_output)}")
    st.write(f"  Reinstatements: {len(r_output)}")
    
    return {
        'nb_data': nb_output,
        'cancellation_data': c_output,
        'reinstatement_data': r_output,
        'combined_data': combined_output,
        'total_records': len(combined_output)
    }

def main():
    st.title("🚀 Karen NCB Data Processor - Iteration 2.0")
    st.markdown("---")
    
    st.write("""
    **Expected Output:** 2k-2500 rows in specific order with proper column mapping
    
    **Data Structure:**
    - **New Business (NB)**: Positive/empty/0 values expected
    - **Reinstatement (R)**: Positive/empty/0 values expected  
    - **Cancellation (C)**: Negative/empty/0 values expected
    
    **Column Mapping:** Labels matched to money values dynamically
    
    **Reference File:** Uses 'Col Ref' sheet or similar reference sheet for column mapping
    """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "📁 Upload Excel file with NCB data and reference sheet (Col Ref, xref, etc.)",
        type=['xlsx', 'xls'],
        help="File should contain data sheet and a reference sheet (like 'Col Ref' or 'xref') for column mapping. The reference sheet should define Admin columns and their relationships."
    )
    
    if uploaded_file is not None:
        st.write(f"📁 **File uploaded:** {uploaded_file.name}")
        
        if st.button("🚀 Process Data (Version 2.0)", type="primary", use_container_width=True):
            
            with st.spinner("🔍 Analyzing file structure..."):
                # Detect column structure
                column_mapping, admin_columns, amount_columns = detect_column_structure_v2(uploaded_file)
                
                if not column_mapping:
                    st.error("❌ Could not detect column structure. Please ensure file has a reference sheet (Col Ref, xref, etc.).")
                    st.write("💡 **Tip:** Your file should have a sheet that defines the Admin column structure (like 'Col Ref' or 'xref').")
                    return
            
            with st.spinner("📊 Processing data..."):
                # Read the main data sheet
                try:
                    excel_data = pd.ExcelFile(uploaded_file)
                    
                    # Find the main data sheet (not the reference sheet)
                    data_sheet = None
                    reference_sheet_names = ['Col Ref', 'ColRef', 'Column Reference', 'ColumnRef', 'Reference', 'Col_Ref', 'xref', 'Xref']
                    
                    for sheet in excel_data.sheet_names:
                        if not any(ref_name.lower() in sheet.lower() for ref_name in reference_sheet_names) and not any(keyword in sheet.lower() for keyword in ['summary', 'ref']):
                            data_sheet = sheet
                            break
                    
                    if not data_sheet:
                        st.error("❌ Could not find main data sheet")
                        return
                    
                    st.write(f"📋 **Processing sheet:** {data_sheet}")
                    
                    # Read data
                    df = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                    st.write(f"📏 **Data shape:** {df.shape}")
                    
                    # Process data
                    output_data = process_data_v2(df, column_mapping, admin_columns, amount_columns)
                    
                    if output_data:
                        st.success(f"✅ **Processing Complete!** Generated {output_data['total_records']} records")
                        
                        # Display results
                        st.subheader("📊 Processing Results")
                        st.write(f"**Total Records:** {output_data['total_records']}")
                        
                        # Show breakdown by transaction type
                        type_counts = output_data['combined_data']['Transaction_Type'].value_counts()
                        st.write("**Breakdown by Transaction Type:**")
                        for trans_type, count in type_counts.items():
                            st.write(f"  {trans_type}: {count} records")
                        
                        # Show sample data in collapsible section
                        with st.expander("🔍 **Sample Output Data** (Click to expand)", expanded=False):
                            st.dataframe(output_data['combined_data'].head(10))
                        
                        # Download buttons for each section
                        st.subheader("💾 Download Results")
                        
                        # Create Excel files in memory for each section
                        nb_buffer = io.BytesIO()
                        c_buffer = io.BytesIO()
                        r_buffer = io.BytesIO()
                        
                        with pd.ExcelWriter(nb_buffer, engine='xlsxwriter') as writer:
                            output_data['nb_data'].to_excel(writer, sheet_name='New_Business', index=False)
                        with pd.ExcelWriter(c_buffer, engine='xlsxwriter') as writer:
                            output_data['cancellation_data'].to_excel(writer, sheet_name='Cancellation', index=False)
                        with pd.ExcelWriter(r_buffer, engine='xlsxwriter') as writer:
                            output_data['reinstatement_data'].to_excel(writer, sheet_name='Reinstatement', index=False)
                        
                        # Generate filenames with timestamp
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        nb_filename = f"Karen_NCB_2_0_New_Business_{timestamp}.xlsx"
                        c_filename = f"Karen_NCB_2_0_Cancellation_{timestamp}.xlsx"
                        r_filename = f"Karen_NCB_2_0_Reinstatement_{timestamp}.xlsx"
                        combined_filename = f"Karen_NCB_2_0_Combined_{timestamp}.xlsx"
                        
                        # Download buttons for each section
                        st.download_button(
                            label="📥 Download New Business Excel File",
                            data=nb_buffer.getvalue(),
                            file_name=nb_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="📥 Download Cancellation Excel File",
                            data=c_buffer.getvalue(),
                            file_name=c_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="📥 Download Reinstatement Excel File",
                            data=r_buffer.getvalue(),
                            file_name=r_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # Create combined download
                        combined_buffer = io.BytesIO()
                        with pd.ExcelWriter(combined_buffer, engine='xlsxwriter') as writer:
                            output_data['combined_data'].to_excel(writer, sheet_name='Combined_Data', index=False)
                        
                        combined_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 Download Combined Excel File",
                            data=combined_buffer.getvalue(),
                            file_name=combined_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.write(f"📁 **Files saved as:**")
                        st.write(f"  New Business: {nb_filename}")
                        st.write(f"  Cancellation: {c_filename}")
                        st.write(f"  Reinstatement: {r_filename}")
                        st.write(f"  Combined: {combined_filename}")
                        
                    else:
                        st.error("❌ No data was processed successfully")
                        
                except Exception as e:
                    st.error(f"❌ Error processing file: {e}")
                    import traceback
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
