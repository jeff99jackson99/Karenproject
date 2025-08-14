#!/usr/bin/env python3
"""
Karen 2.0 NCB Data Processor
NCB Transaction Summarization with specific column requirements
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen 2.0 NCB Data Processor", page_icon="📊", layout="wide")

def process_excel_data_karen_2_0(uploaded_file):
    """Process Excel data from the Data tab for Karen 2.0."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Create expandable section for processing details
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen 2.0 Processing Started**")
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            # Process ONLY the Data sheet (main transaction data)
            data_sheet_name = 'Data'
            
            if data_sheet_name in excel_data.sheet_names:
                st.write(f"📋 **Processing {data_sheet_name} Sheet**")
                
                # Read the Data sheet
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Show first few column names to understand structure
                st.write(f"🔍 **First 20 columns:** {list(df.columns[:20])}")
                
                # Find NCB columns by looking for Admin columns (AO, AQ, AU, AW, AY, BA, BC)
                ncb_columns = find_ncb_columns_simple(df)
                st.write("🔍 **NCB Columns Found:**")
                for ncb_type, col_name in ncb_columns.items():
                    st.write(f"  {ncb_type}: {col_name}")
                
                # Find required columns by position mapping
                required_cols = find_required_columns_simple(df)
                st.write("🔍 **Required Columns Found:**")
                for col_letter, col_name in required_cols.items():
                    st.write(f"  {col_letter}: {col_name}")
                
                if len(ncb_columns) >= 7 and len(required_cols) >= len(['B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U']):
                    # Process the data
                    results = process_transaction_data_karen_2_0(df, ncb_columns, required_cols)
                    
                    if results:
                        results['ncb_columns'] = ncb_columns
                        results['required_cols'] = required_cols
                        
                        return results
                    else:
                        return None
                else:
                    st.error(f"❌ Need at least 7 NCB columns and 10 required columns, found {len(ncb_columns)} NCB and {len(required_cols)} required")
                    return None
            else:
                st.error(f"❌ '{data_sheet_name}' sheet not found. Available sheets: {excel_data.sheet_names}")
                return None
                
    except Exception as e:
        st.error(f"❌ Error in Karen 2.0 processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def find_ncb_columns_simple(df):
    """Find NCB columns by looking for Admin columns based on actual file structure."""
    ncb_columns = {}
    
    # Based on the actual file structure, map Admin columns to NCB types
    admin_mapping = {
        40: 'Agent NCB Fee',      # ADMIN 3 Amount (Agent NCB) - Column 40
        42: 'Dealer NCB Fee',     # ADMIN 4 Amount (Dealer NCB Fee) - Column 42
        46: 'Agent NCB Offset Bucket', # ADMIN 6 Amount (Agent NCB Offset Bucket) - Column 46
        52: 'Agent NCB Offset',   # ADMIN 9 Amount (Agent NCB Offset) - Column 52
        54: 'Dealer NCB Offset Bucket', # ADMIN 10 Amount (Dealer NCB Offset Bucket) - Column 54
        36: 'Admin 1 Amount',     # ADMIN 1 Amount - Column 36 (additional)
        38: 'Admin 2 Amount',     # ADMIN 2 Amount - Column 38 (additional)
    }
    
    # Map Admin columns by their actual positions
    for pos, ncb_type in admin_mapping.items():
        if pos < len(df.columns):
            col = df.columns[pos]
            ncb_columns[ncb_type] = col
            st.write(f"✅ **Found NCB column:** {ncb_type} → {col} (Position {pos})")
    
    # If we still don't have enough, try to find additional Admin columns
    if len(ncb_columns) < 7:
        st.warning(f"⚠️ **Only found {len(ncb_columns)} NCB columns, need 7. Looking for additional Admin columns...**")
        
        # Look for additional Admin columns that might contain NCB amounts
        for i, col in enumerate(df.columns):
            try:
                if 'ADMIN' in str(col) and 'Amount' in str(col):
                    # Check if this column has meaningful values
                    values = df.iloc[1:, i].dropna()  # Skip header row
                    if len(values) > 0:
                        numeric_vals = pd.to_numeric(values, errors='coerce')
                        non_zero = (numeric_vals != 0).sum()
                        if non_zero > 0 and non_zero < len(values) * 0.9:  # Not all zeros, not all non-zero
                            # Find a unique name for this Admin column
                            admin_num = f"Admin_{len(ncb_columns) + 1}"
                            ncb_columns[admin_num] = col
                            st.write(f"✅ **Found additional Admin column:** {admin_num} → {col} (Position {i})")
                            
                            if len(ncb_columns) >= 7:
                                break
            except:
                pass
    
    return ncb_columns

def find_required_columns_simple(df):
    """Find required columns by simple position mapping based on actual file structure."""
    required_cols = {}
    
    # Based on the actual file structure, map column positions to required columns
    position_mapping = {
        1: 'Insurer Code',        # Column 1 (B)
        2: 'Product Type Code',   # Column 2 (C) 
        3: 'Coverage Code',       # Column 3 (D)
        4: 'Dealer Number',       # Column 4 (E)
        5: 'Dealer Name',         # Column 5 (F)
        6: 'Contract Number',     # Column 6 (G)
        7: 'Contract Sale Date',  # Column 7 (H)
        8: 'Transaction Date',    # Column 8 (I)
        9: 'Transaction Type',    # Column 9 (J) - This contains the dropdown values C, R, NB, A
        10: 'Customer Last Name', # Column 10 (K)
        20: 'Contract Term',      # Column 20 (U) - Term Months
        27: 'Cancellation Factor', # Column 27 (AB)
        28: 'Cancellation Reason', # Column 28 (AC)
        26: 'Cancellation Date',  # Column 26 (AA) - Number of Days in Force
    }
    
    # Map by position
    for i, col in enumerate(df.columns):
        if i in position_mapping:
            required_cols[chr(65 + i)] = col  # Convert position to letter (A=0, B=1, etc.)
            st.write(f"✅ **Found required column:** {chr(65 + i)} ({position_mapping[i]}) → {col}")
    
    return required_cols

def process_transaction_data_karen_2_0(df, ncb_columns, required_cols):
    """Process transaction data according to Karen 2.0 specifications."""
    try:
        # Get the NCB column names
        ncb_cols = list(ncb_columns.values())
        
        if len(ncb_cols) < 7:
            st.error(f"❌ Need exactly 7 NCB columns, found {len(ncb_cols)}")
            return None
        
        st.write(f"✅ **Processing with NCB columns:** {ncb_cols}")
        
        # Convert NCB columns to numeric
        for col in ncb_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ Column {col} not found in data")
                return None
        
        # Calculate sum of NCB amounts
        df['NCB_Sum'] = df[ncb_cols].sum(axis=1)
        
        # Find transaction type column
        transaction_col = required_cols.get('J') # Changed from 'M' to 'J'
        if not transaction_col:
            st.error("❌ No transaction type column found")
            return None
        
        st.write(f"✅ **Transaction column found:** {transaction_col}")
        
        # Show unique values in transaction column
        unique_transactions = df[transaction_col].astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")
        
        # Filter by transaction type and NCB amounts according to Karen 2.0 rules
        nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
        c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
        r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        
        # Apply Karen 2.0 filtering rules
        # Data Set 1 (NB): sum > 0
        nb_df = df[nb_mask].copy()
        nb_filtered = nb_df[nb_df['NCB_Sum'] > 0]
        
        # Data Set 2 (R): sum > 0
        r_df = df[r_mask].copy()
        r_filtered = r_df[r_df['NCB_Sum'] > 0]
        
        # Data Set 3 (C): sum < 0
        c_df = df[c_mask].copy()
        c_filtered = c_df[c_df['NCB_Sum'] < 0]
        
        st.write(f"📊 **Filtered results (Karen 2.0 rules):**")
        st.write(f"  New Business (sum > 0): {len(nb_filtered)}")
        st.write(f"  Reinstatements (sum > 0): {len(r_filtered)}")
        st.write(f"  Cancellations (sum < 0): {len(c_filtered)}")
        
        # Create output dataframes with required columns in correct order
        # Data Set 1: New Business (NB)
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('G'),
            required_cols.get('H'), required_cols.get('I'), required_cols.get('J'),
            required_cols.get('K'), ncb_columns.get('Agent NCB Fee'),
            ncb_columns.get('Dealer NCB Fee'), ncb_columns.get('Agent NCB Offset Bucket'),
            ncb_columns.get('Agent NCB Offset'), ncb_columns.get('Dealer NCB Offset Bucket'),
            ncb_columns.get('Admin 1 Amount'), ncb_columns.get('Admin 2 Amount')
        ]
        
        # Data Set 2: Reinstatements (R) - same columns as NB
        r_output_cols = nb_output_cols.copy()
        
        # Data Set 3: Cancellations (C) - additional columns
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('G'),
            required_cols.get('H'), required_cols.get('I'), required_cols.get('J'),
            required_cols.get('K'), required_cols.get('U'), required_cols.get('AA'),
            required_cols.get('AC'), required_cols.get('AB'), ncb_columns.get('Agent NCB Fee'),
            ncb_columns.get('Dealer NCB Fee'), ncb_columns.get('Agent NCB Offset Bucket'),
            ncb_columns.get('Agent NCB Offset'), ncb_columns.get('Dealer NCB Offset Bucket'),
            ncb_columns.get('Admin 1 Amount'), ncb_columns.get('Admin 2 Amount')
        ]
        
        # Filter dataframes to only include available columns
        nb_output_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        r_output_cols = [col for col in r_output_cols if col is not None and col in df.columns]
        c_output_cols = [col for col in c_output_cols if col is not None and col in df.columns]
        
        # Create output dataframes
        nb_output = nb_filtered[nb_output_cols].copy()
        r_output = r_filtered[r_output_cols].copy()
        c_output = c_filtered[c_output_cols].copy()
        
        # Set column headers based on Karen 2.0 specifications
        nb_headers = [
            'Insurer Code', 'Product Type Code', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Customer Last Name',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)', 'Admin 1 Amount', 'Admin 2 Amount'
        ]
        
        r_headers = nb_headers.copy()
        
        c_headers = [
            'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Last Name',
            'Contract Term', 'Cancellation Date', 'Cancellation Reason', 'Cancellation Factor',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)', 'Admin 1 Amount', 'Admin 2 Amount'
        ]
        
        # Apply headers
        if len(nb_headers) == len(nb_output.columns):
            nb_output.columns = nb_headers
        if len(r_headers) == len(r_output.columns):
            r_output.columns = r_headers
        if len(c_headers) == len(c_output.columns):
            c_output.columns = c_headers
        
        st.write(f"🔍 **Output dataframe shapes:**")
        st.write(f"  NB: {nb_output.shape}")
        st.write(f"  R: {r_output.shape}")
        st.write(f"  C: {c_output.shape}")
        
        return {
            'nb_data': nb_output,
            'reinstatement_data': r_output,
            'cancellation_data': c_output,
            'ncb_columns': ncb_columns,
            'total_records': len(nb_output) + len(r_output) + len(c_output)
        }
        
    except Exception as e:
        st.error(f"❌ Error processing transaction data: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_excel_download_karen_2_0(results):
    """Create Excel download with three separate worksheets as per Karen 2.0."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Data Set 1: New Business (NB)
        if len(results['nb_data']) > 0:
            results['nb_data'].to_excel(writer, sheet_name='Data Set 1 - New Business', index=False)
        
        # Data Set 2: Reinstatements (R)
        if len(results['reinstatement_data']) > 0:
            results['reinstatement_data'].to_excel(writer, sheet_name='Data Set 2 - Reinstatements', index=False)
        
        # Data Set 3: Cancellations (C)
        if len(results['cancellation_data']) > 0:
            results['cancellation_data'].to_excel(writer, sheet_name='Data Set 3 - Cancellations', index=False)
    
    return output.getvalue()

def main():
    st.title("📊 Karen 2.0 NCB Data Processor")
    st.markdown("**NCB Transaction Summarization with specific column requirements**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Karen 2.0 Processing")
        st.markdown("**This app will:**")
        st.markdown("• Extract NB, R, and C transaction types")
        st.markdown("• Apply Karen 2.0 filtering rules")
        st.markdown("• Create 3 separate worksheets")
        st.markdown("• Output 2,000-2,500 rows total")
        st.markdown("• Preserve exact column order")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file for processing", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("📊 Start Karen 2.0 Processing", type="primary", use_container_width=True):
            with st.spinner("Karen 2.0 processing in progress..."):
                results = process_excel_data_karen_2_0(uploaded_file)
                
                if results:
                    st.success("✅ Karen 2.0 processing complete!")
                    
                    # Show results
                    st.header("📊 Processing Results - Karen 2.0")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    with col3:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col4:
                        st.metric("Total Records", results['total_records'])
                    
                    # Show NCB columns used
                    st.subheader("🔍 NCB Columns Used")
                    for ncb_type, col_name in results['ncb_columns'].items():
                        st.write(f"  **{ncb_type}:** {col_name}")
                    
                    # Show expected output information
                    st.subheader("📋 Expected Output - Karen 2.0")
                    st.write("**Data Set 1 - New Business (NB):** New contracts with NCB sum > 0")
                    st.write("**Data Set 2 - Reinstatements (R):** Reinstated records with NCB sum > 0")
                    st.write("**Data Set 3 - Cancellations (C):** Cancelled records with NCB sum < 0")
                    st.write(f"**Total expected rows:** 2,000-2,500 (Current: {results['total_records']})")
                    
                    # Show sample data
                    if len(results['nb_data']) > 0:
                        st.subheader("🆕 Data Set 1 - New Business (NB)")
                        st.dataframe(results['nb_data'].head(10))
                    
                    if len(results['reinstatement_data']) > 0:
                        st.subheader("🔄 Data Set 2 - Reinstatements (R)")
                        st.dataframe(results['reinstatement_data'].head(10))
                    
                    if len(results['cancellation_data']) > 0:
                        st.subheader("❌ Data Set 3 - Cancellations (C)")
                        st.dataframe(results['cancellation_data'].head(10))
                    
                    # Download button for all three datasets
                    st.download_button(
                        label="📥 Download All Datasets (Karen 2.0 Format)",
                        data=create_excel_download_karen_2_0(results),
                        file_name=f"NCB_Karen_2_0_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Individual download buttons
                    if len(results['nb_data']) > 0:
                        st.download_button(
                            label="📥 Download New Business Data",
                            data=create_excel_download_karen_2_0({'nb_data': results['nb_data'], 'reinstatement_data': pd.DataFrame(), 'cancellation_data': pd.DataFrame()}),
                            file_name=f"Data_Set_1_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        st.download_button(
                            label="📥 Download Reinstatement Data",
                            data=create_excel_download_karen_2_0({'nb_data': pd.DataFrame(), 'reinstatement_data': results['reinstatement_data'], 'cancellation_data': pd.DataFrame()}),
                            file_name=f"Data_Set_2_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        st.download_button(
                            label="📥 Download Cancellation Data",
                            data=create_excel_download_karen_2_0({'nb_data': pd.DataFrame(), 'reinstatement_data': pd.DataFrame(), 'cancellation_data': results['cancellation_data']}),
                            file_name=f"Data_Set_3_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("❌ Karen 2.0 processing failed. Check the error messages above.")

if __name__ == "__main__":
    main()
