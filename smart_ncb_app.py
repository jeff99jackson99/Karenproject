#!/usr/bin/env python3
"""
Smart NCB Data Processor - Works with unnamed columns by detecting structure from Col Ref sheet
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Smart NCB Data Processor", page_icon="🧠", layout="wide")

def detect_column_structure(excel_file):
    """Detect the actual column structure from the Col Ref sheet."""
    try:
        # Read the Col Ref sheet to understand column mapping
        col_ref_df = pd.read_excel(excel_file, sheet_name='Col Ref')
        
        # Look for the row that contains column descriptions
        column_mapping = {}
        
        # Find the row with column descriptions (usually row 1)
        for idx, row in col_ref_df.iterrows():
            if 'Admin' in str(row.iloc[2]) or 'NCB' in str(row.iloc[7]):
                # This is the row with column descriptions
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and str(value).strip():
                        column_mapping[col_idx] = str(value).strip()
                break
        
        st.write("🔍 **Column Structure Detected:**")
        for col_idx, desc in column_mapping.items():
            st.write(f"  Column {col_idx}: {desc}")
        
        return column_mapping
        
    except Exception as e:
        st.error(f"❌ Error detecting column structure: {e}")
        return {}

def find_admin_columns(df, column_mapping):
    """Find Admin columns based on the detected structure."""
    admin_columns = {}
    
    # Look for Admin columns based on the mapping
    for col_idx, desc in column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            
            if 'Admin' in desc or 'NCB' in desc:
                if 'Amount' in desc or 'Fee' in desc:
                    # This is an Admin amount column
                    if 'Agent' in desc:
                        admin_columns['Admin 3'] = col_name
                    elif 'Dealer' in desc:
                        admin_columns['Admin 4'] = col_name
                    elif 'Offset' in desc:
                        if 'Agent' in desc:
                            admin_columns['Admin 9'] = col_name
                        elif 'Dealer' in desc:
                            admin_columns['Admin 10'] = col_name
    
    # If we don't have enough admin columns, try to find them by position
    if len(admin_columns) < 4:
        st.warning("⚠️ Could not identify all Admin columns from mapping, trying position-based detection...")
        
        # Look for numeric columns that might be admin amounts
        numeric_cols = []
        for col in df.columns:
            try:
                # Skip datetime columns
                if isinstance(col, pd.Timestamp) or 'datetime' in str(type(col)).lower():
                    continue
                    
                # Try to convert to numeric
                numeric_values = pd.to_numeric(df[col], errors='coerce')
                if not numeric_values.isna().all() and numeric_values.dtype in ['int64', 'float64']:
                    # Only include columns that are actually numeric
                    numeric_cols.append(col)
            except:
                pass
        
        st.write(f"🔍 **Numeric columns found:** {len(numeric_cols)}")
        st.write(f"  Sample columns: {numeric_cols[:10]}")
        
        # Assign admin columns by position (assuming they're in order)
        if len(numeric_cols) >= 4:
            admin_columns = {
                'Admin 3': numeric_cols[0] if len(numeric_cols) > 0 else None,
                'Admin 4': numeric_cols[1] if len(numeric_cols) > 1 else None,
                'Admin 9': numeric_cols[2] if len(numeric_cols) > 2 else None,
                'Admin 10': numeric_cols[3] if len(numeric_cols) > 3 else None
            }
            
            st.write(f"✅ **Admin columns assigned by position:**")
            for admin_type, col_name in admin_columns.items():
                st.write(f"  {admin_type}: {col_name}")
        else:
            st.error(f"❌ Not enough numeric columns found. Need 4, found {len(numeric_cols)}")
    
    return admin_columns

def process_excel_data_smart(uploaded_file):
    """Process Excel data using smart column detection."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        st.write("🔍 **Smart Processing Started**")
        st.write(f"📊 Sheets found: {excel_data.sheet_names}")
        
        # Detect column structure
        column_mapping = detect_column_structure(uploaded_file)
        
        # Process the Data sheet (main transaction data)
        if 'Data' in excel_data.sheet_names:
            st.write("📋 **Processing Data Sheet**")
            
            # Read the Data sheet
            df = pd.read_excel(uploaded_file, sheet_name='Data')
            st.write(f"📏 Data shape: {df.shape}")
            
            # Find admin columns
            admin_columns = find_admin_columns(df, column_mapping)
            st.write("🔍 **Admin Columns Found:**")
            for admin_type, col_name in admin_columns.items():
                st.write(f"  {admin_type}: {col_name}")
            
            if len(admin_columns) >= 4:
                # Process the data
                return process_transaction_data(df, admin_columns)
            else:
                st.error(f"❌ Need at least 4 Admin columns, found {len(admin_columns)}")
                return None
        else:
            st.error("❌ No 'Data' sheet found")
            return None
            
    except Exception as e:
        st.error(f"❌ Error in smart processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def process_transaction_data(df, admin_columns):
    """Process transaction data with the detected admin columns."""
    try:
        # Get the admin column names
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
        
        # Convert admin columns to numeric
        for col in admin_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ Column {col} not found in data")
                return None
        
        # Calculate sum of admin amounts
        df['Admin_Sum'] = df[admin_cols].sum(axis=1)
        
        # Find transaction type column
        transaction_col = None
        for col in df.columns:
            if 'transaction' in str(col).lower() or 'type' in str(col).lower():
                transaction_col = col
                break
        
        if not transaction_col:
            st.error("❌ No transaction type column found")
            return None
        
        st.write(f"✅ **Transaction column found:** {transaction_col}")
        
        # Show unique values in transaction column to understand the data
        unique_transactions = df[transaction_col].astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")  # Show first 20
        
        # Filter by transaction type and admin amounts
        nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
        c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
        r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        
        # If no transactions found, try to understand the data better
        if nb_mask.sum() == 0 and c_mask.sum() == 0 and r_mask.sum() == 0:
            st.warning("⚠️ No standard transaction types found. Let me examine the data...")
            
            # Show sample values from the transaction column
            sample_values = df[transaction_col].dropna().head(20).tolist()
            st.write(f"🔍 **Sample transaction values:** {sample_values}")
            
            # Try to find any non-null values that might be transaction types
            non_null_values = df[transaction_col].dropna().unique()
            st.write(f"🔍 **All non-null transaction values:** {non_null_values[:20]}")
            
            # Look for patterns in the data
            for col in df.columns[:10]:  # Check first 10 columns
                if df[col].dtype == 'object':
                    sample_vals = df[col].dropna().head(5).tolist()
                    if any('NB' in str(v) or 'C' in str(v) or 'R' in str(v) for v in sample_vals):
                        st.write(f"🔍 **Potential transaction column {col}:** {sample_vals}")
        
        # Filter NB data (sum > 0 and all 4 admin values present)
        nb_df = df[nb_mask].copy()
        nb_filtered = nb_df[
            (nb_df['Admin_Sum'] > 0) &
            (nb_df[admin_cols[0]] != 0) &
            (nb_df[admin_cols[1]] != 0) &
            (nb_df[admin_cols[2]] != 0) &
            (nb_df[admin_cols[3]] != 0)
        ]
        
        # Filter C/R data (sum != 0)
        cr_df = df[c_mask | r_mask].copy()
        cr_filtered = cr_df[cr_df['Admin_Sum'] != 0]
        
        st.write(f"📊 **Filtered results:**")
        st.write(f"  New Business (filtered): {len(nb_filtered)}")
        st.write(f"  Cancellations/Reinstatements (filtered): {len(cr_filtered)}")
        
        return {
            'nb_data': nb_filtered,
            'cancellation_data': cr_filtered[cr_df[transaction_col] == 'C'],
            'reinstatement_data': cr_filtered[cr_df[transaction_col] == 'R'],
            'admin_columns': admin_cols,
            'total_records': len(nb_filtered) + len(cr_filtered)
        }
        
    except Exception as e:
        st.error(f"❌ Error processing transaction data: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_excel_download(df, sheet_name):
    """Create Excel download for a dataframe."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def main():
    st.title("🧠 Smart NCB Data Processor")
    st.markdown("**Automatically detects column structure and processes NCB data**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Smart Processing")
        st.markdown("**This app will:**")
        st.markdown("• Detect column structure automatically")
        st.markdown("• Map unnamed columns to Admin fields")
        st.markdown("• Process data based on actual structure")
        st.markdown("• Show detailed debugging information")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file for processing", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🧠 Start Smart Processing", type="primary", use_container_width=True):
            with st.spinner("Smart processing in progress..."):
                results = process_excel_data_smart(uploaded_file)
                
                if results:
                    st.success("✅ Smart processing complete!")
                    
                    # Show results
                    st.header("📊 Processing Results")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col3:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    with col4:
                        st.metric("Total Records", results['total_records'])
                    
                    # Show admin columns used
                    st.subheader("🔍 Admin Columns Used")
                    st.write(f"**Columns processed:** {results['admin_columns']}")
                    
                    # Show sample data
                    if len(results['nb_data']) > 0:
                        st.subheader("🆕 Sample New Business Data")
                        st.dataframe(results['nb_data'].head(10))
                        
                        # Download button
                        st.download_button(
                            label="Download New Business Data",
                            data=create_excel_download(results['nb_data'], 'New Business'),
                            file_name=f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        st.subheader("❌ Sample Cancellation Data")
                        st.dataframe(results['cancellation_data'].head(10))
                        
                        st.download_button(
                            label="Download Cancellation Data",
                            data=create_excel_download(results['cancellation_data'], 'Cancellations'),
                            file_name=f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        st.subheader("🔄 Sample Reinstatement Data")
                        st.dataframe(results['reinstatement_data'].head(10))
                        
                        st.download_button(
                            label="Download Reinstatement Data",
                            data=create_excel_download(results['reinstatement_data'], 'Reinstatements'),
                            file_name=f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("❌ Smart processing failed. Check the error messages above.")

if __name__ == "__main__":
    main()
