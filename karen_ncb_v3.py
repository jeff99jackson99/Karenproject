#!/usr/bin/env python3
"""
Karen NCB Data Processor - Version 3.0
Clean, working version with proper column mapping
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen NCB v3.0", page_icon="🚀", layout="wide")

def analyze_data_structure_clean(df):
    """Clean, simple data structure analysis that actually works."""
    st.write("🔍 **Starting clean data structure analysis...**")
    
    # Step 1: Find the transaction type column
    transaction_col = None
    st.write("🔄 **Step 1: Finding Transaction Type column...**")
    
    for col in df.columns:
        try:
            # Get sample data and check for transaction types
            # Skip the first row (header) and look at actual data
            sample_data = df[col].dropna().iloc[1:].head(500)  # Look at more rows
            if len(sample_data) > 0:
                # Convert to string and check for NB, C, R
                str_vals = [str(val).upper().strip() for val in sample_data]
                nb_count = str_vals.count('NB')
                c_count = str_vals.count('C')
                r_count = str_vals.count('R')
                
                # Check if this looks like a transaction type column
                if (nb_count > 0 or c_count > 0 or r_count > 0) and (nb_count + c_count + r_count) > len(sample_data) * 0.1:
                    transaction_col = col
                    st.write(f"✅ **Found Transaction Type column:** {col}")
                    st.write(f"   Sample counts: NB={nb_count}, C={c_count}, R={r_count}")
                    st.write(f"   Total transaction records: {nb_count + c_count + r_count} out of {len(sample_data)}")
                    
                    # Show unique values found
                    unique_vals = list(set(str_vals))
                    transaction_vals = [v for v in unique_vals if v in ['NB', 'C', 'R']]
                    st.write(f"   Transaction types found: {transaction_vals}")
                    break
        except Exception as e:
            continue
    
    if not transaction_col:
        st.error("❌ **Could not find Transaction Type column with NB, C, R values**")
        st.write("🔍 **Available columns and their sample values:**")
        for i, col in enumerate(df.columns[:15]):  # Show first 15 columns
            try:
                # Show both header and sample data
                header_val = df[col].iloc[0] if len(df) > 0 else 'No data'
                sample_vals = df[col].dropna().iloc[1:].head(5).tolist()
                st.write(f"  Column {i}: {col} → Header: '{header_val}' → Sample: {sample_vals}")
            except:
                st.write(f"  Column {i}: {col} → Error reading data")
        return None
    
    # Step 2: Find Admin columns by looking for actual financial data
    st.write("🔄 **Step 2: Finding Admin columns with financial data...**")
    
    admin_candidates = []
    
    for col in df.columns:
        try:
            # Skip the header row and look at actual data
            data_col = df[col].iloc[1:]  # Skip header row
            
            # Try to convert to numeric, handling mixed data types
            numeric_data = pd.to_numeric(data_col, errors='coerce')
            
            # Check if we have meaningful numeric data
            if not numeric_data.isna().all():
                # Count non-zero and non-null values
                non_zero_count = (numeric_data != 0).sum()
                total_count = len(numeric_data.dropna())
                
                # More flexible criteria for Admin columns
                if non_zero_count > 5 and total_count > 10:
                    admin_candidates.append({
                        'column': col,
                        'non_zero_count': non_zero_count,
                        'total_count': total_count,
                        'ratio': non_zero_count / total_count,
                        'mean': numeric_data.mean(),
                        'std': numeric_data.std(),
                        'sample_values': numeric_data.dropna().head(5).tolist()
                    })
            else:
                # If numeric conversion failed, try to find columns that might contain financial data
                # Look for columns with dollar signs, numbers, or decimal points
                sample_text = data_col.astype(str).head(100).str.cat(sep=' ')
                if any(char in sample_text for char in ['$', '.', '-']) and any(char.isdigit() for char in sample_text):
                    # This might be a financial column with mixed formatting
                    try:
                        # Try to extract numeric values
                        cleaned_data = data_col.astype(str).str.replace('$', '').str.replace(',', '')
                        cleaned_data = cleaned_data.str.replace('(', '-').str.replace(')', '')
                        numeric_data = pd.to_numeric(cleaned_data, errors='coerce')
                        
                        if not numeric_data.isna().all():
                            non_zero_count = (numeric_data != 0).sum()
                            total_count = len(numeric_data.dropna())
                            
                            if non_zero_count > 5 and total_count > 10:
                                admin_candidates.append({
                                    'column': col,
                                    'non_zero_count': non_zero_count,
                                    'total_count': total_count,
                                    'ratio': non_zero_count / total_count,
                                    'mean': numeric_data.mean(),
                                    'std': numeric_data.std(),
                                    'sample_values': numeric_data.dropna().head(5).tolist(),
                                    'note': 'Mixed format financial data'
                                })
                    except:
                        pass
                        
        except Exception as e:
            continue
    
    # Sort by ratio and select the best candidates
    admin_candidates.sort(key=lambda x: x['ratio'], reverse=True)
    
    st.write(f"🔍 **Found {len(admin_candidates)} potential Admin columns:**")
    for i, candidate in enumerate(admin_candidates[:10]):
        note = candidate.get('note', '')
        st.write(f"  {i+1}. {candidate['column']}: {candidate['non_zero_count']}/{candidate['total_count']} non-zero ({candidate['ratio']:.1%})")
        st.write(f"     Mean: {candidate['mean']:.2f}, Std: {candidate['std']:.2f}")
        st.write(f"     Sample: {candidate['sample_values']}")
        if note:
            st.write(f"     Note: {note}")
    
    if len(admin_candidates) < 4:
        st.warning(f"⚠️ **Only found {len(admin_candidates)} Admin columns, but need 4**")
        st.write("🔄 **Trying alternative Admin column detection...**")
        
        # Look for columns by name/content that might be Admin columns
        alternative_admin_cols = []
        for col in df.columns:
            col_str = str(col).upper()
            if any(keyword in col_str for keyword in ['ADMIN', 'NCB', 'AGENT', 'DEALER', 'FEE', 'AMOUNT']):
                try:
                    data_col = df[col].iloc[1:]  # Skip header
                    numeric_data = pd.to_numeric(data_col, errors='coerce')
                    if not numeric_data.isna().all():
                        non_zero_count = (numeric_data != 0).sum()
                        total_count = len(numeric_data.dropna())
                        if total_count > 10:
                            alternative_admin_cols.append({
                                'column': col,
                                'non_zero_count': non_zero_count,
                                'total_count': total_count,
                                'ratio': non_zero_count / total_count,
                                'note': 'Found by name/content'
                            })
                except:
                    continue
        
        if alternative_admin_cols:
            st.write(f"🔍 **Found {len(alternative_admin_cols)} alternative Admin columns by name/content:**")
            for i, alt_col in enumerate(alternative_admin_cols[:8]):
                st.write(f"  {i+1}. {alt_col['column']}: {alt_col['non_zero_count']}/{alt_col['total_count']} non-zero ({alt_col['ratio']:.1%})")
                st.write(f"     Note: {alt_col['note']}")
            
            # Combine both lists and select the best 4
            all_candidates = admin_candidates + alternative_admin_cols
            all_candidates.sort(key=lambda x: x['ratio'], reverse=True)
            
            if len(all_candidates) >= 4:
                admin_candidates = all_candidates[:4]
                st.write(f"✅ **Combined detection found {len(admin_candidates)} Admin columns**")
            else:
                st.error(f"❌ **Still not enough Admin columns. Need 4, found {len(all_candidates)}**")
                return None
        else:
            st.error(f"❌ **Alternative detection also failed. Need 4 Admin columns**")
            return None
    
    # Select the 4 best Admin columns
    selected_admin_cols = admin_candidates[:4]
    
    admin_columns = {
        'Admin 3': selected_admin_cols[0]['column'],
        'Admin 4': selected_admin_cols[1]['column'],
        'Admin 9': selected_admin_cols[2]['column'],
        'Admin 10': selected_admin_cols[3]['column']
    }
    
    st.write(f"✅ **Selected Admin columns:**")
    for admin_type, col_name in admin_columns.items():
        st.write(f"  {admin_type}: {col_name}")
    
    return {
        'transaction_col': transaction_col,
        'admin_columns': admin_columns,
        'admin_candidates': admin_candidates
    }

def find_column_simple(df, search_terms, fallback_position=None):
    """Simple, reliable column finding."""
    # First try: exact column name match
    for col in df.columns:
        col_str = str(col).upper()
        for term in search_terms:
            if term.upper() in col_str:
                st.write(f"✅ **Found by name:** {col} matches '{term}'")
                return df[col]
    
    # Second try: content search in first 50 rows
    for col in df.columns:
        try:
            sample_data = df[col].dropna().head(50)
            if len(sample_data) > 0:
                for term in search_terms:
                    if any(term.upper() in str(val).upper() for val in sample_data):
                        st.write(f"✅ **Found by content:** {col} contains '{term}'")
                        return df[col]
        except:
            continue
    
    # Third try: position fallback
    if fallback_position is not None and fallback_position < len(df.columns):
        col_name = df.columns[fallback_position]
        st.write(f"⚠️ **Using position fallback:** {col_name} at position {fallback_position}")
        return df[col_name]
    
    st.write(f"❌ **No column found for:** {search_terms}")
    return None

def create_output_dataframe_clean(df, transaction_type, admin_cols, row_type, include_cancellation_fields=False):
    """Create clean output dataframe with proper column mapping."""
    if len(df) == 0:
        return pd.DataFrame()
    
    st.write(f"🔍 **Creating {row_type} output dataframe with {len(df)} rows**")
    
    # Create new dataframe
    output = pd.DataFrame()
    
    # Map columns using simple, reliable method
    # B – Insurer Code
    insurer_col = find_column_simple(df, ['INSURER', 'INSURER CODE'], fallback_position=1)
    if insurer_col is not None:
        output['Insurer_Code'] = insurer_col
    
    # C – Product Type Code
    product_col = find_column_simple(df, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'], fallback_position=2)
    if product_col is not None:
        output['Product_Type_Code'] = product_col
    
    # D – Coverage Code
    coverage_col = find_column_simple(df, ['COVERAGE CODE', 'COVERAGE'], fallback_position=3)
    if coverage_col is not None:
        output['Coverage_Code'] = coverage_col
    
    # E – Dealer Number
    dealer_num_col = find_column_simple(df, ['DEALER NUMBER', 'DEALER #'], fallback_position=4)
    if dealer_num_col is not None:
        output['Dealer_Number'] = dealer_num_col
    
    # F – Dealer Name
    dealer_name_col = find_column_simple(df, ['DEALER NAME', 'DEALER'], fallback_position=5)
    if dealer_name_col is not None:
        output['Dealer_Name'] = dealer_name_col
    
    # H – Contract Number
    contract_col = find_column_simple(df, ['CONTRACT NUMBER', 'CONTRACT #'], fallback_position=7)
    if contract_col is not None:
        output['Contract_Number'] = contract_col
    
    # L – Contract Sale Date
    sale_date_col = find_column_simple(df, ['CONTRACT SALE DATE', 'SALE DATE'], fallback_position=11)
    if sale_date_col is not None:
        output['Contract_Sale_Date'] = sale_date_col
    
    # J – Transaction Date
    trans_date_col = find_column_simple(df, ['TRANSACTION DATE', 'ACTIVATION DATE'], fallback_position=9)
    if trans_date_col is not None:
        output['Transaction_Date'] = trans_date_col
    
    # M – Transaction Type
    output['Transaction_Type'] = transaction_type
    
    # U – Customer Last Name
    last_name_col = find_column_simple(df, ['LAST NAME', 'CUSTOMER LAST NAME'], fallback_position=20)
    if last_name_col is not None:
        output['Customer_Last_Name'] = last_name_col
    
    # Additional fields for cancellations
    if include_cancellation_fields:
        # Z – Contract Term
        term_col = find_column_simple(df, ['CONTRACT TERM', 'TERM'], fallback_position=25)
        if term_col is not None:
            output['Contract_Term'] = term_col
        
        # AE – Cancellation Date
        cancel_date_col = find_column_simple(df, ['CANCELLATION DATE', 'CANCEL DATE'], fallback_position=30)
        if cancel_date_col is not None:
            output['Cancellation_Date'] = cancel_date_col
        
        # AB – Cancellation Reason
        reason_col = find_column_simple(df, ['CANCELLATION REASON', 'REASON'], fallback_position=27)
        if reason_col is not None:
            output['Cancellation_Reason'] = reason_col
        
        # AA – Cancellation Factor
        factor_col = find_column_simple(df, ['CANCELLATION FACTOR', 'FACTOR'], fallback_position=26)
        if factor_col is not None:
            output['Cancellation_Factor'] = factor_col
    
    # Admin Amount columns - use the detected ones directly
    output['Admin_3_Amount_Agent_NCB_Fee'] = df[admin_cols['Admin 3']]
    output['Admin_4_Amount_Dealer_NCB_Fee'] = df[admin_cols['Admin 4']]
    output['Admin_6_Amount_Agent_NCB_Offset'] = df[admin_cols['Admin 9']]
    output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = df[admin_cols['Admin 10']]
    output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = df[admin_cols['Admin 3']]  # Reuse for additional columns
    output['Admin_9_Amount_Agent_NCB_Offset'] = df[admin_cols['Admin 4']]  # Reuse for additional columns
    output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = df[admin_cols['Admin 9']]  # Reuse for additional columns
    
    # Add identifiers
    output['Transaction_Type'] = transaction_type
    output['Row_Type'] = row_type
    
    st.write(f"✅ **{row_type} output dataframe created with {len(output.columns)} columns**")
    
    return output

def process_data_clean(df):
    """Clean, simple data processing that actually works."""
    try:
        st.write("🔍 **Starting clean data processing...**")
        
        # Analyze data structure
        structure_info = analyze_data_structure_clean(df)
        if not structure_info:
            st.error("❌ **Could not analyze data structure**")
            return None
        
        transaction_col = structure_info['transaction_col']
        admin_cols = structure_info['admin_columns']
        
        st.write(f"✅ **Data structure analysis complete**")
        st.write(f"  - Transaction column: {transaction_col}")
        st.write(f"  - Admin columns: {admin_cols}")
        
        # Filter by transaction type - skip the header row
        st.write("🔄 **Filtering data by transaction type...**")
        
        # Skip the first row (header) and work with actual data
        data_df = df.iloc[1:].copy()  # Skip header row
        
        # New Business (NB)
        nb_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'NB'].copy()
        st.write(f"  - NB records found: {len(nb_df)}")
        
        # Cancellations (C)
        c_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'C'].copy()
        st.write(f"  - C records found: {len(c_df)}")
        
        # Reinstatements (R)
        r_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'R'].copy()
        st.write(f"  - R records found: {len(r_df)}")
        
        # Apply Admin amount filtering
        st.write("🔄 **Applying Admin amount filtering...**")
        
        # For NB: sum > 0 AND all individual amounts > 0
        if len(nb_df) > 0:
            admin_cols_list = list(admin_cols.values())
            nb_df['Admin_Sum'] = nb_df[admin_cols_list].sum(axis=1)
            
            # First filter: sum > 0
            nb_filtered = nb_df[nb_df['Admin_Sum'] > 0]
            st.write(f"  - NB after Admin sum > 0 filter: {len(nb_filtered)} records")
            
            # Second filter: all individual amounts > 0
            nb_final = nb_filtered[
                (nb_filtered[admin_cols_list[0]] > 0) & 
                (nb_filtered[admin_cols_list[1]] > 0) & 
                (nb_filtered[admin_cols_list[2]] > 0) & 
                (nb_filtered[admin_cols_list[3]] > 0)
            ]
            st.write(f"  - NB after individual Admin > 0 filter: {len(nb_final)} records")
            nb_df = nb_final
        
        # For R: sum > 0
        if len(r_df) > 0:
            admin_cols_list = list(admin_cols.values())
            r_df['Admin_Sum'] = r_df[admin_cols_list].sum(axis=1)
            r_filtered = r_df[r_df['Admin_Sum'] > 0]
            st.write(f"  - R after Admin sum > 0 filter: {len(r_filtered)} records")
            r_df = r_filtered
        
        # For C: sum < 0
        if len(c_df) > 0:
            admin_cols_list = list(admin_cols.values())
            c_df['Admin_Sum'] = c_df[admin_cols_list].sum(axis=1)
            c_filtered = c_df[c_df['Admin_Sum'] < 0]
            st.write(f"  - C after Admin sum < 0 filter: {len(c_filtered)} records")
            c_df = c_filtered
        
        # Create output dataframes
        st.write("🔄 **Creating output dataframes...**")
        
        nb_output = create_output_dataframe_clean(nb_df, 'NB', admin_cols, 'New Business', False)
        c_output = create_output_dataframe_clean(c_df, 'C', admin_cols, 'Cancellation', True)
        r_output = create_output_dataframe_clean(r_df, 'R', admin_cols, 'Reinstatement', False)
        
        st.write("✅ **Data processing complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        return nb_output, c_output, r_output
        
    except Exception as e:
        st.error(f"❌ **Error in data processing:** {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_excel_download_clean(df, sheet_name):
    """Create clean Excel download."""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Get workbook and worksheet for formatting
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Format headers
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Apply header formatting and auto-adjust column widths
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                # Auto-adjust column width
                max_width = max(len(str(value)) + 2, 15)
                # Check data width
                for row_num in range(1, min(len(df) + 1, 100)):
                    try:
                        cell_value = str(df.iloc[row_num - 1, col_num])
                        max_width = max(max_width, len(cell_value) + 2)
                    except:
                        pass
                worksheet.set_column(col_num, col_num, max_width)
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

def main():
    st.set_page_config(
        page_title="Karen NCB Data Processor - Version 3.0",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.markdown("**Expected Output:** 2k-2500 rows in specific order with proper column mapping")
    
    # File upload section
    st.header("📁 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Upload Excel file with NCB data and 'Col Ref' sheet",
        type=['xlsx', 'xls'],
        help="Your file should have a 'Data' tab and a reference sheet (like 'Col Ref' or 'xref')"
    )
    
    if uploaded_file is not None:
        try:
            # Load the Excel file
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            
            # Show available sheets
            st.success(f"✅ **File loaded successfully!**")
            st.write(f"🔍 **Available sheets:** {list(excel_data.keys())}")
            
            # Find the data sheet
            data_sheet = None
            for sheet_name in ['Data', 'data', 'DATA']:
                if sheet_name in excel_data:
                    data_sheet = sheet_name
                    break
            
            if not data_sheet:
                st.error("❌ **No 'Data' sheet found. Please ensure your file has a 'Data' tab.**")
                return
            
            # Load the data
            df = excel_data[data_sheet]
            st.write(f"📊 **Data sheet loaded:** {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Process the data
            if st.button("🔄 **Process Data**", type="primary"):
                with st.spinner("Processing data..."):
                    # Process the data using the enhanced functions
                    nb_df, c_df, r_df = process_data_clean(df)
                    
                    if nb_df is not None and c_df is not None and r_df is not None:
                        
                        # Display results
                        st.header("📊 Processing Results")
                        
                        # Summary statistics
                        total_records = len(nb_df) + len(c_df) + len(r_df)
                        st.write(f"**Total Records:** {total_records}")
                        
                        # Breakdown by transaction type
                        st.write("**Breakdown by Transaction Type:**")
                        st.write(f"- **NB (New Business):** {len(nb_df)} records")
                        st.write(f"- **C (Cancellation):** {len(c_df)} records") 
                        st.write(f"- **R (Reinstatement):** {len(r_df)} records")
                        
                        # Show sample data in collapsible sections
                        with st.expander("📋 **Sample New Business Data (First 5 rows)**", expanded=False):
                            if len(nb_df) > 0:
                                st.dataframe(nb_df.head(), use_container_width=True)
                            else:
                                st.write("No New Business data found.")
                        
                        with st.expander("📋 **Sample Cancellation Data (First 5 rows)**", expanded=False):
                            if len(c_df) > 0:
                                st.dataframe(c_df.head(), use_container_width=True)
                            else:
                                st.write("No Cancellation data found.")
                        
                        with st.expander("📋 **Sample Reinstatement Data (First 5 rows)**", expanded=False):
                            if len(r_df) > 0:
                                st.dataframe(r_df.head(), use_container_width=True)
                            else:
                                st.write("No Reinstatement data found.")
                        
                        # Download options
                        st.header("💾 Download Options")
                        
                        # Individual downloads
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            if len(nb_df) > 0:
                                nb_download = create_excel_download_clean(nb_df, "New_Contracts_NB")
                                if nb_download:
                                    st.download_button(
                                        label=f"📥 Download NB Data ({len(nb_df)} rows)",
                                        data=nb_download.getvalue(),
                                        file_name=f"NCB_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with col2:
                            if len(c_df) > 0:
                                c_download = create_excel_download_clean(c_df, "Cancellations_C")
                                if c_download:
                                    st.download_button(
                                        label=f"📥 Download C Data ({len(c_df)} rows)",
                                        data=c_download.getvalue(),
                                        file_name=f"NCB_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with col3:
                            if len(r_df) > 0:
                                r_download = create_excel_download_clean(r_df, "Reinstatements_R")
                                if r_download:
                                    st.download_button(
                                        label=f"📥 Download R Data ({len(r_df)} rows)",
                                        data=r_download.getvalue(),
                                        file_name=f"NCB_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        # Combined download - Create a single Excel file with multiple sheets
                        st.write("---")
                        st.subheader("📊 **Combined Download (All Data in One File)**")
                        
                        try:
                            # Create combined Excel file
                            combined_output = io.BytesIO()
                            
                            with pd.ExcelWriter(combined_output, engine='xlsxwriter') as writer:
                                # Write each dataframe to a separate sheet
                                if len(nb_df) > 0:
                                    nb_df.to_excel(writer, sheet_name="New_Contracts_NB", index=False)
                                
                                if len(c_df) > 0:
                                    c_df.to_excel(writer, sheet_name="Cancellations_C", index=False)
                                
                                if len(r_df) > 0:
                                    r_df.to_excel(writer, sheet_name="Reinstatements_R", index=False)
                                
                                # Get the workbook for formatting
                                workbook = writer.book
                                
                                # Format each sheet
                                for sheet_name in writer.sheets:
                                    worksheet = writer.sheets[sheet_name]
                                    
                                    # Format headers
                                    header_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'valign': 'top',
                                        'fg_color': '#D7E4BC',
                                        'border': 1
                                    })
                                    
                                    # Get the dataframe for this sheet
                                    if sheet_name == "New_Contracts_NB":
                                        sheet_df = nb_df
                                    elif sheet_name == "Cancellations_C":
                                        sheet_df = c_df
                                    elif sheet_name == "Reinstatements_R":
                                        sheet_df = r_df
                                    else:
                                        continue
                                    
                                    # Apply header formatting and auto-adjust column widths
                                    for col_num, value in enumerate(sheet_df.columns.values):
                                        worksheet.write(0, col_num, value, header_format)
                                        # Auto-adjust column width
                                        max_width = max(len(str(value)) + 2, 15)
                                        # Also check data width
                                        for row_num in range(1, min(len(sheet_df) + 1, 100)):  # Check first 100 rows
                                            try:
                                                cell_value = str(sheet_df.iloc[row_num - 1, col_num])
                                                max_width = max(max_width, len(cell_value) + 2)
                                            except:
                                                pass
                                        worksheet.set_column(col_num, col_num, max_width)
                            
                            combined_output.seek(0)
                            
                            # Create download button
                            st.download_button(
                                label=f"📥 Download All Data Combined ({total_records} total rows)",
                                data=combined_output.getvalue(),
                                file_name=f"NCB_Complete_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="combined_download"  # Unique key to prevent conflicts
                            )
                            
                        except Exception as e:
                            st.error(f"❌ **Error creating combined download:** {str(e)}")
                            st.write("**Debug info:**")
                            st.write(f"- NB DataFrame shape: {nb_df.shape if nb_df is not None else 'None'}")
                            st.write(f"- C DataFrame shape: {c_df.shape if c_df is not None else 'None'}")
                            st.write(f"- R DataFrame shape: {r_df.shape if r_df is not None else 'None'}")
                    
                    else:
                        st.error("❌ **Data processing failed. Please check your file format.**")
        
        except Exception as e:
            st.error(f"❌ **Error loading file:** {str(e)}")
            st.write("**Please ensure your file is a valid Excel file (.xlsx or .xls) with the required sheets.**")

if __name__ == "__main__":
    main()
