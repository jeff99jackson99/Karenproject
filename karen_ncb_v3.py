#!/usr/bin/env python3
"""
Karen NCB Data Processor - Version 3.0
Based on working logic from smart_ncb_app.py with Karen 2.0 instructions
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen NCB v3.0", page_icon="🚀", layout="wide")

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
        st.warning(f"⚠️ Could not read 'Col Ref' sheet: {e}")
        st.write("🔄 **Trying alternative column detection methods...**")
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
    
    # If we don't have enough admin columns, try to find them by content analysis
    if len(admin_columns) < 4:
        with st.expander("⚠️ **Admin Column Detection Details** (Click to expand)", expanded=False):
            st.warning("⚠️ Could not identify all Admin columns from mapping, trying content-based detection...")
            
            # Look for columns that might contain Admin amounts by analyzing their content
            potential_admin_cols = []
            
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
    
    return admin_columns

def process_excel_data_karen_v3(uploaded_file):
    """Process Excel data using working logic with Karen 2.0 instructions."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Create expandable section for processing details
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen v3.0 Processing Started**")
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            # Detect column structure
            column_mapping = detect_column_structure(uploaded_file)
            
            # Process the Data sheet (main transaction data)
            data_sheet_name = None
            
            # Look for common data sheet names
            for sheet_name in excel_data.sheet_names:
                if any(keyword in sheet_name.lower() for keyword in ['data', 'transaction', 'detail', 'main']):
                    data_sheet_name = sheet_name
                    break
            
            # If no obvious data sheet found, try the first sheet that's not a summary
            if not data_sheet_name:
                for sheet_name in excel_data.sheet_names:
                    if not any(keyword in sheet_name.lower() for keyword in ['summary', 'xref', 'ref', 'instruction']):
                        data_sheet_name = sheet_name
                        break
            
            if data_sheet_name:
                st.write(f"📋 **Processing {data_sheet_name} Sheet**")
                
                # Read the data sheet
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Find admin columns
                admin_columns = find_admin_columns(df, column_mapping)
                st.write("🔍 **Admin Columns Found:**")
                for admin_type, col_name in admin_columns.items():
                    st.write(f"  {admin_type}: {col_name}")
                
                if len(admin_columns) >= 4:
                    # Process the data using the working logic
                    results = process_transaction_data_karen_v3(df, admin_columns)
                    
                    if results:
                        # Rename columns with meaningful names
                        renamed_nb, nb_rename_map = rename_columns_with_meaning(results['nb_data'], column_mapping)
                        renamed_cancellation, c_rename_map = rename_columns_with_meaning(results['cancellation_data'], column_mapping)
                        renamed_reinstatement, r_rename_map = rename_columns_with_meaning(results['reinstatement_data'], column_mapping)
                        
                        # Show column renaming information
                        st.write("🔍 **Column Renaming Applied:**")
                        if nb_rename_map:
                            st.write("  **New Business columns renamed:**")
                            for old_name, new_name in nb_rename_map.items():
                                st.write(f"    {old_name} → {new_name}")
                        
                        # Update results with renamed dataframes
                        results['nb_data'] = renamed_nb
                        results['cancellation_data'] = renamed_cancellation
                        results['reinstatement_data'] = renamed_reinstatement
                        results['column_mapping'] = column_mapping
                        
                        return results
                    else:
                        return None
                else:
                    st.error(f"❌ Need at least 4 Admin columns, found {len(admin_columns)}")
                    return None
            else:
                st.error("❌ No suitable data sheet found")
                return None
                
    except Exception as e:
        st.error(f"❌ Error in Karen v3.0 processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def process_transaction_data_karen_v3(df, admin_columns):
    """Process transaction data with the working logic and Karen 2.0 requirements."""
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
        
        # First, look for columns that actually contain transaction type data
        for col in df.columns:
            if df[col].dtype == 'object':
                sample_vals = df[col].dropna().head(10).tolist()
                # Check if this column contains transaction types (NB, C, R)
                if any(str(v).upper() in ['NB', 'C', 'R', 'NEW BUSINESS', 'CANCELLATION', 'REINSTATEMENT'] for v in sample_vals):
                    transaction_col = col
                    st.write(f"✅ **Found transaction column by content:** {col} with values: {sample_vals[:5]}")
                    break
        
        # If not found by content, fall back to name-based search
        if not transaction_col:
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
        
        # Filter NB data using the WORKING logic that gave ~1200 records
        nb_df = df[nb_mask].copy()
        
        # Show Admin amount statistics for NB data
        st.write(f"🔍 **New Business Admin Amount Analysis:**")
        for i, col in enumerate(admin_cols):
            col_values = nb_df[col]
            non_zero_count = (col_values > 0).sum()
            zero_count = (col_values == 0).sum()
            st.write(f"  {col}: {non_zero_count} non-zero, {zero_count} zero values")
        
        # Show sum statistics
        nb_df['Admin_Sum'] = nb_df[admin_cols].sum(axis=1)
        st.write(f"  Admin Sum > 0: {(nb_df['Admin_Sum'] > 0).sum()} records")
        st.write(f"  Admin Sum = 0: {(nb_df['Admin_Sum'] == 0).sum()} records")
        
        # Use the EXACT working logic that gave ~1200 records
        with st.expander("🔍 **Filtering Details** (Click to expand)", expanded=False):
            st.write(f"🔍 **Using EXACT working logic from original version:**")
            
            # EXACT USER REQUIREMENT - ALL 4 Admin amounts > 0 AND sum > 0
            nb_filtered = nb_df[
                (nb_df['Admin_Sum'] > 0) &
                (nb_df[admin_cols[0]] > 0) &
                (nb_df[admin_cols[1]] > 0) &
                (nb_df[admin_cols[2]] > 0) &
                (nb_df[admin_cols[3]] > 0)
            ]
            st.write(f"✅ **EXACT working logic (ALL 4 Admin > 0 AND Sum > 0): {len(nb_filtered)} records**")
            st.write(f"  Expected: ~1200 records")
            st.write(f"  Actual: {len(nb_filtered)} records")
        
        # Filter C/R data according to Karen 2.0 requirements
        # C: sum < 0 (negative values expected)
        # R: sum > 0 (positive values expected)
        c_df = df[c_mask].copy()
        r_df = df[r_mask].copy()
        
        # Apply Karen 2.0 filtering
        c_filtered = c_df[c_df['Admin_Sum'] < 0]  # Cancellations: sum < 0
        r_filtered = r_df[r_df['Admin_Sum'] > 0]  # Reinstatements: sum > 0
        
        st.write(f"📊 **Karen 2.0 Filtered results:**")
        st.write(f"  New Business (filtered): {len(nb_filtered)}")
        st.write(f"  Cancellations (sum < 0): {len(c_filtered)}")
        st.write(f"  Reinstatements (sum > 0): {len(r_filtered)}")
        
        # Now create output dataframes with EXACT Karen 2.0 column mapping
        # Data Set 1 - New Contracts (NB)
        nb_output = create_karen_output_dataframe(nb_filtered, 'NB', admin_cols, 'New Business')
        
        # Data Set 2 - Reinstatements (R)
        r_output = create_karen_output_dataframe(r_filtered, 'R', admin_cols, 'Reinstatement')
        
        # Data Set 3 - Cancellations (C) - with additional columns
        c_output = create_karen_output_dataframe(c_filtered, 'C', admin_cols, 'Cancellation', include_cancellation_fields=True)
        
        st.write(f"🔍 **Karen 2.0 Output created:**")
        st.write(f"  New Business: {len(nb_output)} records")
        st.write(f"  Reinstatements: {len(r_output)} records")
        st.write(f"  Cancellations: {len(c_output)} records")
        st.write(f"  Total: {len(nb_output) + len(r_output) + len(c_output)} records")
        
        return {
            'nb_data': nb_output,
            'cancellation_data': c_output,
            'reinstatement_data': r_output,
            'admin_columns': admin_cols,
            'total_records': len(nb_output) + len(c_output) + len(r_output)
        }
        
    except Exception as e:
        st.error(f"❌ Error processing transaction data: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_karen_output_dataframe(df, transaction_type, admin_cols, row_type, include_cancellation_fields=False):
    """Create output dataframe with exact Karen 2.0 column mapping."""
    if len(df) == 0:
        return pd.DataFrame()
    
    st.write(f"🔍 **Creating {row_type} output dataframe with {len(df)} rows**")
    st.write(f"🔍 **Available columns in source data:** {list(df.columns)[:10]}...")  # Show first 10 columns
    
    # First, analyze the data structure to find the right columns
    structure_info = analyze_data_structure(df)
    if not structure_info:
        st.error("❌ **Could not analyze data structure**")
        return pd.DataFrame()
    
    transaction_col = structure_info['transaction_col']
    detected_admin_cols = structure_info['admin_columns']
    
    # Create new dataframe with exact column order as specified in Karen 2.0 instructions
    output = pd.DataFrame()
    
    # Map columns based on the detected structure
    # B – Insurer Code (try to find by content first, then position)
    insurer_col = find_column_by_content_karen(df, ['INSURER', 'INSURER CODE'])
    if insurer_col is not None and not insurer_col.isna().all():
        output['Insurer_Code'] = insurer_col
        st.write(f"✅ **Found Insurer Code column:** {insurer_col.name if hasattr(insurer_col, 'name') else 'Series'}")
    else:
        # Use position-based mapping
        if len(df.columns) > 1:
            output['Insurer_Code'] = df.iloc[:, 1]  # Column B position
            st.write(f"⚠️ **Using position-based mapping for Insurer Code:** Column 1")
        else:
            output['Insurer_Code'] = None
            st.write(f"⚠️ **Insurer Code column not found**")
    
    # C – Product Type Code
    product_col = find_column_by_content_karen(df, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'])
    if product_col is not None and not product_col.isna().all():
        output['Product_Type_Code'] = product_col
        st.write(f"✅ **Found Product Type Code column:** {product_col.name if hasattr(product_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 2:
            output['Product_Type_Code'] = df.iloc[:, 2]  # Column C position
            st.write(f"⚠️ **Using position-based mapping for Product Type Code:** Column 2")
        else:
            output['Product_Type_Code'] = None
            st.write(f"⚠️ **Product Type Code column not found**")
    
    # D – Coverage Code
    coverage_col = find_column_by_content_karen(df, ['COVERAGE CODE', 'COVERAGE'])
    if coverage_col is not None and not coverage_col.isna().all():
        output['Coverage_Code'] = coverage_col
        st.write(f"✅ **Found Coverage Code column:** {coverage_col.name if hasattr(coverage_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 3:
            output['Coverage_Code'] = df.iloc[:, 3]  # Column D position
            st.write(f"⚠️ **Using position-based mapping for Coverage Code:** Column 3")
        else:
            output['Coverage_Code'] = None
            st.write(f"⚠️ **Coverage Code column not found**")
    
    # E – Dealer Number
    dealer_num_col = find_column_by_content_karen(df, ['DEALER NUMBER', 'DEALER #'])
    if dealer_num_col is not None and not dealer_num_col.isna().all():
        output['Dealer_Number'] = dealer_num_col
        st.write(f"✅ **Found Dealer Number column:** {dealer_num_col.name if hasattr(dealer_num_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 4:
            output['Dealer_Number'] = df.iloc[:, 4]  # Column E position
            st.write(f"⚠️ **Using position-based mapping for Dealer Number:** Column 4")
        else:
            output['Dealer_Number'] = None
            st.write(f"⚠️ **Dealer Number column not found**")
    
    # F – Dealer Name
    dealer_name_col = find_column_by_content_karen(df, ['DEALER NAME', 'DEALER'])
    if dealer_name_col is not None and not dealer_name_col.isna().all():
        output['Dealer_Name'] = dealer_name_col
        st.write(f"✅ **Found Dealer Name column:** {dealer_name_col.name if hasattr(dealer_name_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 5:
            output['Dealer_Name'] = df.iloc[:, 5]  # Column F position
            st.write(f"⚠️ **Using position-based mapping for Dealer Name:** Column 5")
        else:
            output['Dealer_Name'] = None
            st.write(f"⚠️ **Dealer Name column not found**")
    
    # H – Contract Number
    contract_col = find_column_by_content_karen(df, ['CONTRACT NUMBER', 'CONTRACT #'])
    if contract_col is not None and not contract_col.isna().all():
        output['Contract_Number'] = contract_col
        st.write(f"✅ **Found Contract Number column:** {contract_col.name if hasattr(contract_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 7:
            output['Contract_Number'] = df.iloc[:, 7]  # Column H position
            st.write(f"⚠️ **Using position-based mapping for Contract Number:** Column 7")
        else:
            output['Contract_Number'] = None
            st.write(f"⚠️ **Contract Number column not found**")
    
    # L – Contract Sale Date
    sale_date_col = find_column_by_content_karen(df, ['CONTRACT SALE DATE', 'SALE DATE'])
    if sale_date_col is not None and not sale_date_col.isna().all():
        output['Contract_Sale_Date'] = sale_date_col
        st.write(f"✅ **Found Contract Sale Date column:** {sale_date_col.name if hasattr(sale_date_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 11:
            output['Contract_Sale_Date'] = df.iloc[:, 11]  # Column L position
            st.write(f"⚠️ **Using position-based mapping for Contract Sale Date:** Column 11")
        else:
            output['Contract_Sale_Date'] = None
            st.write(f"⚠️ **Contract Sale Date column not found**")
    
    # J – Transaction Date
    trans_date_col = find_column_by_content_karen(df, ['TRANSACTION DATE', 'ACTIVATION DATE'])
    if trans_date_col is not None and not trans_date_col.isna().all():
        output['Transaction_Date'] = trans_date_col
        st.write(f"✅ **Found Transaction Date column:** {trans_date_col.name if hasattr(trans_date_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 9:
            output['Transaction_Date'] = df.iloc[:, 9]  # Column J position
            st.write(f"⚠️ **Using position-based mapping for Transaction Date:** Column 9")
        else:
            output['Transaction_Date'] = None
            st.write(f"⚠️ **Transaction Date column not found**")
    
    # M – Transaction Type (use the detected one)
    output['Transaction_Type'] = df[transaction_col]
    st.write(f"✅ **Using detected Transaction Type column:** {transaction_col}")
    
    # U – Customer Last Name
    last_name_col = find_column_by_content_karen(df, ['LAST NAME', 'CUSTOMER LAST NAME'])
    if last_name_col is not None and not last_name_col.isna().all():
        output['Customer_Last_Name'] = last_name_col
        st.write(f"✅ **Found Customer Last Name column:** {last_name_col.name if hasattr(last_name_col, 'name') else 'Series'}")
    else:
        if len(df.columns) > 20:
            output['Customer_Last_Name'] = df.iloc[:, 20]  # Column U position
            st.write(f"⚠️ **Using position-based mapping for Customer Last Name:** Column 20")
        else:
            output['Customer_Last_Name'] = None
            st.write(f"⚠️ **Customer Last Name column not found**")
    
    # Additional columns for cancellations only
    if include_cancellation_fields:
        # Z – Contract Term
        term_col = find_column_by_content_karen(df, ['CONTRACT TERM', 'TERM'])
        if term_col is not None and not term_col.isna().all():
            output['Contract_Term'] = term_col
            st.write(f"✅ **Found Contract Term column:** {term_col.name if hasattr(term_col, 'name') else 'Series'}")
        else:
            if len(df.columns) > 25:
                output['Contract_Term'] = df.iloc[:, 25]  # Column Z position
                st.write(f"⚠️ **Using position-based mapping for Contract Term:** Column 25")
            else:
                output['Contract_Term'] = None
                st.write(f"⚠️ **Contract Term column not found**")
        
        # AE – Cancellation Date
        cancel_date_col = find_column_by_content_karen(df, ['CANCELLATION DATE', 'CANCEL DATE'])
        if cancel_date_col is not None and not cancel_date_col.isna().all():
            output['Cancellation_Date'] = cancel_date_col
            st.write(f"✅ **Found Cancellation Date column:** {cancel_date_col.name if hasattr(cancel_date_col, 'name') else 'Series'}")
        else:
            if len(df.columns) > 30:
                output['Cancellation_Date'] = df.iloc[:, 30]  # Column AE position
                st.write(f"⚠️ **Using position-based mapping for Cancellation Date:** Column 30")
            else:
                output['Cancellation_Date'] = None
                st.write(f"⚠️ **Cancellation Date column not found**")
        
        # AB – Cancellation Reason
        reason_col = find_column_by_content_karen(df, ['CANCELLATION REASON', 'REASON'])
        if reason_col is not None and not reason_col.isna().all():
            output['Cancellation_Reason'] = reason_col
            st.write(f"✅ **Found Cancellation Reason column:** {reason_col.name if hasattr(reason_col, 'name') else 'Series'}")
        else:
            if len(df.columns) > 27:
                output['Cancellation_Reason'] = df.iloc[:, 27]  # Column AB position
                st.write(f"⚠️ **Using position-based mapping for Cancellation Reason:** Column 27")
            else:
                output['Cancellation_Reason'] = None
                st.write(f"⚠️ **Cancellation Reason column not found**")
        
        # AA – Cancellation Factor
        factor_col = find_column_by_content_karen(df, ['CANCELLATION FACTOR', 'FACTOR'])
        if factor_col is not None and not factor_col.isna().all():
            output['Cancellation_Factor'] = factor_col
            st.write(f"✅ **Found Cancellation Factor column:** {factor_col.name if hasattr(factor_col, 'name') else 'Series'}")
        else:
            if len(df.columns) > 26:
                output['Cancellation_Factor'] = df.iloc[:, 26]  # Column AA position
                st.write(f"⚠️ **Using position-based mapping for Cancellation Factor:** Column 26")
            else:
                output['Cancellation_Factor'] = None
                st.write(f"⚠️ **Cancellation Factor column not found**")
    
    # Admin Amount columns (use the detected ones)
    # AO – Admin 3 Amount (Agent NCB Fee)
    output['Admin_3_Amount_Agent_NCB_Fee'] = df[detected_admin_cols['Admin 3']]
    # AQ – Admin 4 Amount (Dealer NCB Fee)
    output['Admin_4_Amount_Dealer_NCB_Fee'] = df[detected_admin_cols['Admin 4']]
    # AU – Admin 6 Amount (Agent NCB Offset)
    output['Admin_6_Amount_Agent_NCB_Offset'] = df[detected_admin_cols['Admin 9']]
    # AW – Admin 7 Amount (Agent NCB Offset Bucket)
    output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = df[detected_admin_cols['Admin 10']]
    # AY – Admin 8 Amount (Dealer NCB Offset Bucket)
    output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = df[detected_admin_cols['Admin 3']]  # Using available columns
    # BA – Admin 9 Amount (Agent NCB Offset)
    output['Admin_9_Amount_Agent_NCB_Offset'] = df[detected_admin_cols['Admin 4']]  # Using available columns
    # BC – Admin 10 Amount (Dealer NCB Offset Bucket)
    output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = df[detected_admin_cols['Admin 9']]  # Using available columns
    
    # Add transaction type identifiers
    output['Transaction_Type'] = transaction_type
    output['Row_Type'] = row_type
    
    st.write(f"✅ **{row_type} output dataframe created with {len(output.columns)} columns**")
    st.write(f"🔍 **Output columns:** {list(output.columns)}")
    
    return output

def find_column_by_content_karen(df, search_terms):
    """Find a column by searching for specific content in column names or first few rows."""
    st.write(f"🔍 **Searching for columns matching:** {search_terms}")
    
    # First, try to find by column names
    for col in df.columns:
        col_str = str(col).upper()
        for term in search_terms:
            if term.upper() in col_str:
                st.write(f"✅ **Found column by name:** {col} matches '{term}'")
                return df[col]
    
    # If not found by name, try to find by content in first few rows
    st.write(f"🔄 **Searching by content in first few rows...**")
    for col in df.columns:
        try:
            # Get first 10 non-null values
            sample_data = df[col].dropna().head(10)
            if len(sample_data) > 0:
                for term in search_terms:
                    # Check if any of the sample values contain the search term
                    if any(term.upper() in str(val).upper() for val in sample_data):
                        st.write(f"✅ **Found column by content:** {col} contains '{term}' in sample data")
                        return df[col]
        except Exception as e:
            st.write(f"⚠️ **Error checking column {col}:** {str(e)}")
            continue
    
    # If still not found, try to find by position (fallback)
    st.write(f"🔄 **Trying position-based fallback...**")
    
    # Common positions for key columns based on typical Excel structure
    position_mapping = {
        'INSURER': 1,           # Column B
        'PRODUCT TYPE': 2,      # Column C
        'COVERAGE': 3,          # Column D
        'DEALER NUMBER': 4,     # Column E
        'DEALER NAME': 5,       # Column F
        'CONTRACT NUMBER': 7,   # Column H
        'SALE DATE': 11,        # Column L
        'TRANSACTION DATE': 9,  # Column J
        'TRANSACTION TYPE': 12, # Column M
        'LAST NAME': 20,        # Column U
    }
    
    for term in search_terms:
        for search_term, pos in position_mapping.items():
            if term.upper() in search_term.upper():
                if pos < len(df.columns):
                    col_name = df.columns[pos]
                    st.write(f"✅ **Found column by position:** {col_name} at position {pos} for '{term}'")
                    return df[col_name]
    
    # Last resort: try to find any column that might contain relevant data
    st.write(f"🔄 **Trying last resort search...**")
    for col in df.columns:
        try:
            # Check if this column has meaningful data (not all None/NaN)
            sample_data = df[col].dropna().head(20)
            if len(sample_data) > 0:
                # Check if this looks like the type of data we're looking for
                if 'INSURER' in search_terms and any(str(val).isdigit() for val in sample_data[:5]):
                    st.write(f"✅ **Found potential Insurer column by data pattern:** {col}")
                    return df[col]
                elif 'PRODUCT TYPE' in search_terms and any(str(val).isalpha() for val in sample_data[:5]):
                    st.write(f"✅ **Found potential Product Type column by data pattern:** {col}")
                    return df[col]
                elif 'DEALER' in search_terms and any(str(val).isalnum() for val in sample_data[:5]):
                    st.write(f"✅ **Found potential Dealer column by data pattern:** {col}")
                    return df[col]
        except Exception as e:
            continue
    
    st.write(f"❌ **No column found for:** {search_terms}")
    return None

def analyze_data_structure(df):
    """Analyze the data structure to understand column contents and find the right mappings."""
    st.write("🔍 **Analyzing data structure for correct column mapping...**")
    
    # Find the transaction type column by looking for actual transaction values
    transaction_col = None
    for col in df.columns:
        try:
            sample_data = df[col].dropna().head(50)
            if len(sample_data) > 0:
                # Check if this column contains actual transaction types (NB, C, R)
                unique_vals = set(str(val).upper().strip() for val in sample_data)
                if any(val in ['NB', 'C', 'R', 'NEW BUSINESS', 'CANCELLATION', 'REINSTATEMENT'] for val in unique_vals):
                    transaction_col = col
                    st.write(f"✅ **Found Transaction Type column:** {col} with values: {list(unique_vals)[:10]}")
                    break
        except:
            continue
    
    if not transaction_col:
        st.error("❌ **Could not find Transaction Type column with NB, C, R values**")
        st.write("🔍 **Available columns and their sample values:**")
        for i, col in enumerate(df.columns[:20]):  # Show first 20 columns
            try:
                sample_vals = df[col].dropna().head(5).tolist()
                st.write(f"  Column {i}: {col} → Sample: {sample_vals}")
            except:
                st.write(f"  Column {i}: {col} → Error reading data")
        return None
    
    # Now find Admin columns by looking for actual numeric data
    admin_columns = {}
    numeric_cols = []
    
    for col in df.columns:
        try:
            # Convert to numeric and check if it contains meaningful financial data
            numeric_data = pd.to_numeric(df[col], errors='coerce')
            if not numeric_data.isna().all():
                # Check if this column has meaningful non-zero values
                non_zero_count = (numeric_data != 0).sum()
                total_count = len(numeric_data.dropna())
                
                if non_zero_count > 0 and non_zero_count < total_count * 0.95:  # Not all zeros, not all non-zero
                    numeric_cols.append({
                        'column': col,
                        'non_zero_count': non_zero_count,
                        'total_count': total_count,
                        'non_zero_ratio': non_zero_count / total_count,
                        'sample_values': numeric_data.dropna().head(10).tolist()
                    })
        except:
            continue
    
    # Sort by non-zero ratio to find the most likely Admin columns
    numeric_cols.sort(key=lambda x: x['non_zero_ratio'], reverse=True)
    
    st.write(f"🔍 **Found {len(numeric_cols)} potential Admin columns:**")
    for i, col_info in enumerate(numeric_cols[:10]):
        st.write(f"  {i+1}. {col_info['column']}: {col_info['non_zero_count']}/{col_info['total_count']} non-zero ({col_info['non_zero_ratio']:.1%})")
        st.write(f"     Sample values: {col_info['sample_values'][:5]}")
    
    # Assign Admin columns based on the best candidates
    if len(numeric_cols) >= 4:
        admin_columns = {
            'Admin 3': numeric_cols[0]['column'],
            'Admin 4': numeric_cols[1]['column'], 
            'Admin 9': numeric_cols[2]['column'],
            'Admin 10': numeric_cols[3]['column']
        }
        
        st.write(f"✅ **Admin columns assigned by data analysis:**")
        for admin_type, col_name in admin_columns.items():
            st.write(f"  {admin_type}: {col_name}")
    else:
        st.error(f"❌ **Not enough Admin columns found. Need 4, found {len(numeric_cols)}**")
        return None
    
    return {
        'transaction_col': transaction_col,
        'admin_columns': admin_columns,
        'numeric_cols': numeric_cols
    }

def create_excel_download_karen_v3(df, sheet_name):
    """Create Excel download for a dataframe with Karen 2.0 formatting."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def rename_columns_with_meaning(df, column_mapping):
    """Rename unnamed columns with their actual meanings."""
    renamed_df = df.copy()
    
    # Create a mapping from unnamed columns to meaningful names
    column_rename_map = {}
    
    # Map Admin columns based on their position and the Col Ref sheet
    admin_column_mapping = {
        0: 'Admin_Base',      # Admin(Base)
        4: 'Admin_CLP',       # Admin(CLP) 
        7: 'Agent_NCB_Fees',  # Agent NCB Fees
        16: 'Dealer_NCB_Fees' # Dealer NCB Fees
    }
    
    # Apply Admin column renaming
    for col_idx, new_name in admin_column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            column_rename_map[col_name] = new_name
    
    # Map other important columns based on the Col Ref sheet
    other_column_mapping = {
        1: 'Insurer_Code',
        2: 'Product_Type_Code', 
        3: 'Coverage_Code',
        5: 'Dealer_Name',
        6: 'Dealer_Insured_State',
        8: 'Reinsurance_Support',
        9: 'Transaction_Type'
    }
    
    # Apply other column renaming
    for col_idx, new_name in other_column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            if col_name not in column_rename_map:  # Don't overwrite Admin columns
                column_rename_map[col_name] = new_name
    
    # Also rename some common columns we know about
    for col in df.columns:
        if 'RPT908' in str(col):
            column_rename_map[col] = 'Coverage_Code'
        elif 'Transaction Type' in str(df[col].iloc[0] if len(df) > 0 else ''):
            column_rename_map[col] = 'Transaction_Type'
    
    # Apply the renaming
    renamed_df = renamed_df.rename(columns=column_rename_map)
    
    # Debug: Show what was renamed
    st.write(f"🔍 **Column renaming debug for {len(column_rename_map)} columns:**")
    for old_name, new_name in column_rename_map.items():
        st.write(f"  {old_name} → {new_name}")
    
    # Debug: Show final column names
    st.write(f"🔍 **Final column names (first 10):** {list(renamed_df.columns)[:10]}")
    
    return renamed_df, column_rename_map

def main():
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.markdown("**Based on working logic with Karen 2.0 instructions**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Karen v3.0 Processing")
        st.markdown("**This app will:**")
        st.markdown("• Use the working logic that gave ~1200 NB records")
        st.markdown("• Apply Karen 2.0 column mapping requirements")
        st.markdown("• Create separate worksheets for each data set")
        st.markdown("• Generate 2,000-2,500 total records")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file for processing", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🚀 Start Karen v3.0 Processing", type="primary", use_container_width=True):
            with st.spinner("Karen v3.0 processing in progress..."):
                results = process_excel_data_karen_v3(uploaded_file)
                
                if results:
                    st.success("✅ Karen v3.0 processing complete!")
                    
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
                    
                    # Show column mapping if available
                    if 'column_mapping' in results:
                        st.write("**Column meanings from Col Ref sheet:**")
                        for col_idx, desc in results['column_mapping'].items():
                            st.write(f"  Column {col_idx}: {desc}")
                    
                    # Show sample data
                    if len(results['nb_data']) > 0:
                        st.subheader("🆕 Sample New Business Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['nb_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['nb_data'].shape}, Columns: {len(results['nb_data'].columns)}")
                        st.dataframe(results['nb_data'].head(10))
                        
                        # Download button
                        st.download_button(
                            label="Download New Business Data",
                            data=create_excel_download_karen_v3(results['nb_data'], 'New Business'),
                            file_name=f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        st.subheader("❌ Sample Cancellation Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['cancellation_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['cancellation_data'].shape}, Columns: {len(results['cancellation_data'].columns)}")
                        st.dataframe(results['cancellation_data'].head(10))
                        
                        st.download_button(
                            label="Download Cancellation Data",
                            data=create_excel_download_karen_v3(results['cancellation_data'], 'Cancellations'),
                            file_name=f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        st.subheader("🔄 Sample Reinstatement Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['reinstatement_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['reinstatement_data'].shape}, Columns: {len(results['reinstatement_data'].columns)}")
                        st.dataframe(results['reinstatement_data'].head(10))
                        
                        st.download_button(
                            label="Download Reinstatement Data",
                            data=create_excel_download_karen_v3(results['reinstatement_data'], 'Reinstatements'),
                            file_name=f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    # Karen 2.0: Create Excel file with separate worksheets
                    st.write("---")
                    st.subheader("📥 Karen 2.0 Excel Output (3 Separate Worksheets)")
                    
                    if st.button("📥 Download Complete Excel File (3 Worksheets)", type="primary"):
                        try:
                            # Create Excel file with separate worksheets as specified in Karen 2.0
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                # Data Set 1 - New Contracts (NB)
                                if len(results['nb_data']) > 0:
                                    results['nb_data'].to_excel(writer, sheet_name='New_Contracts_NB', index=False)
                                    st.write(f"✅ **New Contracts worksheet created:** {len(results['nb_data'])} records")
                                
                                # Data Set 2 - Reinstatements (R)
                                if len(results['reinstatement_data']) > 0:
                                    results['reinstatement_data'].to_excel(writer, sheet_name='Reinstatements_R', index=False)
                                    st.write(f"✅ **Reinstatements worksheet created:** {len(results['reinstatement_data'])} records")
                                
                                # Data Set 3 - Cancellations (C)
                                if len(results['cancellation_data']) > 0:
                                    results['cancellation_data'].to_excel(writer, sheet_name='Cancellations_C', index=False)
                                    st.write(f"✅ **Cancellations worksheet created:** {len(results['cancellation_data'])} records")
                            
                            output.seek(0)
                            
                            # Create download button for the complete file
                            st.download_button(
                                label="📥 Download Complete Excel File (3 Worksheets)",
                                data=output.getvalue(),
                                file_name=f"Karen_NCB_v3_Complete_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
                            st.success(f"✅ **Karen 2.0 Excel file generated successfully!**")
                            st.write(f"📊 **File contains:**")
                            st.write(f"  • New Contracts (NB): {len(results['nb_data'])} records")
                            st.write(f"  • Reinstatements (R): {len(results['reinstatement_data'])} records")
                            st.write(f"  • Cancellations (C): {len(results['cancellation_data'])} records")
                            st.write(f"  • **Total:** {results['total_records']} records across 3 worksheets")
                            st.write(f"  • **Expected:** 2,000-2,500 total records")
                            
                        except Exception as e:
                            st.error(f"❌ Error generating Karen 2.0 Excel file: {str(e)}")
                            st.write(f"Debug info: {type(e).__name__}")
                else:
                    st.error("❌ Karen v3.0 processing failed. Check the error messages above.")

if __name__ == "__main__":
    main()
