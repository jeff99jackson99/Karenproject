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

def process_excel_data_karen_v3(df):
    """Process Excel data using Karen v3.0 logic with enhanced column detection."""
    try:
        st.write("🔍 **Starting Karen v3.0 data processing...**")
        
        # First, analyze the data structure to find the right columns
        structure_info = analyze_data_structure(df)
        if not structure_info:
            st.error("❌ **Could not analyze data structure**")
            return None
        
        transaction_col = structure_info['transaction_col']
        detected_admin_cols = structure_info['admin_columns']
        
        st.write(f"✅ **Data structure analysis complete:**")
        st.write(f"  - Transaction column: {transaction_col}")
        st.write(f"  - Admin columns: {detected_admin_cols}")
        
        # Filter data by transaction type
        st.write("🔄 **Filtering data by transaction type...**")
        
        # New Business (NB) - Transaction Type = "NB"
        nb_df = df[df[transaction_col].astype(str).str.upper().str.strip() == 'NB'].copy()
        st.write(f"  - NB records found: {len(nb_df)}")
        
        # Cancellations (C) - Transaction Type = "C"
        c_df = df[df[transaction_col].astype(str).str.upper().str.strip() == 'C'].copy()
        st.write(f"  - C records found: {len(c_df)}")
        
        # Reinstatements (R) - Transaction Type = "R"
        r_df = df[df[transaction_col].astype(str).str.upper().str.strip() == 'R'].copy()
        st.write(f"  - R records found: {len(r_df)}")
        
        # Apply Admin amount filtering
        st.write("🔄 **Applying Admin amount filtering...**")
        
        # For NB and R: sum > 0
        if len(nb_df) > 0:
            nb_df['Admin_Sum'] = nb_df[list(detected_admin_cols.values())].sum(axis=1)
            nb_filtered = nb_df[nb_df['Admin_Sum'] > 0]
            st.write(f"  - NB after Admin sum > 0 filter: {len(nb_filtered)} records")
            
            # Additional filter: all individual Admin amounts must be > 0
            admin_cols_list = list(detected_admin_cols.values())
            nb_final = nb_filtered[
                (nb_filtered[admin_cols_list[0]] > 0) & 
                (nb_filtered[admin_cols_list[1]] > 0) & 
                (nb_filtered[admin_cols_list[2]] > 0) & 
                (nb_filtered[admin_cols_list[3]] > 0)
            ]
            st.write(f"  - NB after individual Admin > 0 filter: {len(nb_final)} records")
            nb_df = nb_final
        
        if len(r_df) > 0:
            r_df['Admin_Sum'] = r_df[list(detected_admin_cols.values())].sum(axis=1)
            r_filtered = r_df[r_df['Admin_Sum'] > 0]
            st.write(f"  - R after Admin sum > 0 filter: {len(r_filtered)} records")
            r_df = r_filtered
        
        # For C: sum < 0
        if len(c_df) > 0:
            c_df['Admin_Sum'] = c_df[list(detected_admin_cols.values())].sum(axis=1)
            c_filtered = c_df[c_df['Admin_Sum'] < 0]
            st.write(f"  - C after Admin sum < 0 filter: {len(c_filtered)} records")
            c_df = c_filtered
        
        # Create output dataframes with exact Karen 2.0 column mapping
        st.write("🔄 **Creating output dataframes...**")
        
        # New Business output
        nb_output = create_karen_output_dataframe(nb_df, 'NB', detected_admin_cols, 'New Business', include_cancellation_fields=False)
        
        # Cancellation output (includes additional fields)
        c_output = create_karen_output_dataframe(c_df, 'C', detected_admin_cols, 'Cancellation', include_cancellation_fields=True)
        
        # Reinstatement output
        r_output = create_karen_output_dataframe(r_df, 'R', detected_admin_cols, 'Reinstatement', include_cancellation_fields=False)
        
        st.write("✅ **Data processing complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        return nb_output, c_output, r_output
        
    except Exception as e:
        st.error(f"❌ **Error in Karen v3.0 processing:** {str(e)}")
        st.write("**Full error details:**")
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
    try:
        # Create a BytesIO buffer for the Excel file
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Get the workbook and worksheet objects
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
            
            # Apply header formatting
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                # Auto-adjust column width
                worksheet.set_column(col_num, col_num, max(len(str(value)) + 2, 15))
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

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
                    nb_df, c_df, r_df = process_excel_data_karen_v3(df)
                    
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
                                nb_download = create_excel_download_karen_v3(nb_df, "New_Contracts_NB")
                                if nb_download:
                                    st.download_button(
                                        label=f"📥 Download NB Data ({len(nb_df)} rows)",
                                        data=nb_download.getvalue(),
                                        file_name=f"NCB_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with col2:
                            if len(c_df) > 0:
                                c_download = create_excel_download_karen_v3(c_df, "Cancellations_C")
                                if c_download:
                                    st.download_button(
                                        label=f"📥 Download C Data ({len(c_df)} rows)",
                                        data=c_download.getvalue(),
                                        file_name=f"NCB_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with col3:
                            if len(r_df) > 0:
                                r_download = create_excel_download_karen_v3(r_df, "Reinstatements_R")
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
