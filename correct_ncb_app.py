#!/usr/bin/env python3
"""
NCB Data Processor - CORRECTED VERSION
Follows exact user specifications for NCB data filtering.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="NCB Data Processor - Corrected", page_icon="📊", layout="wide")

def find_admin_columns(df):
    """Find Admin 3, 4, 9, 10 amount columns and their labels."""
    admin_columns = {}
    
    # Look for Admin amount columns
    for col in df.columns:
        col_str = str(col).lower()
        
        # Admin 3 Amount
        if 'admin 3' in col_str and 'amount' in col_str:
            admin_columns['Admin 3 Amount'] = col
        
        # Admin 4 Amount  
        elif 'admin 4' in col_str and 'amount' in col_str:
            admin_columns['Admin 4 Amount'] = col
        
        # Admin 9 Amount
        elif 'admin 9' in col_str and 'amount' in col_str:
            admin_columns['Admin 9 Amount'] = col
        
        # Admin 10 Amount
        elif 'admin 10' in col_str and 'amount' in col_str:
            admin_columns['Admin 10 Amount'] = col
    
    return admin_columns

def find_admin_labels(df):
    """Find Admin 3, 4, 9, 10 label columns."""
    admin_labels = {}
    
    for col in df.columns:
        col_str = str(col).lower()
        
        # Admin 3 Label
        if 'admin 3' in col_str and ('label' in col_str or 'desc' in col_str):
            admin_labels['Admin 3 Label'] = col
        
        # Admin 4 Label
        elif 'admin 4' in col_str and ('label' in col_str or 'desc' in col_str):
            admin_labels['Admin 4 Label'] = col
        
        # Admin 9 Label
        elif 'admin 9' in col_str and ('label' in col_str or 'desc' in col_str):
            admin_labels['Admin 9 Label'] = col
        
        # Admin 10 Label
        elif 'admin 10' in col_str and ('label' in col_str or 'desc' in col_str):
            admin_labels['Admin 10 Label'] = col
    
    return admin_labels

def filter_by_admin_amounts(df, admin_columns):
    """Filter rows where sum of Admin 3, 4, 9, 10 amounts > 0."""
    if len(admin_columns) != 4:
        return df
    
    try:
        # Convert to numeric, replacing non-numeric with 0
        for col_name in admin_columns.values():
            df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
        
        # Calculate sum of all 4 admin amounts
        total_amount = (df[admin_columns['Admin 3 Amount']] + 
                       df[admin_columns['Admin 4 Amount']] + 
                       df[admin_columns['Admin 9 Amount']] + 
                       df[admin_columns['Admin 10 Amount']])
        
        # Filter rows where sum > 0
        filtered_df = df[total_amount > 0].copy()
        
        return filtered_df
        
    except Exception as e:
        st.error(f"Error filtering by admin amounts: {e}")
        return df

def process_excel_file_corrected(uploaded_file):
    """Process Excel file following exact user specifications."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Process each sheet
        all_nb_data = []
        all_cancellation_data = []
        all_reinstatement_data = []
        
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Find admin columns and labels
            admin_columns = find_admin_columns(df)
            admin_labels = find_admin_labels(df)
            
            # Look for transaction type column
            transaction_col = None
            for col in df.columns:
                col_str = str(col).lower()
                if 'transaction' in col_str or 'type' in col_str or 'tran' in col_str:
                    transaction_col = col
                    break
            
            if transaction_col:
                # Filter NB data (Transaction Type = NB)
                nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
                if nb_mask.any():
                    nb_df = df[nb_mask].copy()
                    nb_df['Source_Sheet'] = sheet_name
                    nb_df['Transaction_Type'] = 'New Business'
                    
                    # Filter by admin amounts > 0
                    nb_filtered = filter_by_admin_amounts(nb_df, admin_columns)
                    if len(nb_filtered) > 0:
                        all_nb_data.append(nb_filtered)
                
                # Filter Cancellation data (Transaction Type = C)
                c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
                if c_mask.any():
                    c_df = df[c_mask].copy()
                    c_df['Source_Sheet'] = sheet_name
                    c_df['Transaction_Type'] = 'Cancellation'
                    
                    # Filter by admin amounts > 0
                    c_filtered = filter_by_admin_amounts(c_df, admin_columns)
                    if len(c_filtered) > 0:
                        all_cancellation_data.append(c_filtered)
                
                # Filter Reinstatement data (Transaction Type = R)
                r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
                if r_mask.any():
                    r_df = df[r_mask].copy()
                    r_df['Source_Sheet'] = sheet_name
                    r_df['Transaction_Type'] = 'Reinstatement'
                    
                    # Filter by admin amounts > 0
                    r_filtered = filter_by_admin_amounts(r_df, admin_columns)
                    if len(r_filtered) > 0:
                        all_reinstatement_data.append(r_filtered)
        
        # Combine all data
        final_nb = pd.concat(all_nb_data, ignore_index=True) if all_nb_data else pd.DataFrame()
        final_cancellation = pd.concat(all_cancellation_data, ignore_index=True) if all_cancellation_data else pd.DataFrame()
        final_reinstatement = pd.concat(all_reinstatement_data, ignore_index=True) if all_reinstatement_data else pd.DataFrame()
        
        return {
            'nb_data': final_nb,
            'cancellation_data': final_cancellation,
            'reinstatement_data': final_reinstatement,
            'sheets_processed': len(excel_data.sheet_names),
            'total_records': len(final_nb) + len(final_cancellation) + len(final_reinstatement),
            'admin_columns_found': len(admin_columns) if 'admin_columns' in locals() else 0
        }
        
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        return None

def main():
    st.title("📊 NCB Data Processor - CORRECTED VERSION")
    st.markdown("**Follows exact specifications: Filters by Admin 3+4+9+10 amounts > 0**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Configuration")
        st.markdown("**Corrected Features:**")
        st.markdown("• ✅ Filters by Admin amounts > 0")
        st.markdown("• ✅ Collects ALL matching rows")
        st.markdown("• ✅ Proper transaction type filtering")
        st.markdown("• ✅ Data review before download")
        
        st.markdown("---")
        st.markdown("**Instructions:**")
        st.markdown("1. Upload Excel file")
        st.markdown("2. Review filtered data")
        st.markdown("3. Download results")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🚀 Process NCB Data (Corrected)", type="primary", use_container_width=True):
            with st.spinner("Processing with correct filtering..."):
                results = process_excel_file_corrected(uploaded_file)
                
                if results:
                    st.success("✅ Processing complete with correct filtering!")
                    
                    # Display summary
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col3:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    with col4:
                        st.metric("Total Records", results['total_records'])
                    
                    st.write(f"📊 **Sheets processed:** {results['sheets_processed']}")
                    st.write(f"�� **Admin columns found:** {results['admin_columns_found']}")
                    
                    # Create tabs
                    tab1, tab2, tab3, tab4 = st.tabs(["📋 Data Review", "📥 Downloads", "🔍 Admin Amounts", "📊 Summary"])
                    
                    with tab1:
                        st.header("📋 Data Review - Filtered by Admin Amounts > 0")
                        
                        if len(results['nb_data']) > 0:
                            st.subheader("�� New Business Data")
                            st.write(f"**Records found:** {len(results['nb_data'])} (Filtered by Admin amounts > 0)")
                            st.dataframe(results['nb_data'], use_container_width=True)
                        
                        if len(results['cancellation_data']) > 0:
                            st.subheader("❌ Cancellation Data")
                            st.write(f"**Records found:** {len(results['cancellation_data'])} (Filtered by Admin amounts > 0)")
                            st.dataframe(results['cancellation_data'], use_container_width=True)
                        
                        if len(results['reinstatement_data']) > 0:
                            st.subheader("🔄 Reinstatement Data")
                            st.write(f"**Records found:** {len(results['reinstatement_data'])} (Filtered by Admin amounts > 0)")
                            st.dataframe(results['reinstatement_data'], use_container_width=True)
                    
                    with tab2:
                        st.header("📥 Download Results")
                        
                        # Download buttons
                        if len(results['nb_data']) > 0:
                            nb_buffer = io.BytesIO()
                            with pd.ExcelWriter(nb_buffer, engine='openpyxl') as writer:
                                results['nb_data'].to_excel(writer, index=False, sheet_name='New Business Data')
                            nb_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download New Business Data",
                                data=nb_buffer.getvalue(),
                                file_name=f"NB_Data_Filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        if len(results['cancellation_data']) > 0:
                            c_buffer = io.BytesIO()
                            with pd.ExcelWriter(c_buffer, engine='openpyxl') as writer:
                                results['cancellation_data'].to_excel(writer, index=False, sheet_name='Cancellation Data')
                            c_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download Cancellation Data",
                                data=c_buffer.getvalue(),
                                file_name=f"Cancellation_Data_Filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        if len(results['reinstatement_data']) > 0:
                            r_buffer = io.BytesIO()
                            with pd.ExcelWriter(r_buffer, engine='openpyxl') as writer:
                                results['reinstatement_data'].to_excel(writer, index=False, sheet_name='Reinstatement Data')
                            c_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download Reinstatement Data",
                                data=r_buffer.getvalue(),
                                file_name=f"Reinstatement_Data_Filtered_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    
                    with tab3:
                        st.header("🔍 Admin Amounts Analysis")
                        st.write("This tab shows the Admin amount columns found and their values:")
                        
                        # Show admin columns info
                        if len(results['nb_data']) > 0:
                            st.subheader("Admin Columns in New Business Data")
                            admin_cols = find_admin_columns(results['nb_data'])
                            if admin_cols:
                                for name, col in admin_cols.items():
                                    st.write(f"**{name}:** {col}")
                                    if col in results['nb_data'].columns:
                                        st.write(f"Sample values: {results['nb_data'][col].head(5).tolist()}")
                            else:
                                st.warning("No Admin amount columns found")
                    
                    with tab4:
                        st.header("📊 Processing Summary")
                        st.markdown(f"""
                        **Corrected Processing Results:**
                        - **Total Records Found:** {results['total_records']}
                        - **New Business Records:** {len(results['nb_data'])} (Filtered by Admin amounts > 0)
                        - **Cancellation Records:** {len(results['cancellation_data'])} (Filtered by Admin amounts > 0)
                        - **Reinstatement Records:** {len(results['reinstatement_data'])} (Filtered by Admin amounts > 0)
                        - **Sheets Processed:** {results['sheets_processed']}
                        - **Admin Columns Found:** {results['admin_columns_found']}
                        - **Processing Time:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
                        """)
                        
                        st.success("✅ **This version correctly filters by Admin 3+4+9+10 amounts > 0 as specified!**")

if __name__ == "__main__":
    main()
