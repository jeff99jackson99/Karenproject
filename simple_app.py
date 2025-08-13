#!/usr/bin/env python3
"""
NCB Data Processor - Simple Working Version
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Page configuration
st.set_page_config(page_title="NCB Data Processor", page_icon="📊", layout="wide")

def process_excel_file(uploaded_file):
    """Process Excel file and extract NCB data."""
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Process each sheet
        nb_data = []
        cancellation_data = []
        reinstatement_data = []
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Look for transaction type column
            for col in df.columns:
                col_str = str(col).lower()
                if 'transaction' in col_str or 'type' in col_str:
                    # Filter NB data
                    nb_mask = df[col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
                    if nb_mask.any():
                        nb_df = df[nb_mask].copy()
                        nb_df['Source_Sheet'] = sheet_name
                        nb_data.append(nb_df)
                    
                    # Filter cancellation data
                    c_mask = df[col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
                    if c_mask.any():
                        c_df = df[c_mask].copy()
                        c_df['Source_Sheet'] = sheet_name
                        cancellation_data.append(c_df)
                    
                    # Filter reinstatement data
                    r_mask = df[col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
                    if r_mask.any():
                        r_df = df[r_mask].copy()
                        r_df['Source_Sheet'] = sheet_name
                        reinstatement_data.append(r_df)
                    break
        
        # Combine all data
        final_nb = pd.concat(nb_data, ignore_index=True) if nb_data else pd.DataFrame()
        final_cancellation = pd.concat(cancellation_data, ignore_index=True) if cancellation_data else pd.DataFrame()
        final_reinstatement = pd.concat(reinstatement_data, ignore_index=True) if reinstatement_data else pd.DataFrame()
        
        return {
            'nb_data': final_nb,
            'cancellation_data': final_cancellation,
            'reinstatement_data': final_reinstatement,
            'sheets_processed': len(excel_file.sheet_names)
        }
        
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def main():
    st.title("📊 NCB Data Processor")
    st.markdown("Process New Business, Cancellation, and Reinstatement data from Excel files")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        st.write(f"Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🚀 Process Data", type="primary"):
            with st.spinner("Processing..."):
                results = process_excel_file(uploaded_file)
                
                if results:
                    st.success("✅ Processing complete!")
                    
                    # Display results
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col3:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    
                    # Download buttons
                    st.subheader("📥 Download Results")
                    
                    if len(results['nb_data']) > 0:
                        nb_buffer = io.BytesIO()
                        results['nb_data'].to_excel(nb_buffer, index=False)
                        nb_buffer.seek(0)
                        st.download_button(
                            "Download New Business Data",
                            nb_buffer.getvalue(),
                            f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        c_buffer = io.BytesIO()
                        results['cancellation_data'].to_excel(c_buffer, index=False)
                        c_buffer.seek(0)
                        st.download_button(
                            "Download Cancellation Data",
                            c_buffer.getvalue(),
                            f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        r_buffer = io.BytesIO()
                        results['reinstatement_data'].to_excel(r_buffer, index=False)
                        r_buffer.seek(0)
                        st.download_button(
                            "Download Reinstatement Data",
                            r_buffer.getvalue(),
                            f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        )
                    
                    # Data previews
                    st.subheader("🔍 Data Preview")
                    
                    if len(results['nb_data']) > 0:
                        st.write("**New Business Data:**")
                        st.dataframe(results['nb_data'].head(5))
                    
                    if len(results['cancellation_data']) > 0:
                        st.write("**Cancellation Data:**")
                        st.dataframe(results['cancellation_data'].head(5))
                    
                    if len(results['reinstatement_data']) > 0:
                        st.write("**Reinstatement Data:**")
                        st.dataframe(results['reinstatement_data'].head(5))

if __name__ == "__main__":
    main()
