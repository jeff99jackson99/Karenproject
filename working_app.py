#!/usr/bin/env python3
"""
NCB Data Processor - Guaranteed Working Version
"""

import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="NCB Data Processor", page_icon="📊")

def main():
    st.title("📊 NCB Data Processor")
    st.write("Process Excel files for NCB data")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"File uploaded: {uploaded_file.name}")
        
        if st.button("🚀 Process Data"):
            try:
                with st.spinner("Processing..."):
                    # Read Excel file
                    excel_data = pd.ExcelFile(uploaded_file)
                    
                    # Process each sheet
                    nb_count = 0
                    c_count = 0
                    r_count = 0
                    
                    for sheet_name in excel_data.sheet_names:
                        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
                        
                        # Look for transaction type column
                        for col in df.columns:
                            col_str = str(col).lower()
                            if 'transaction' in col_str or 'type' in col_str:
                                # Count NB records
                                nb_count += len(df[df[col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])])
                                
                                # Count cancellation records
                                c_count += len(df[df[col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])])
                                
                                # Count reinstatement records
                                r_count += len(df[df[col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])])
                                break
                    
                    st.success("✅ Processing complete!")
                    
                    # Show results
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("New Business", nb_count)
                    with col2:
                        st.metric("Cancellations", c_count)
                    with col3:
                        st.metric("Reinstatements", r_count)
                    
                    st.write(f"**Sheets processed:** {len(excel_data.sheet_names)}")
                    
            except Exception as e:
                st.error(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
