#!/usr/bin/env python3
"""
NCB Data Processor - Streamlit Cloud Version
Simplified version for cloud deployment.
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

# Page configuration
st.set_page_config(
    page_title="NCB Data Processor",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #ff7f0e;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

def process_excel_data(uploaded_file):
    """Process the uploaded Excel file."""
    try:
        # Read the Excel file
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Process each sheet
        all_nb_data = []
        all_cancellation_data = []
        all_reinstatement_data = []
        
        for sheet_name in excel_data.sheet_names:
            sheet_data = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Look for transaction type column
            transaction_cols = [col for col in sheet_data.columns 
                              if 'transaction' in str(col).lower() or 'type' in str(col).lower()]
            
            if transaction_cols:
                for col in transaction_cols:
                    # Filter NB data
                    nb_mask = sheet_data[col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
                    if nb_mask.any():
                        nb_data = sheet_data[nb_mask].copy()
                        nb_data['Source_Sheet'] = sheet_name
                        all_nb_data.append(nb_data)
                    
                    # Filter cancellation data
                    c_mask = sheet_data[col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
                    if c_mask.any():
                        c_data = sheet_data[c_mask].copy()
                        c_data['Source_Sheet'] = sheet_name
                        all_cancellation_data.append(c_data)
                    
                    # Filter reinstatement data
                    r_mask = sheet_data[col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
                    if r_mask.any():
                        r_data = sheet_data[r_mask].copy()
                        r_data['Source_Sheet'] = sheet_name
                        all_reinstatement_data.append(r_data)
        
        # Combine data
        nb_data = pd.concat(all_nb_data, ignore_index=True) if all_nb_data else pd.DataFrame()
        cancellation_data = pd.concat(all_cancellation_data, ignore_index=True) if all_cancellation_data else pd.DataFrame()
        reinstatement_data = pd.concat(all_reinstatement_data, ignore_index=True) if all_reinstatement_data else pd.DataFrame()
        
        return {
            'nb_data': nb_data,
            'cancellation_data': cancellation_data,
            'reinstatement_data': reinstatement_data,
            'sheets_processed': len(excel_data.sheet_names)
        }
        
    except Exception as e:
        st.error(f"Error processing file: {e}")
        return None

def main():
    """Main application."""
    
    # Header
    st.markdown('<h1 class="main-header">📊 NCB Data Processor</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; font-size: 1.2rem; color: #666;">Process New Business, Cancellation, and Reinstatement data from Excel spreadsheets</p>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Configuration")
        st.markdown("**Instructions:**")
        st.markdown("1. Upload your Excel file")
        st.markdown("2. Process the data")
        st.markdown("3. Download results")
        
        st.markdown("---")
        st.markdown("**About:**")
        st.markdown("This app processes Excel files to extract:")
        st.markdown("• New Business (NB) transactions")
        st.markdown("• Cancellation (C) transactions")
        st.markdown("• Reinstatement (R) transactions")
    
    # Main content
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<h2 class="sub-header">📁 Upload Excel File</h2>', unsafe_allow_html=True)
        
        # File upload
        uploaded_file = st.file_uploader(
            "Choose an Excel file (.xlsx or .xls)",
            type=['xlsx', 'xls'],
            help="Upload your Excel file containing NCB data"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            
            # File info
            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }
            st.json(file_details)
            
            # Process button
            if st.button("🚀 Process NCB Data", type="primary", use_container_width=True):
                with st.spinner("Processing your Excel file..."):
                    results = process_excel_data(uploaded_file)
                    
                    if results:
                        st.session_state['results'] = results
                        st.session_state['processed'] = True
                        st.success("✅ Data processing completed successfully!")
                    else:
                        st.error("❌ Processing failed")
    
    with col2:
        st.markdown('<h2 class="sub-header">📊 Quick Stats</h2>', unsafe_allow_html=True)
        
        if 'processed' in st.session_state and st.session_state['processed']:
            results = st.session_state.get('results', {})
            
            st.metric("New Business Records", len(results.get('nb_data', pd.DataFrame())))
            st.metric("Cancellation Records", len(results.get('cancellation_data', pd.DataFrame())))
            st.metric("Reinstatement Records", len(results.get('reinstatement_data', pd.DataFrame())))
            st.metric("Sheets Processed", results.get('sheets_processed', 0))
        else:
            st.info("📤 Upload and process a file to see statistics")
    
    # Results section
    if 'processed' in st.session_state and st.session_state['processed']:
        st.markdown('<h2 class="sub-header">📋 Processing Results</h2>', unsafe_allow_html=True)
        
        results = st.session_state.get('results', {})
        
        # Tabs for different views
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📥 Download Files", "🔍 Data Preview", "📋 Raw Data"])
        
        with tab1:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Processing Summary**")
            st.markdown(f"- **New Business Records**: {len(results.get('nb_data', pd.DataFrame()))}")
            st.markdown(f"- **Cancellation Records**: {len(results.get('cancellation_data', pd.DataFrame()))}")
            st.markdown(f"- **Reinstatement Records**: {len(results.get('reinstatement_data', pd.DataFrame()))}")
            st.markdown(f"- **Sheets Processed**: {results.get('sheets_processed', 0)}")
            st.markdown(f"- **Processing Time**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            st.markdown("</div>", unsafe_allow_html=True)
        
        with tab2:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Download Processed Files**")
            st.markdown("Your processed files are ready for download:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Download buttons
            if len(results.get('nb_data', pd.DataFrame())) > 0:
                nb_buffer = io.BytesIO()
                with pd.ExcelWriter(nb_buffer, engine='openpyxl') as writer:
                    results['nb_data'].to_excel(writer, index=False, sheet_name='New Business Data')
                nb_buffer.seek(0)
                st.download_button(
                    label="📥 Download New Business Data",
                    data=nb_buffer.getvalue(),
                    file_name=f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            if len(results.get('cancellation_data', pd.DataFrame())) > 0:
                c_buffer = io.BytesIO()
                with pd.ExcelWriter(c_buffer, engine='openpyxl') as writer:
                    results['cancellation_data'].to_excel(writer, index=False, sheet_name='Cancellation Data')
                c_buffer.seek(0)
                st.download_button(
                    label="📥 Download Cancellation Data",
                    data=c_buffer.getvalue(),
                    file_name=f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            if len(results.get('reinstatement_data', pd.DataFrame())) > 0:
                r_buffer = io.BytesIO()
                with pd.ExcelWriter(r_buffer, engine='openpyxl') as writer:
                    results['reinstatement_data'].to_excel(writer, index=False, sheet_name='Reinstatement Data')
                r_buffer.seek(0)
                st.download_button(
                    label="📥 Download Reinstatement Data",
                    data=r_buffer.getvalue(),
                    file_name=f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        
        with tab3:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Data Preview**")
            st.markdown("Preview of processed data:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Data previews
            if len(results.get('nb_data', pd.DataFrame())) > 0:
                st.subheader("New Business Data Preview")
                st.dataframe(results['nb_data'].head(10))
                st.caption(f"Showing first 10 rows of {len(results['nb_data'])} total records")
            
            if len(results.get('cancellation_data', pd.DataFrame())) > 0:
                st.subheader("Cancellation Data Preview")
                st.dataframe(results['cancellation_data'].head(10))
                st.caption(f"Showing first 10 rows of {len(results['cancellation_data'])} total records")
            
            if len(results.get('reinstatement_data', pd.DataFrame())) > 0:
                st.subheader("Reinstatement Data Preview")
                st.dataframe(results['reinstatement_data'].head(10))
                st.caption(f"Showing first 10 rows of {len(results['reinstatement_data'])} total records")
        
        with tab4:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Raw Data Analysis**")
            st.markdown("Detailed view of your data:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Show all data
            if len(results.get('nb_data', pd.DataFrame())) > 0:
                st.subheader("All New Business Data")
                st.dataframe(results['nb_data'])
            
            if len(results.get('cancellation_data', pd.DataFrame())) > 0:
                st.subheader("All Cancellation Data")
                st.dataframe(results['cancellation_data'])
            
            if len(results.get('reinstatement_data', pd.DataFrame())) > 0:
                st.subheader("All Reinstatement Data")
                st.dataframe(results['reinstatement_data'])

if __name__ == "__main__":
    main()
