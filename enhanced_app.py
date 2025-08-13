#!/usr/bin/env python3
"""
NCB Data Processor - Enhanced Version with Data Review and Downloads
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="NCB Data Processor", page_icon="📊", layout="wide")

def process_excel_file(uploaded_file):
    """Process Excel file and extract NCB data with full data collection."""
    try:
        # Read Excel file
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Process each sheet
        all_nb_data = []
        all_cancellation_data = []
        all_reinstatement_data = []
        
        for sheet_name in excel_data.sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Look for transaction type column
            for col in df.columns:
                col_str = str(col).lower()
                if 'transaction' in col_str or 'type' in col_str:
                    # Collect NB data
                    nb_mask = df[col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
                    if nb_mask.any():
                        nb_df = df[nb_mask].copy()
                        nb_df['Source_Sheet'] = sheet_name
                        nb_df['Transaction_Type'] = 'New Business'
                        all_nb_data.append(nb_df)
                    
                    # Collect cancellation data
                    c_mask = df[col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
                    if c_mask.any():
                        c_df = df[c_mask].copy()
                        c_df['Source_Sheet'] = sheet_name
                        c_df['Transaction_Type'] = 'Cancellation'
                        all_cancellation_data.append(c_df)
                    
                    # Collect reinstatement data
                    r_mask = df[col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
                    if r_mask.any():
                        r_df = df[r_mask].copy()
                        r_df['Source_Sheet'] = sheet_name
                        r_df['Transaction_Type'] = 'Reinstatement'
                        all_reinstatement_data.append(r_df)
                    break
        
        # Combine all data
        final_nb = pd.concat(all_nb_data, ignore_index=True) if all_nb_data else pd.DataFrame()
        final_cancellation = pd.concat(all_cancellation_data, ignore_index=True) if all_cancellation_data else pd.DataFrame()
        final_reinstatement = pd.concat(all_reinstatement_data, ignore_index=True) if all_reinstatement_data else pd.DataFrame()
        
        return {
            'nb_data': final_nb,
            'cancellation_data': final_cancellation,
            'reinstatement_data': final_reinstatement,
            'sheets_processed': len(excel_data.sheet_names),
            'total_records': len(final_nb) + len(final_cancellation) + len(final_reinstatement)
        }
        
    except Exception as e:
        st.error(f"Error: {str(e)}")
        return None

def main():
    st.title("📊 NCB Data Processor - Enhanced")
    st.write("Process Excel files for NCB data with full review and download capabilities")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Configuration")
        st.markdown("**Features:**")
        st.markdown("• Full data collection")
        st.markdown("• Data review & preview")
        st.markdown("• Excel file downloads")
        st.markdown("• Record counting")
        
        st.markdown("---")
        st.markdown("**Instructions:**")
        st.markdown("1. Upload Excel file")
        st.markdown("2. Review collected data")
        st.markdown("3. Download results")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🚀 Process Data", type="primary", use_container_width=True):
            with st.spinner("Processing your Excel file..."):
                results = process_excel_file(uploaded_file)
                
                if results:
                    st.success("✅ Processing complete!")
                    
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
                    
                    # Create tabs for different views
                    tab1, tab2, tab3, tab4 = st.tabs(["📋 Data Review", "📥 Downloads", "🔍 Detailed Preview", "📊 Summary"])
                    
                    with tab1:
                        st.header("📋 Data Review")
                        st.write("Review the data collected from your Excel file:")
                        
                        if len(results['nb_data']) > 0:
                            st.subheader("🆕 New Business Data")
                            st.write(f"**Records found:** {len(results['nb_data'])}")
                            st.dataframe(results['nb_data'], use_container_width=True)
                        
                        if len(results['cancellation_data']) > 0:
                            st.subheader("❌ Cancellation Data")
                            st.write(f"**Records found:** {len(results['cancellation_data'])}")
                            st.dataframe(results['cancellation_data'], use_container_width=True)
                        
                        if len(results['reinstatement_data']) > 0:
                            st.subheader("🔄 Reinstatement Data")
                            st.write(f"**Records found:** {len(results['reinstatement_data'])}")
                            st.dataframe(results['reinstatement_data'], use_container_width=True)
                    
                    with tab2:
                        st.header("📥 Download Results")
                        st.write("Download your processed data as Excel files:")
                        
                        # Download buttons
                        if len(results['nb_data']) > 0:
                            nb_buffer = io.BytesIO()
                            with pd.ExcelWriter(nb_buffer, engine='openpyxl') as writer:
                                results['nb_data'].to_excel(writer, index=False, sheet_name='New Business Data')
                            nb_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download New Business Data",
                                data=nb_buffer.getvalue(),
                                file_name=f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
                                file_name=f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        if len(results['reinstatement_data']) > 0:
                            r_buffer = io.BytesIO()
                            with pd.ExcelWriter(r_buffer, engine='openpyxl') as writer:
                                results['reinstatement_data'].to_excel(writer, index=False, sheet_name='Reinstatement Data')
                            r_buffer.seek(0)
                            
                            st.download_button(
                                label="📥 Download Reinstatement Data",
                                data=r_buffer.getvalue(),
                                file_name=f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        
                        # Combined download
                        if results['total_records'] > 0:
                            st.markdown("---")
                            st.subheader("📦 Combined Download")
                            
                            combined_buffer = io.BytesIO()
                            with pd.ExcelWriter(combined_buffer, engine='openpyxl') as writer:
                                if len(results['nb_data']) > 0:
                                    results['nb_data'].to_excel(writer, index=False, sheet_name='New Business')
                                if len(results['cancellation_data']) > 0:
                                    results['cancellation_data'].to_excel(writer, index=False, sheet_name='Cancellations')
                                if len(results['reinstatement_data']) > 0:
                                    results['reinstatement_data'].to_excel(writer, index=False, sheet_name='Reinstatements')
                            
                            combined_buffer.seek(0)
                            st.download_button(
                                label="📥 Download All Data (Combined)",
                                data=combined_buffer.getvalue(),
                                file_name=f"NCB_All_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                    
                    with tab3:
                        st.header("🔍 Detailed Preview")
                        st.write("Preview of your data (first 10 rows of each type):")
                        
                        if len(results['nb_data']) > 0:
                            st.subheader("🆕 New Business Preview")
                            st.dataframe(results['nb_data'].head(10), use_container_width=True)
                            st.caption(f"Showing first 10 of {len(results['nb_data'])} records")
                        
                        if len(results['cancellation_data']) > 0:
                            st.subheader("❌ Cancellation Preview")
                            st.dataframe(results['cancellation_data'].head(10), use_container_width=True)
                            st.caption(f"Showing first 10 of {len(results['cancellation_data'])} records")
                        
                        if len(results['reinstatement_data']) > 0:
                            st.subheader("🔄 Reinstatement Preview")
                            st.dataframe(results['reinstatement_data'].head(10), use_container_width=True)
                            st.caption(f"Showing first 10 of {len(results['reinstatement_data'])} records")
                    
                    with tab4:
                        st.header("📊 Processing Summary")
                        st.markdown("""
                        **Data Collection Summary:**
                        - **Total Records Found:** {total}
                        - **New Business Records:** {nb}
                        - **Cancellation Records:** {c}
                        - **Reinstatement Records:** {r}
                        - **Sheets Processed:** {sheets}
                        - **Processing Time:** {time}
                        """.format(
                            total=results['total_records'],
                            nb=len(results['nb_data']),
                            c=len(results['cancellation_data']),
                            r=len(results['reinstatement_data']),
                            sheets=results['sheets_processed'],
                            time=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        ))
                        
                        # Data quality info
                        st.subheader("🔍 Data Quality Information")
                        if len(results['nb_data']) > 0:
                            st.write(f"**New Business Data Columns:** {list(results['nb_data'].columns)}")
                        if len(results['cancellation_data']) > 0:
                            st.write(f"**Cancellation Data Columns:** {list(results['cancellation_data'].columns)}")
                        if len(results['reinstatement_data']) > 0:
                            st.write(f"**Reinstatement Data Columns:** {list(results['reinstatement_data'].columns)}")

if __name__ == "__main__":
    main()
