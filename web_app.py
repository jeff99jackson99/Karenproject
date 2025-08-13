#!/usr/bin/env python3
"""
NCB Data Processor - Web Interface
Beautiful Streamlit web application for processing NCB data from Excel files.
"""

import streamlit as st
import pandas as pd
import os
import sys
from datetime import datetime
import tempfile
import base64
from pathlib import Path

# Add current directory to path to import our modules
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Page configuration
st.set_page_config(
    page_title="NCB Data Processor",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
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
    .upload-section {
        border: 2px dashed #ccc;
        border-radius: 10px;
        padding: 2rem;
        text-align: center;
        margin: 2rem 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    """Main web application."""
    
    # Header
    st.markdown('<h1 class="main-header">📊 NCB Data Processor</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align: center; font-size: 1.2rem; color: #666;">Process New Business, Cancellation, and Reinstatement data from Excel spreadsheets</p>', unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Configuration")
        
        # GitHub Configuration
        st.subheader("GitHub Settings")
        github_token = st.text_input("GitHub Personal Access Token", type="password", help="Enter your GitHub token for uploading results")
        github_repo = st.text_input("GitHub Repository", value="jeff99jackson99/Karenproject", help="Format: username/repository")
        
        # Processing Options
        st.subheader("Processing Options")
        output_dir = st.text_input("Output Directory", value="output", help="Directory to save processed files")
        auto_upload = st.checkbox("Auto-upload to GitHub", value=False, help="Automatically upload results after processing")
        
        st.markdown("---")
        st.markdown("**Instructions:**")
        st.markdown("1. Upload your Excel file")
        st.markdown("2. Configure GitHub settings")
        st.markdown("3. Process the data")
        st.markdown("4. Download or upload results")
    
    # Main content area
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown('<h2 class="sub-header">📁 Upload Excel File</h2>', unsafe_allow_html=True)
        
        # File upload
        uploaded_file = st.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your Excel file containing NCB data"
        )
        
        if uploaded_file is not None:
            st.success(f"✅ File uploaded: {uploaded_file.name}")
            
            # Show file info
            file_details = {
                "Filename": uploaded_file.name,
                "File size": f"{uploaded_file.size / 1024:.2f} KB",
                "File type": uploaded_file.type
            }
            st.json(file_details)
            
            # Process button
            if st.button("🚀 Process NCB Data", type="primary", use_container_width=True):
                with st.spinner("Processing your Excel file..."):
                    try:
                        # Process the file
                        results = process_excel_file(uploaded_file, output_dir)
                        st.session_state['results'] = results
                        st.session_state['processed'] = True
                        st.success("✅ Data processing completed successfully!")
                        
                    except Exception as e:
                        st.error(f"❌ Error processing file: {str(e)}")
                        st.session_state['processed'] = False
    
    with col2:
        st.markdown('<h2 class="sub-header">📊 Quick Stats</h2>', unsafe_allow_html=True)
        
        if 'processed' in st.session_state and st.session_state['processed']:
            results = st.session_state.get('results', {})
            
            # Display statistics
            st.metric("New Business Records", results.get('nb_count', 0))
            st.metric("Cancellation Records", results.get('cancellation_count', 0))
            st.metric("Reinstatement Records", results.get('reinstatement_count', 0))
            
            # Total records
            total = sum([
                results.get('nb_count', 0),
                results.get('cancellation_count', 0),
                results.get('reinstatement_count', 0)
            ])
            st.metric("Total Records", total)
        else:
            st.info("📤 Upload and process a file to see statistics")
    
    # Results section
    if 'processed' in st.session_state and st.session_state['processed']:
        st.markdown('<h2 class="sub-header">📋 Processing Results</h2>', unsafe_allow_html=True)
        
        results = st.session_state.get('results', {})
        
        # Display results in tabs
        tab1, tab2, tab3, tab4 = st.tabs(["📊 Summary", "📥 Download Files", "🌐 GitHub Upload", "🔍 Data Preview"])
        
        with tab1:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Processing Summary**")
            st.markdown(f"- **New Business Records**: {results.get('nb_count', 0)}")
            st.markdown(f"- **Cancellation Records**: {results.get('cancellation_count', 0)}")
            st.markdown(f"- **Reinstatement Records**: {results.get('reinstatement_count', 0)}")
            st.markdown(f"- **Output Directory**: {results.get('output_dir', 'output')}")
            st.markdown(f"- **Processing Time**: {results.get('processing_time', 'N/A')}")
            st.markdown("</div>", unsafe_allow_html=True)
        
        with tab2:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Download Processed Files**")
            st.markdown("Your processed files are ready for download:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Download buttons for each file type
            for file_type, file_path in results.get('files', {}).items():
                if os.path.exists(file_path):
                    with open(file_path, 'rb') as f:
                        file_data = f.read()
                    
                    st.download_button(
                        label=f"📥 Download {file_type}",
                        data=file_data,
                        file_name=os.path.basename(file_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        with tab3:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Upload to GitHub**")
            st.markdown("Upload your processed results to GitHub:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            if github_token and github_repo:
                if st.button("🚀 Upload to GitHub", type="primary"):
                    with st.spinner("Uploading to GitHub..."):
                        try:
                            upload_success = upload_to_github(results.get('output_dir', 'output'), github_token, github_repo)
                            if upload_success:
                                st.success("✅ Successfully uploaded to GitHub!")
                            else:
                                st.error("❌ Failed to upload to GitHub")
                        except Exception as e:
                            st.error(f"❌ Upload error: {str(e)}")
            else:
                st.warning("⚠️ Please configure GitHub settings in the sidebar")
        
        with tab4:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown("**Data Preview**")
            st.markdown("Preview of processed data:")
            st.markdown("</div>", unsafe_allow_html=True)
            
            # Show data previews
            for data_type, data in results.get('data_previews', {}).items():
                if data is not None and len(data) > 0:
                    st.subheader(f"{data_type} Preview")
                    st.dataframe(data.head(10))
                    st.caption(f"Showing first 10 rows of {len(data)} total records")

def process_excel_file(uploaded_file, output_dir):
    """Process the uploaded Excel file."""
    start_time = datetime.now()
    
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_file_path = tmp_file.name
    
    try:
        # Import and use our NCB processor
        from main import NCBDataProcessor
        
        # Process the file
        processor = NCBDataProcessor()
        processor.process_data(tmp_file_path)
        
        # Generate output files
        processor.generate_output_files(output_dir)
        
        # Get results
        results = {
            'nb_count': len(processor.nb_data) if processor.nb_data is not None else 0,
            'cancellation_count': len(processor.cancellation_data) if processor.cancellation_data is not None else 0,
            'reinstatement_count': len(processor.reinstatement_data) if processor.reinstatement_data is not None else 0,
            'output_dir': output_dir,
            'processing_time': str(datetime.now() - start_time),
            'files': {},
            'data_previews': {}
        }
        
        # Get file paths
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if processor.nb_data is not None and len(processor.nb_data) > 0:
            results['files']['New Business Data'] = f"{output_dir}/NB_Data_{timestamp}.xlsx"
            results['data_previews']['New Business'] = processor.nb_data
        
        if processor.cancellation_data is not None and len(processor.cancellation_data) > 0:
            results['files']['Cancellation Data'] = f"{output_dir}/Cancellation_Data_{timestamp}.xlsx"
            results['data_previews']['Cancellation'] = processor.cancellation_data
        
        if processor.reinstatement_data is not None and len(processor.reinstatement_data) > 0:
            results['files']['Reinstatement Data'] = f"{output_dir}/Reinstatement_Data_{timestamp}.xlsx"
            results['data_previews']['Reinstatement'] = processor.reinstatement_data
        
        return results
        
    finally:
        # Clean up temporary file
        if os.path.exists(tmp_file_path):
            os.unlink(tmp_file_path)

def upload_to_github(output_dir, token, repo):
    """Upload processed files to GitHub."""
    try:
        from github_uploader import GitHubUploader
        
        uploader = GitHubUploader(token, repo)
        return uploader.upload_directory(output_dir, "NCB Data Processing Results via Web Interface")
        
    except Exception as e:
        st.error(f"GitHub upload error: {e}")
        return False

if __name__ == "__main__":
    main()
