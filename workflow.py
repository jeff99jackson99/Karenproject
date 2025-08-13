#!/usr/bin/env python3
"""
Complete workflow script for NCB Data Processing
Combines data processing and GitHub upload in one script.
"""

import os
import sys
import argparse
from datetime import datetime
import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('workflow.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

def run_workflow(input_file, output_dir, github_token=None, github_repo=None):
    """Run the complete NCB data processing workflow."""
    logger.info("Starting NCB Data Processing Workflow")
    
    try:
        # Step 1: Process the Excel file
        logger.info("Step 1: Processing Excel file...")
        from main import NCBDataProcessor
        
        processor = NCBDataProcessor()
        processor.process_data(input_file)
        processor.generate_output_files(output_dir)
        
        logger.info("✓ Data processing completed")
        
        # Step 2: Upload to GitHub if credentials provided
        if github_token and github_repo:
            logger.info("Step 2: Uploading to GitHub...")
            from github_uploader import GitHubUploader
            
            uploader = GitHubUploader(github_token, github_repo)
            if uploader.upload_directory(output_dir, "NCB Data Processing Results"):
                logger.info("✓ GitHub upload completed")
            else:
                logger.warning("⚠ GitHub upload failed")
        else:
            logger.info("Step 2: Skipping GitHub upload (no credentials provided)")
        
        logger.info("✓ Workflow completed successfully!")
        return True
        
    except Exception as e:
        logger.error(f"Workflow failed: {e}")
        return False

def main():
    """Main function for the workflow script."""
    parser = argparse.ArgumentParser(description='Complete NCB Data Processing Workflow')
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('--output-dir', default='output', help='Output directory for processed files')
    parser.add_argument('--github-token', help='GitHub personal access token')
    parser.add_argument('--github-repo', help='GitHub repository name (owner/repo)')
    
    args = parser.parse_args()
    
    # Check if input file exists
    if not os.path.exists(args.input_file):
        logger.error(f"Input file not found: {args.input_file}")
        sys.exit(1)
    
    # Run workflow
    success = run_workflow(
        args.input_file,
        args.output_dir,
        args.github_token,
        args.github_repo
    )
    
    if success:
        logger.info("Workflow completed successfully!")
        sys.exit(0)
    else:
        logger.error("Workflow failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
