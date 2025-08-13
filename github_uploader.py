#!/usr/bin/env python3
"""
GitHub Uploader for NCB Data Processing Results
Uploads processed Excel files to a GitHub repository.
"""

import os
import sys
from github import Github
from datetime import datetime
import logging
import argparse
from pathlib import Path

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class GitHubUploader:
    """Handles uploading files to GitHub repository."""
    
    def __init__(self, token, repo_name, branch="main"):
        self.github = Github(token)
        self.repo_name = repo_name
        self.branch = branch
        self.repo = None
        
    def connect_to_repo(self):
        """Connect to the specified GitHub repository."""
        try:
            self.repo = self.github.get_repo(self.repo_name)
            logger.info(f"Connected to repository: {self.repo_name}")
            return True
        except Exception as e:
            logger.error(f"Failed to connect to repository: {e}")
            return False
    
    def upload_file(self, file_path, commit_message=None):
        """Upload a single file to GitHub."""
        if not self.repo:
            if not self.connect_to_repo():
                return False
        
        try:
            # Read file content
            with open(file_path, 'rb') as file:
                content = file.read()
            
            # Create commit message
            if not commit_message:
                commit_message = f"Upload {os.path.basename(file_path)} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            
            # Upload file
            file_name = os.path.basename(file_path)
            self.repo.create_file(
                path=f"data/{file_name}",
                message=commit_message,
                content=content,
                branch=self.branch
            )
            
            logger.info(f"Successfully uploaded: {file_name}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to upload {file_path}: {e}")
            return False
    
    def upload_directory(self, directory_path, commit_message=None):
        """Upload all files from a directory to GitHub."""
        if not os.path.exists(directory_path):
            logger.error(f"Directory not found: {directory_path}")
            return False
        
        success_count = 0
        total_count = 0
        
        for file_path in Path(directory_path).rglob('*'):
            if file_path.is_file():
                total_count += 1
                if self.upload_file(str(file_path), commit_message):
                    success_count += 1
        
        logger.info(f"Uploaded {success_count}/{total_count} files successfully")
        return success_count == total_count

def main():
    """Main function for GitHub uploader."""
    parser = argparse.ArgumentParser(description='Upload NCB data files to GitHub')
    parser.add_argument('--token', required=True, help='GitHub personal access token')
    parser.add_argument('--repo', required=True, help='GitHub repository name (owner/repo)')
    parser.add_argument('--path', required=True, help='Path to file or directory to upload')
    parser.add_argument('--branch', default='main', help='GitHub branch (default: main)')
    parser.add_argument('--message', help='Commit message')
    
    args = parser.parse_args()
    
    # Initialize uploader
    uploader = GitHubUploader(args.token, args.repo, args.branch)
    
    # Upload file or directory
    if os.path.isfile(args.path):
        success = uploader.upload_file(args.path, args.message)
    elif os.path.isdir(args.path):
        success = uploader.upload_directory(args.path, args.message)
    else:
        logger.error(f"Path not found: {args.path}")
        sys.exit(1)
    
    if success:
        logger.info("Upload completed successfully!")
        sys.exit(0)
    else:
        logger.error("Upload failed!")
        sys.exit(1)

if __name__ == "__main__":
    main()
