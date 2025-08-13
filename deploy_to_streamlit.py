#!/usr/bin/env python3
"""
Deployment Helper for NCB Data Processor
Helps you deploy to Streamlit Cloud and other platforms.
"""

import os
import sys
import webbrowser
import subprocess

def deploy_to_streamlit():
    """Guide user through Streamlit Cloud deployment."""
    print("🚀 NCB Data Processor - Streamlit Cloud Deployment")
    print("=" * 50)
    
    print("\n📋 Prerequisites Check:")
    
    # Check if code is on GitHub
    print("✅ Your code is on GitHub at: https://github.com/jeff99jackson99/Karenproject")
    
    # Check if web_app.py exists
    if os.path.exists("web_app.py"):
        print("✅ Web application file exists")
    else:
        print("❌ Web application file missing")
        return False
    
    # Check if requirements.txt exists
    if os.path.exists("requirements.txt"):
        print("✅ Requirements file exists")
    else:
        print("❌ Requirements file missing")
        return False
    
    print("\n🌐 Deployment Steps:")
    print("1. Go to: https://share.streamlit.io/")
    print("2. Sign in with your GitHub account")
    print("3. Click 'New app'")
    print("4. Select repository: jeff99jackson99/Karenproject")
    print("5. Set main file path: web_app.py")
    print("6. Click 'Deploy!'")
    
    print("\n💡 Your app will be available at:")
    print("   https://your-app-name.streamlit.app")
    
    # Ask if user wants to open the deployment page
    choice = input("\n🌐 Open Streamlit Cloud deployment page? (y/n): ").strip().lower()
    if choice == 'y':
        webbrowser.open("https://share.streamlit.io/")
        print("✅ Opened Streamlit Cloud in your browser!")
    
    return True

def deploy_to_railway():
    """Guide user through Railway deployment."""
    print("\n🚂 Railway Deployment Alternative:")
    print("1. Go to: https://railway.app/")
    print("2. Sign in with GitHub")
    print("3. Create new project")
    print("4. Connect your repository")
    print("5. Deploy automatically")
    
    choice = input("\n�� Open Railway deployment page? (y/n): ").strip().lower()
    if choice == 'y':
        webbrowser.open("https://railway.app/")
        print("✅ Opened Railway in your browser!")

def main():
    """Main deployment function."""
    print("🌐 NCB Data Processor - Web Deployment Options")
    print("=" * 50)
    
    print("\nChoose your deployment option:")
    print("1. Streamlit Cloud (Recommended - FREE)")
    print("2. Railway (Alternative - FREE)")
    print("3. View deployment guide")
    print("4. Exit")
    
    choice = input("\nEnter your choice (1-4): ").strip()
    
    if choice == "1":
        deploy_to_streamlit()
    elif choice == "2":
        deploy_to_railway()
    elif choice == "3":
        print("\n📖 Opening deployment guide...")
        if os.path.exists("DEPLOYMENT_GUIDE.md"):
            subprocess.run(["open", "DEPLOYMENT_GUIDE.md"])
        else:
            print("Deployment guide: DEPLOYMENT_GUIDE.md")
    elif choice == "4":
        print("👋 Goodbye!")
        return
    else:
        print("❌ Invalid choice. Please try again.")
        main()

if __name__ == "__main__":
    main()
