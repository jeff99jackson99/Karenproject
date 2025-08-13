#!/usr/bin/env python3
"""
Web App Launcher for NCB Data Processor
Starts the Streamlit web interface.
"""

import subprocess
import sys
import os

def start_web_app():
    """Start the Streamlit web application."""
    print("🚀 Starting NCB Data Processor Web Interface...")
    print("=" * 50)
    
    # Check if streamlit is installed
    try:
        import streamlit
        print("✅ Streamlit is installed")
    except ImportError:
        print("❌ Streamlit not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "streamlit"])
        print("✅ Streamlit installed successfully")
    
    # Get the current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    web_app_path = os.path.join(current_dir, "web_app.py")
    
    print(f"📁 Web app location: {web_app_path}")
    print("🌐 Starting web server...")
    print("\n💡 The web interface will open in your browser automatically")
    print("💡 If it doesn't open, go to: http://localhost:8501")
    print("\n🛑 To stop the web app, press Ctrl+C in this terminal")
    print("=" * 50)
    
    # Start the Streamlit app
    try:
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", web_app_path,
            "--server.port", "8501",
            "--server.address", "localhost"
        ])
    except KeyboardInterrupt:
        print("\n🛑 Web app stopped by user")
    except Exception as e:
        print(f"❌ Error starting web app: {e}")

if __name__ == "__main__":
    start_web_app()
