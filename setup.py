#!/usr/bin/env python3
"""
Setup script for Excel Sheets Agent
This script helps set up the virtual environment and install dependencies
"""

import subprocess
import sys
import os
from pathlib import Path

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed successfully")
        if result.stdout:
            print(result.stdout)
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed")
        print(f"Error: {e.stderr}")
        return False

def main():
    """Main setup function"""
    print("🚀 Setting up Excel Sheets Agent - Phase 1")
    print("=" * 50)
    
    # Check if we're on Windows
    is_windows = os.name == 'nt'
    
    # Create virtual environment
    if not os.path.exists('.venv'):
        if not run_command('python -m venv .venv', 'Creating virtual environment'):
            return False
    else:
        print("✅ Virtual environment already exists")
    
    # Activate virtual environment and install dependencies
    if is_windows:
        activate_cmd = '.venv\\Scripts\\activate'
        pip_cmd = '.venv\\Scripts\\pip'
    else:
        activate_cmd = 'source .venv/bin/activate'
        pip_cmd = '.venv/bin/pip'
    
    # Install dependencies
    if not run_command(f'{pip_cmd} install -r requirements.txt', 'Installing dependencies'):
        return False
    
    print("\n🎉 Setup completed successfully!")
    print("\nNext steps:")
    print("1. Activate the virtual environment:")
    if is_windows:
        print("   .venv\\Scripts\\activate")
    else:
        print("   source .venv/bin/activate")
    print("2. Run the application:")
    print("   streamlit run app.py")
    
    return True

if __name__ == "__main__":
    main()
