#!/usr/bin/env python3
"""
Dependency checker for GPT-PowerPoint Generator
This script checks that all required dependencies are installed and configured correctly.
"""

import sys
import importlib
import os
from dotenv import load_dotenv

def check_module(module_name):
    """Check if a module is installed."""
    try:
        importlib.import_module(module_name)
        return True
    except ImportError:
        return False

def main():
    # Required modules
    required_modules = [
        "flask", "openai", "python-dotenv", "python-pptx"
    ]
    
    # Check Python version
    print(f"Python version: {sys.version}")
    if sys.version_info < (3, 7):
        print("❌ Python 3.7 or higher is required")
    else:
        print("✅ Python version OK")
    
    # Check required modules
    print("\nChecking required modules:")
    all_modules_installed = True
    for module in required_modules:
        module_name = module
        if module == "python-dotenv":
            module_name = "dotenv"
        elif module == "python-pptx":
            module_name = "pptx"
            
        if check_module(module_name):
            print(f"✅ {module} installed")
        else:
            print(f"❌ {module} not installed")
            all_modules_installed = False
    
    # Check environment variables
    print("\nChecking environment configuration:")
    load_dotenv()
    if os.getenv("OPENAI_API_KEY"):
        print("✅ OPENAI_API_KEY found")
    else:
        print("❌ OPENAI_API_KEY not found. Please create a .env file with your API key.")
    
    # Check for output directory
    print("\nChecking file structure:")
    if os.path.exists("generated"):
        print("✅ 'generated' directory exists")
    else:
        print("ℹ️ 'generated' directory will be created when the app runs")
    
    # Check for web templates
    if os.path.exists("templates/index.html"):
        print("✅ Web templates found")
    else:
        print("❌ Web templates not found")
    
    # Final status
    print("\nResults:")
    if all_modules_installed and os.getenv("OPENAI_API_KEY"):
        print("✅ All dependencies are installed and configured correctly. You're good to go!")
    else:
        print("❌ Some dependencies are missing. Please install them using:")
        print("   pip install -r requirements.txt")
        print("   Also make sure you have a .env file with your OPENAI_API_KEY.")

if __name__ == "__main__":
    main() 