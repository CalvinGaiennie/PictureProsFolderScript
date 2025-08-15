#!/usr/bin/env python3
"""
Test script for Picture Pros Folder Script
Creates test files and runs the main script for testing.
"""

import os
import time
import subprocess
import threading
from pathlib import Path

# Configuration - update this path
MASTER_FOLDER = r"C:\Users\colem\Desktop\MasterPrintFolder"

def create_test_files():
    """Create test files in the master folder."""
    master_path = Path(MASTER_FOLDER)
    master_path.mkdir(parents=True, exist_ok=True)
    
    # Create test files
    test_files = [
        "photo800.jpg",
        "label800.pdf", 
        "photo123.png",
        "label123.txt"
    ]
    
    print("Creating test files...")
    for filename in test_files:
        file_path = master_path / filename
        with open(file_path, 'w') as f:
            f.write(f"Test content for {filename}")
        print(f"Created: {filename}")
        time.sleep(1)  # Wait 1 second between files

def run_main_script():
    """Run the main picture pros script."""
    print("Starting main script...")
    subprocess.run(["python", "picture_pros_folder_script.py"])

def main():
    print("=== Picture Pros Script Test ===")
    print(f"Master folder: {MASTER_FOLDER}")
    print()
    
    # Create test files in a separate thread
    test_thread = threading.Thread(target=create_test_files)
    test_thread.daemon = True
    test_thread.start()
    
    # Wait a moment, then start the main script
    time.sleep(2)
    
    try:
        run_main_script()
    except KeyboardInterrupt:
        print("\nTest stopped by user.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
