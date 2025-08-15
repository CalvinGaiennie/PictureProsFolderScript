#!/usr/bin/env python3
"""
Test script to check which printer pairs are actually available
"""

import win32print
import logging

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Printer configuration from main script
PRINTER_PAIRS = [
    {"photo": "PhotoPool1", "label": "LabelPool1"},
    {"photo": "PhotoPool2", "label": "LabelPool2"},
    {"photo": "PhotoPool3", "label": "LabelPool3"},
    {"photo": "PhotoPool4", "label": "LabelPool4"},
    {"photo": "PhotoPool5", "label": "LabelPool5"},
    {"photo": "PhotoPool6", "label": "LabelPool6"},
    {"photo": "PhotoPool7", "label": "LabelPool7"},
    {"photo": "PhotoPool8", "label": "LabelPool8"},
    {"photo": "PhotoPool9", "label": "LabelPool9"},
    {"photo": "PhotoPool10", "label": "LabelPool10"},
    {"photo": "PhotoPool11", "label": "LabelPool11"},
    {"photo": "PhotoPool12", "label": "LabelPool12"}
]

PRINTER_FOLDER_MAP = {
    "PhotoPool1": "P1", "LabelPool1": "LP-1",
    "PhotoPool2": "P2", "LabelPool2": "LP-2",
    "PhotoPool3": "P3", "LabelPool3": "LP-3",
    "PhotoPool4": "P4", "LabelPool4": "LP-4",
    "PhotoPool5": "P5", "LabelPool5": "LP-5",
    "PhotoPool6": "P6", "LabelPool6": "LP-6",
    "PhotoPool7": "P7", "LabelPool7": "LP-7",
    "PhotoPool8": "P8", "LabelPool8": "LP-8",
    "PhotoPool9": "P9", "LabelPool9": "LP-9",
    "PhotoPool10": "P10", "LabelPool10": "LP-10",
    "PhotoPool11": "P11", "LabelPool11": "LP-11",
    "PhotoPool12": "P12", "LabelPool12": "LP-12"
}

def is_printer_available(printer_name: str) -> bool:
    """Check if a printer is available and connected."""
    try:
        # Get all printers
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        
        # Look for our printer
        for printer in printers:
            if printer[2] == printer_name:  # printer[2] is the printer name
                # Check if printer is ready
                try:
                    handle = win32print.OpenPrinter(printer_name)
                    printer_info = win32print.GetPrinter(handle, 2)  # Level 2 info
                    win32print.ClosePrinter(handle)
                    
                    # Check printer status
                    status = printer_info['Status']
                    
                    # Check for specific error conditions
                    if status & 0x00000080:  # PRINTER_STATUS_OFFLINE
                        print(f"  ✗ {printer_name}: OFFLINE")
                        return False
                    elif status & 0x00000002:  # PRINTER_STATUS_ERROR
                        print(f"  ✗ {printer_name}: ERROR")
                        return False
                    elif status & 0x00000001:  # PRINTER_STATUS_PAUSED
                        print(f"  ✗ {printer_name}: PAUSED")
                        return False
                    elif status & 0x00001000:  # PRINTER_STATUS_NOT_AVAILABLE
                        print(f"  ✗ {printer_name}: NOT_AVAILABLE")
                        return False
                    elif status & 0x00000008:  # PRINTER_STATUS_PAPER_JAM
                        print(f"  ✗ {printer_name}: PAPER_JAM")
                        return False
                    elif status & 0x00000010:  # PRINTER_STATUS_PAPER_OUT
                        print(f"  ✗ {printer_name}: OUT_OF_PAPER")
                        return False
                    elif status & 0x00000400:  # PRINTER_STATUS_PRINTING
                        print(f"  ✗ {printer_name}: PRINTING")
                        return False
                    elif status & 0x00000200:  # PRINTER_STATUS_BUSY
                        print(f"  ✗ {printer_name}: BUSY")
                        return False
                    elif status == 0:  # PRINTER_STATUS_READY
                        print(f"  ✓ {printer_name}: READY")
                        return True
                    else:
                        print(f"  ✗ {printer_name}: UNKNOWN STATUS ({status})")
                        return False
                        
                except Exception as e:
                    print(f"  ✗ {printer_name}: ERROR - {e}")
                    return False
        
        print(f"  ✗ {printer_name}: NOT_FOUND")
        return False
        
    except Exception as e:
        print(f"  ✗ {printer_name}: ERROR - {e}")
        return False

def test_all_printer_pairs():
    """Test all printer pairs and show which ones are available."""
    print("=== TESTING ALL PRINTER PAIRS ===")
    print()
    
    available_pairs = []
    
    for i, pair in enumerate(PRINTER_PAIRS, 1):
        photo_printer = PRINTER_FOLDER_MAP[pair["photo"]]
        label_printer = PRINTER_FOLDER_MAP[pair["label"]]
        
        print(f"Pair {i}: {pair['photo']} + {pair['label']}")
        print(f"  Checking: {photo_printer} + {label_printer}")
        
        photo_available = is_printer_available(photo_printer)
        label_available = is_printer_available(label_printer)
        
        if photo_available and label_available:
            print(f"  ✅ BOTH AVAILABLE - This pair can be used!")
            available_pairs.append(pair)
        else:
            print(f"  ❌ NOT AVAILABLE - Cannot use this pair")
        
        print()

    print("=== SUMMARY ===")
    print(f"Total pairs: {len(PRINTER_PAIRS)}")
    print(f"Available pairs: {len(available_pairs)}")
    
    if available_pairs:
        print("\nAvailable pairs:")
        for i, pair in enumerate(available_pairs, 1):
            photo_printer = PRINTER_FOLDER_MAP[pair["photo"]]
            label_printer = PRINTER_FOLDER_MAP[pair["label"]]
            print(f"  {i}. {pair['photo']} ({photo_printer}) + {pair['label']} ({label_printer})")
    else:
        print("\n❌ NO AVAILABLE PRINTER PAIRS FOUND!")
        print("Check your printer connections and status.")

if __name__ == "__main__":
    test_all_printer_pairs()
