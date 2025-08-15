#!/usr/bin/env python3
"""
Debug script to check printer names and status
"""

import win32print
import win32evtlog
import win32evtlogutil

def list_all_printers():
    """List all printers with detailed information."""
    print("=== ALL AVAILABLE PRINTERS ===")
    try:
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        for i, printer in enumerate(printers):
            print(f"{i+1}. Name: {printer[2]}")
            print(f"   Server: {printer[1]}")
            print(f"   Share Name: {printer[3]}")
            print(f"   Port Name: {printer[4]}")
            print(f"   Driver Name: {printer[5]}")
            print(f"   Comment: {printer[6]}")
            print(f"   Location: {printer[7]}")
            print(f"   Print Processor: {printer[8]}")
            print(f"   Data Type: {printer[9]}")
            print(f"   Parameters: {printer[10]}")
            print()
    except Exception as e:
        print(f"Error: {e}")

def check_printer_status(printer_name):
    """Check detailed status of a specific printer."""
    print(f"=== DETAILED STATUS FOR: {printer_name} ===")
    try:
        handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(handle, 2)
        win32print.ClosePrinter(handle)
        
        print(f"Status: {printer_info['Status']}")
        print(f"Attributes: {printer_info['Attributes']}")
        print(f"Priority: {printer_info['Priority']}")
        print(f"Default Priority: {printer_info['DefaultPriority']}")
        print(f"Start Time: {printer_info['StartTime']}")
        print(f"Until Time: {printer_info['UntilTime']}")
        print(f"Jobs: {printer_info['Jobs']}")
        print(f"Average PPM: {printer_info['AveragePPM']}")
        
        # Status codes
        status_codes = {
            0x00000000: "PRINTER_STATUS_READY",
            0x00000001: "PRINTER_STATUS_PAUSED", 
            0x00000002: "PRINTER_STATUS_ERROR",
            0x00000004: "PRINTER_STATUS_PENDING_DELETION",
            0x00000008: "PRINTER_STATUS_PAPER_JAM",
            0x00000010: "PRINTER_STATUS_PAPER_OUT",
            0x00000020: "PRINTER_STATUS_MANUAL_FEED",
            0x00000040: "PRINTER_STATUS_PAPER_PROBLEM",
            0x00000080: "PRINTER_STATUS_OFFLINE",
            0x00000100: "PRINTER_STATUS_IO_ACTIVE",
            0x00000200: "PRINTER_STATUS_BUSY",
            0x00000400: "PRINTER_STATUS_PRINTING",
            0x00000800: "PRINTER_STATUS_OUTPUT_BIN_FULL",
            0x00001000: "PRINTER_STATUS_NOT_AVAILABLE",
            0x00002000: "PRINTER_STATUS_WAITING",
            0x00004000: "PRINTER_STATUS_PROCESSING",
            0x00008000: "PRINTER_STATUS_INITIALIZING",
            0x00010000: "PRINTER_STATUS_WARMING_UP",
            0x00020000: "PRINTER_STATUS_TONER_LOW",
            0x00040000: "PRINTER_STATUS_NO_TONER",
            0x00080000: "PRINTER_STATUS_PAGE_PUNT",
            0x00100000: "PRINTER_STATUS_USER_INTERVENTION",
            0x00200000: "PRINTER_STATUS_OUT_OF_MEMORY",
            0x00400000: "PRINTER_STATUS_DOOR_OPEN",
            0x00800000: "PRINTER_STATUS_SERVER_UNKNOWN",
            0x01000000: "PRINTER_STATUS_POWER_SAVE"
        }
        
        print(f"Status Description: {status_codes.get(printer_info['Status'], 'UNKNOWN')}")
        
    except Exception as e:
        print(f"Error checking printer {printer_name}: {e}")

def check_script_printers():
    """Check the printers that the script is looking for."""
    print("=== SCRIPT PRINTER MAPPING ===")
    script_printers = {
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
    
    try:
        system_printers = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        
        for script_name, printer_name in script_printers.items():
            if printer_name in system_printers:
                print(f"✓ {script_name} -> {printer_name} (FOUND)")
                check_printer_status(printer_name)
            else:
                print(f"✗ {script_name} -> {printer_name} (NOT FOUND)")
            print()
            
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    list_all_printers()
    print("\n" + "="*50 + "\n")
    check_script_printers()
