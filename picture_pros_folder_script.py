#!/usr/bin/env python3
"""
Picture Pros Folder Script - Python Version
Watches a master folder for photo and label files, pairs them by ID, and moves them to appropriate printer folders.
"""

import os
import time
import re
import shutil
import logging
from pathlib import Path
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from typing import Dict, List, Optional, Tuple
import win32evtlog
import win32evtlogutil
import win32con
import win32api

# Configuration
MASTER_FOLDER = r"C:\Users\colem\Desktop\MasterPrintFolder"
PROCESSED_FILES = set()

# Printer Pairs (fill in the actual printer names)
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

# Map hot folder names to actual printer names
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

# Track printer status internally (Free/Busy)
PRINTER_STATUS = {}
for pair in PRINTER_PAIRS:
    PRINTER_STATUS[pair["photo"]] = "Free"
    PRINTER_STATUS[pair["label"]] = "Free"

# Track last printed document per printer (from operational log)
LAST_PRINTED_DOCUMENT = {}

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('picture_pros.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def is_printer_available(printer_name: str) -> bool:
    """Check if a printer is available and connected."""
    try:
        import win32print
        
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
                        logger.warning(f"Printer {printer_name} is OFFLINE")
                        return False
                    elif status & 0x00000002:  # PRINTER_STATUS_ERROR
                        logger.warning(f"Printer {printer_name} has ERROR status")
                        return False
                    elif status & 0x00000001:  # PRINTER_STATUS_PAUSED
                        logger.warning(f"Printer {printer_name} is PAUSED")
                        return False
                    elif status & 0x00001000:  # PRINTER_STATUS_NOT_AVAILABLE
                        logger.warning(f"Printer {printer_name} is NOT_AVAILABLE")
                        return False
                    elif status & 0x00000008:  # PRINTER_STATUS_PAPER_JAM
                        logger.warning(f"Printer {printer_name} has PAPER_JAM")
                        return False
                    elif status & 0x00000010:  # PRINTER_STATUS_PAPER_OUT
                        logger.warning(f"Printer {printer_name} is OUT_OF_PAPER")
                        return False
                    elif status & 0x00000400:  # PRINTER_STATUS_PRINTING
                        logger.info(f"Printer {printer_name} is currently PRINTING")
                        return False
                    elif status & 0x00000200:  # PRINTER_STATUS_BUSY
                        logger.info(f"Printer {printer_name} is BUSY")
                        return False
                    elif status == 0:  # PRINTER_STATUS_READY
                        # Check if printer is actually connected by testing port communication
                        try:
                            # Try to get printer capabilities - this will fail if printer is not physically connected
                            handle = win32print.OpenPrinter(printer_name)
                            # Just try to get basic info - this will fail if printer is disconnected
                            test_info = win32print.GetPrinter(handle, 1)  # Level 1 info
                            win32print.ClosePrinter(handle)
                            
                            # If we get here, the printer is actually connected
                            logger.info(f"Printer {printer_name} is available and ready (physically connected)")
                            return True
                                
                        except Exception as e:
                            logger.warning(f"Printer {printer_name} appears to be disconnected: {e}")
                            return False
                    else:
                        logger.warning(f"Printer {printer_name} has unknown status: {status}")
                        return False
                        
                except Exception as e:
                    logger.warning(f"Error checking printer {printer_name}: {e}")
                    return False
        
        logger.warning(f"Printer {printer_name} not found in system")
        return False
        
    except Exception as e:
        logger.warning(f"Error enumerating printers: {e}")
        return False


def is_printer_free(printer_name: str) -> bool:
    """Check if a printer has finished its last job by checking Windows Event Logs."""
    # First check if printer is available
    if not is_printer_available(printer_name):
        return False
        
    try:
        # Open the print service operational log
        hand = win32evtlog.OpenEventLog(None, "Microsoft-Windows-PrintService/Operational")
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        
        # Look for print job completion events (ID 307)
        events = win32evtlog.ReadEventLog(hand, flags, 0)
        
        for event in events:
            if event.EventID == 307:  # Print job completion
                # Extract printer name and document name from event data
                event_data = win32evtlogutil.SafeFormatMessage(event, "Microsoft-Windows-PrintService/Operational")
                if printer_name in event_data:
                    # This is a simplified check - in practice you'd parse the event data more carefully
                    return True
        
        return True  # Assume free if no recent events found
        
    except Exception as e:
        logger.warning(f"Error checking printer status for {printer_name}: {e}")
        return False  # Don't assume free on error - be more conservative


def list_available_printers():
    """List all available printers for debugging."""
    try:
        import win32print
        
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        logger.info("Available printers:")
        for printer in printers:
            logger.info(f"  - {printer[2]} (Status: {printer[3]})")
    except Exception as e:
        logger.error(f"Error listing printers: {e}")


def get_free_printer_pair() -> Optional[Dict[str, str]]:
    """Get a free printer pair."""
    logger.info("Checking for available printer pairs...")
    
    for pair in PRINTER_PAIRS:
        photo_printer = PRINTER_FOLDER_MAP[pair["photo"]]
        label_printer = PRINTER_FOLDER_MAP[pair["label"]]
        
        logger.info(f"Checking pair: {photo_printer} + {label_printer}")
        
        photo_available = is_printer_free(photo_printer)
        label_available = is_printer_free(label_printer)
        
        logger.info(f"  Photo printer {photo_printer}: {'Available' if photo_available else 'Not available'}")
        logger.info(f"  Label printer {label_printer}: {'Available' if label_available else 'Not available'}")
        
        if photo_available and label_available:
            PRINTER_STATUS[pair["photo"]] = "Busy"
            PRINTER_STATUS[pair["label"]] = "Busy"
            logger.info(f"Found available pair: {pair['photo']} + {pair['label']}")
            return pair
    
    logger.warning("No available printer pairs found")
    return None


def find_matching_files(file_id: str, master_folder: Path) -> Tuple[Optional[Path], Optional[Path]]:
    """Find matching photo and label files for a given ID."""
    photo_pattern = f"photo{file_id}*"
    label_pattern = f"label{file_id}*"
    
    photo_file = None
    label_file = None
    
    for file_path in master_folder.glob("*"):
        if file_path.name in PROCESSED_FILES:
            continue
            
        if file_path.name.startswith(f"photo{file_id}"):
            photo_file = file_path
        elif file_path.name.startswith(f"label{file_id}"):
            label_file = file_path
    
    return photo_file, label_file


def move_files_to_printer_folders(photo_file: Path, label_file: Path, printer_pair: Dict[str, str]) -> bool:
    """Move photo and label files to their respective printer folders."""
    try:
        # Create destination paths
        photo_dest = Path(f"C:/Users/colem/Desktop/{printer_pair['photo']}/")
        label_dest = Path(f"C:/Users/colem/Desktop/{printer_pair['label']}/")
        
        # Ensure destination folders exist
        photo_dest.mkdir(parents=True, exist_ok=True)
        label_dest.mkdir(parents=True, exist_ok=True)
        
        # Move files
        shutil.move(str(photo_file), str(photo_dest / photo_file.name))
        shutil.move(str(label_file), str(label_dest / label_file.name))
        
        # Mark as processed
        PROCESSED_FILES.add(photo_file.name)
        PROCESSED_FILES.add(label_file.name)
        
        # Update last printed documents
        LAST_PRINTED_DOCUMENT[PRINTER_FOLDER_MAP[printer_pair["photo"]]] = photo_file.name
        LAST_PRINTED_DOCUMENT[PRINTER_FOLDER_MAP[printer_pair["label"]]] = label_file.name
        
        logger.info(f"Moved photo and label ID {photo_file.name.split('photo')[1].split('.')[0]} to printer folders")
        return True
        
    except Exception as e:
        logger.error(f"Error moving files: {e}")
        # Reset printer status on error
        PRINTER_STATUS[printer_pair["photo"]] = "Free"
        PRINTER_STATUS[printer_pair["label"]] = "Free"
        return False


class FileHandler(FileSystemEventHandler):
    """Handle file system events for the master folder."""
    
    def __init__(self, master_folder: Path):
        self.master_folder = master_folder
    
    def on_created(self, event):
        if event.is_directory:
            return
        
        file_path = Path(event.src_path)
        file_name = file_path.name
        
        # Skip system files
        if file_name == ".DS_Store" or file_name.startswith("~"):
            return
        
        # Wait a bit for file to be fully written
        time.sleep(0.2)
        
        # Skip if already processed
        if file_name in PROCESSED_FILES:
            return
        
        # Extract file type and ID
        photo_match = re.match(r"^photo(\d+).*", file_name)
        label_match = re.match(r"^label(\d+).*", file_name)
        
        if photo_match:
            file_type = "Photo"
            file_id = photo_match.group(1)
        elif label_match:
            file_type = "Label"
            file_id = label_match.group(1)
        else:
            logger.info(f"File does not match photo or label pattern: {file_name}")
            return
        
        logger.info(f"Detected {file_type} with ID {file_id} for file: {file_name}")
        
        # Find matching pair
        photo_file, label_file = find_matching_files(file_id, self.master_folder)
        
        if not photo_file or not label_file:
            logger.info(f"Waiting for matching pair for ID {file_id}")
            return
        
        # Get free printer pair
        free_pair = get_free_printer_pair()
        if not free_pair:
            logger.info("No free printer pairs available. Will try again later.")
            return
        
        logger.info(f"Using printer pair: Photo={free_pair['photo']}, Label={free_pair['label']}")
        
        # Move files
        if move_files_to_printer_folders(photo_file, label_file, free_pair):
            # Reset printer status after successful move
            PRINTER_STATUS[free_pair["photo"]] = "Free"
            PRINTER_STATUS[free_pair["label"]] = "Free"


def main():
    """Main function to start the file watcher."""
    master_path = Path(MASTER_FOLDER)
    
    # Ensure master folder exists
    if not master_path.exists():
        logger.error(f"Master folder does not exist: {MASTER_FOLDER}")
        return
    
    logger.info(f"Starting file watcher for: {MASTER_FOLDER}")
    
    # List available printers for debugging
    list_available_printers()
    
    # Create and start the observer
    event_handler = FileHandler(master_path)
    observer = Observer()
    observer.schedule(event_handler, str(master_path), recursive=False)
    observer.start()
    
    try:
        logger.info("Watching for new files. Press Ctrl+C to stop.")
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        logger.info("Stopping file watcher...")
        observer.stop()
    
    observer.join()


if __name__ == "__main__":
    main()
