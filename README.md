# Picture Pros Folder Script

A Python script that automatically watches a master folder for photo and label files, pairs them by ID, and moves them to appropriate printer folders.

## Features

- **File Watching**: Monitors a master folder for new photo and label files
- **Automatic Pairing**: Matches photo and label files by their ID number
- **Printer Management**: Checks printer availability via Windows Event Logs
- **Duplicate Prevention**: Tracks processed files to avoid duplicates
- **Logging**: Comprehensive logging to both console and file

## Prerequisites

- Windows 10/11
- Python 3.8 or higher
- Administrative privileges (for service installation)

## Installation

### 1. Install Python Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure the Script

Edit `picture_pros_folder_script.py` and update these settings:

```python
# Change this to your master folder path
MASTER_FOLDER = r"C:\Users\colem\Desktop\MasterPrintFolder"

# Update printer folder paths if needed
photo_dest = Path(f"C:/Users/colem/Desktop/{printer_pair['photo']}/")
label_dest = Path(f"C:/Users/colem/Desktop/{printer_pair['label']}/")
```

### 3. Choose Your Installation Method

#### Option A: Windows Service (Recommended)

Run as administrator:

```bash
install_service.bat
```

This creates a Windows service that:

- Starts automatically when Windows boots
- Restarts automatically if it crashes
- Runs in the background

**Service Commands:**

- Start: `net start PictureProsFolderScript`
- Stop: `net stop PictureProsFolderScript`
- Status: `sc query PictureProsFolderScript`

#### Option B: Scheduled Task

Run as administrator:

```bash
create_scheduled_task.bat
```

This creates a scheduled task that runs at startup.

**Task Commands:**

- Start: `schtasks /run /tn "PictureProsFolderScript"`
- Stop: `schtasks /end /tn "PictureProsFolderScript"`

#### Option C: Manual Run

For testing or temporary use:

```bash
python picture_pros_folder_script.py
```

## File Naming Convention

The script expects files to follow this naming pattern:

- Photos: `photo{ID}.*` (e.g., `photo800.jpg`, `photo123.png`)
- Labels: `label{ID}.*` (e.g., `label800.pdf`, `label123.txt`)

Files with the same ID will be paired and moved together.

## Logging

The script creates a log file `picture_pros.log` in the same directory with detailed information about:

- File detection and pairing
- Printer status checks
- File movements
- Errors and warnings

## Troubleshooting

### Common Issues

1. **Permission Denied**: Run installation scripts as Administrator
2. **Python Not Found**: Ensure Python is installed and in PATH
3. **Dependencies Missing**: Run `pip install -r requirements.txt`
4. **Folder Not Found**: Check that `MASTER_FOLDER` path exists

### Check Service Status

```bash
sc query PictureProsFolderScript
```

### View Recent Logs

```bash
type picture_pros.log
```

### Restart Service

```bash
net stop PictureProsFolderScript
net start PictureProsFolderScript
```

## Configuration

### Printer Pairs

Edit the `PRINTER_PAIRS` list in the script to match your printer setup:

```python
PRINTER_PAIRS = [
    {"photo": "PhotoPool1", "label": "LabelPool1"},
    {"photo": "PhotoPool2", "label": "LabelPool2"},
    # Add more pairs as needed
]
```

### Printer Names

Update `PRINTER_FOLDER_MAP` to match your actual printer names:

```python
PRINTER_FOLDER_MAP = {
    "PhotoPool1": "P1", "LabelPool1": "LP-1",
    "PhotoPool2": "P2", "LabelPool2": "LP-2",
    # Update with your actual printer names
}
```

## Support

For issues or questions:

1. Check the log file `picture_pros.log`
2. Verify all paths and permissions
3. Ensure Python dependencies are installed
4. Run the script manually first to test configuration
