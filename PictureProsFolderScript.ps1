# Master folder to drop files into
$MasterFolder = "C:\Users\colem\Desktop\MasterPrintFolder"

# Keep track of files already handled to avoid duplicate processing
$ProcessedFiles = @{}

# Printer Pairs (fill in the actual printer names)
$PrinterPairs = @(
    @{ Photo = "PhotoPool1"; Label = "LabelPool1" },
    @{ Photo = "PhotoPool2"; Label = "LabelPool2" },
    @{ Photo = "PhotoPool3"; Label = "LabelPool3" },
    @{ Photo = "PhotoPool4"; Label = "LabelPool4" },
    @{ Photo = "PhotoPool5"; Label = "LabelPool5" },
    @{ Photo = "PhotoPool6"; Label = "LabelPool6" },
    @{ Photo = "PhotoPool7"; Label = "LabelPool7" },
    @{ Photo = "PhotoPool8"; Label = "LabelPool8" },
    @{ Photo = "PhotoPool9"; Label = "LabelPool9" },
    @{ Photo = "PhotoPool10"; Label = "LabelPool10" },
    @{ Photo = "PhotoPool11"; Label = "LabelPool11" },
    @{ Photo = "PhotoPool12"; Label = "LabelPool12" }
)

# Map hot folder names to actual printer names
$PrinterFolderMap = @{
    "PhotoPool1" = "P1"; "LabelPool1" = "LP-1"
    "PhotoPool2" = "P2"; "LabelPool2" = "LP-2"
    "PhotoPool3" = "P3"; "LabelPool3" = "LP-3"
    "PhotoPool4" = "P4"; "LabelPool4" = "LP-4"
    "PhotoPool5" = "P5"; "LabelPool5" = "LP-5"
    "PhotoPool6" = "P6"; "LabelPool6" = "LP-6"
    "PhotoPool7" = "P7"; "LabelPool7" = "LP-7"
    "PhotoPool8" = "P8"; "LabelPool8" = "LP-8"
    "PhotoPool9" = "P9"; "LabelPool9" = "LP-9"
    "PhotoPool10" = "P10"; "LabelPool10" = "LP-10"
    "PhotoPool11" = "P11"; "LabelPool11" = "LP-11"
    "PhotoPool12" = "P12"; "LabelPool12" = "LP-12"
}

# Track printer status internally (Free/Busy)
$PrinterStatus = @{}
foreach ($pair in $PrinterPairs) {
    $PrinterStatus[$pair.Photo] = "Free"
    $PrinterStatus[$pair.Label] = "Free"
}

# Track last printed document per printer (from operational log)
$LastPrintedDocument = @{}

# Function to check if a printer has finished its last job
function Is-PrinterFree($PrinterName) {
    $lastEvent = Get-WinEvent -LogName "Microsoft-Windows-PrintService/Operational" |
        Where-Object { $_.Id -eq 307 -and $_.Properties[1].Value -eq $PrinterName } |
        Sort-Object TimeCreated -Descending |
        Select-Object -First 1

    if (-not $lastEvent) { return $true }

    $lastDoc = $lastEvent.Properties[0].Value
    if ($LastPrintedDocument[$PrinterName] -eq $lastDoc) {
        return $true
    } else {
        return $false
    }
}

# Function to get a free printer pair
function Get-FreePrinterPair {
    foreach ($pair in $PrinterPairs) {
        $photoPrinter = $PrinterFolderMap[$pair.Photo]
        $labelPrinter = $PrinterFolderMap[$pair.Label]

        if ((Is-PrinterFree $photoPrinter) -and (Is-PrinterFree $labelPrinter)) {
            $PrinterStatus[$pair.Photo] = "Busy"
            $PrinterStatus[$pair.Label] = "Busy"
            return $pair
        }
    }
    return $null
}

# Watch the master folder for new files
$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path = $MasterFolder
$Watcher.Filter = "*.*"
$Watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite'
$Watcher.IncludeSubdirectories = $false
$Watcher.EnableRaisingEvents = $true

# Define action for new files
$Action = {
    $Name = $Event.SourceEventArgs.Name
    $FullPath = $Event.SourceEventArgs.FullPath

    if ($Name -eq ".DS_Store") { return }

    Start-Sleep -Milliseconds 200

    if (ProcessedFiles.ContainsKey($Name)) { return }

    if ($Name -match "^photo(\d+).*") {
        $FileType = "Photo"
        $FileID = $Matches[1]
    } elseif ($Name -match "^label(\d+).*") {
        $FileType = "Label"
        $FileID = $Matches[1]
    } else {
        Write-Host "File does not match photo or label pattern: $Name"
        return
    }

    Write-Host "Detected $FileType with ID $FileID for file: $Name"

    $PhotoFile = Get-ChildItem $MasterFolder -Filter "photo$FileID*" | Where-Object { $_.Name -notin $ProcessedFiles.Keys }
    $LabelFile = Get-ChildItem $MasterFolder -Filter "label$FileID*" | Where-Object { $_.Name -notin $ProcessedFiles.Keys }

    if (-not $PhotoFile -or -not $LabelFile) {
        Write-Host "Waiting for matching pair for ID $FileID"
        return
    }

    $FreePair = Get-FreePrinterPair
    if ($FreePair -eq $null) {
        Write-Host "No free printer pairs available. Will try again later."
        return
    }

    Write-Host "Using printer pair: Photo=$($FreePair.Photo), Label=$($FreePair.Label)"

    try {
        Move-Item $PhotoFile.FullName -Destination "C:\Users\colem\Desktop\$($FreePair.Photo)\"
        Move-Item $LabelFile.FullName -Destination "C:\Users\colem\Desktop\$($FreePair.Label)\"
        Write-Host "Moved photo and label ID $FileID to printer folders"

        ProcessedFiles[$PhotoFile.Name] = $true
        ProcessedFiles[$LabelFile.Name] = $true

        $LastPrintedDocument[$PrinterFolderMap[$FreePair.Photo]] = $PhotoFile.Name
        $LastPrintedDocument[$PrinterFolderMap[$FreePair.Label]] = $LabelFile.Name
    } catch {
        Write-Host "Error moving files: $_"
        $PrinterStatus[$FreePair.Photo] = "Free"
        $PrinterStatus[$FreePair.Label] = "Free"
    }

    $PrinterStatus[$FreePair.Photo] = "Free"
    $PrinterStatus[$FreePair.Label] = "Free"
}

# Register the watcher event
Register-ObjectEvent $Watcher Created -Action $Action
Write-Host "Watching $MasterFolder for new files. Press Ctrl+C to stop."

# Keep script running
while ($true) { Start-Sleep -Seconds 1 }
