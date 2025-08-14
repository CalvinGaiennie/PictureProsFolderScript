# $MasterFolder = "/Users/calvingaiennie/Desktop/TestMasterFolder"
#Master folder to drop files to
$MasterFolder = "C:\Users\colem\Desktop\MasterPrintFolder"

# Keep track of files already handled to avoid duplicate processing
$ProcessedFiles = @{}

#Printer Pairs (fill in the actual printer names)
$PrinterPairs = @(
    @{ Photo = "PhotoPool1"; Label = "LabelPool1" },
    @{ Photo = "PhotoPool2"; Label = "LabelPool2" },
    @{ Photo = "PhotoPool3"; Label = "LabelPool3" },
    @{ Photo = "PhotoPool4"; Label = "LabelPool4" },
    @{ Photo = "PhotoPool5"; Label = "LabelPool5" },
    @{ Photo = "PhotoPool6"; Label = "LabelPool6" },
    @{ Photo = "PrinterPool7"; Label = "LabelPool7" },
    @{ Photo = "PrinterPool8"; Label = "LabelPool8" },
    @{ Photo = "PrinterPool9"; Label = "LabelPool9" },
    @{ Photo = "PrinterPool10"; Label = "LabelPool10" },
    @{ Photo = "PrinterPool11"; Label = "LabelPool11" },
    @{ Photo = "PrinterPool12"; Label = "LabelPool12" }
)

# Map hot folder names to actual printer names (fill these in)
$PrinterFolderMap = @{
    "PhotoPool1" = "P1"
    "LabelPool1" = "LP-1"
    "PhotoPool2" = "P2"
    "LabelPool2" = "LP-2"
    "PhotoPool3" = "P3"
    "LabelPool3" = "LP-3"
    "PhotoPool4" = "P4"
    "LabelPool4" = "LP-4"
    "PhotoPool5" = "P5"
    "LabelPool5" = "LP-5"
    "PhotoPool6" = "P6"
    "LabelPool6" = "LP-6"
    "PhotoPool7" = "P7"
    "LabelPool7" = "LP-7"
    "PhotoPool8" = "P8"
    "LabelPool8" = "LP-8"
    "PhotoPool9" = "P9"
    "LabelPool9" = "LP-9"
    "PhotoPool10" = "P10"
    "LabelPool10" = "LP-10"
    "PhotoPool11" = "P11"
    "LabelPool11" = "LP-1"
    "PhotoPool12" = "P12"
    "LabelPool12" = "LP-12"
}

# Track printer status internally (Free/Busy)
$PrinterStatus = @{}
foreach ($pair in $PrinterPairs) {
    $PrinterStatus[$pair.Photo] = "Free"
    $PrinterStatus[$pair.Label] = "Free"
}

# Track last printed document per printer (from Operational log)
$LastPrintedDocument = @{}

# Watch the master folder for new files
$Watcher = New-Object System.IO.FileSystemWatcher
$Watcher.Path = $MasterFolder
$Watcher.Filter = "*.*"          #Watch all files
$Watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite'
$Watcher.IncludeSubdirectories = $false
$Watcher.EnableRaisingEvents = $true

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

# Action when a new file appears
$Action = {
    $Path = $Event.SourceEventArgs.FullPath
    $Name = $Event.SourceEventArgs.Name

    if ($Name -eq ".DS_Store") { return }
    Start-Sleep -Milliseconds 200
    if ($ProcessedFiles.ContainsKey($Name)) { return }

    # Determine file type and numeric ID
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

    Write-Host "Detected $FileType with ID $FileID"

    # Find matching files
    $PhotoFile = Get-ChildItem $MasterFolder -Filter "photo$FileID*" | Where-Object { -not $ProcessedFiles.ContainsKey($_.Name) }
    $LabelFile = Get-ChildItem $MasterFolder -Filter "label$FileID*" | Where-Object { -not $ProcessedFiles.ContainsKey($_.Name) }

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

    # Move files
    try {
        Move-Item $PhotoFile.FullName -Destination "C:\Users\colem\Desktop\$($FreePair.Photo)\"
        Move-Item $LabelFile.FullName -Destination "C:\Users\colem\Desktop\$($FreePair.Label)\"
        Write-Host "Moved photo and label ID $FileID to printer folders"

        $ProcessedFiles[$PhotoFile.Name] = $true
        $ProcessedFiles[$LabelFile.Name] = $true

        # Update last printed document
        $LastPrintedDocument[$PrinterFolderMap[$FreePair.Photo]] = $PhotoFile.Name
        $LastPrintedDocument[$PrinterFolderMap[$FreePair.Label]] = $LabelFile.Name
    } catch {
        Write-Host "Error moving files: $_"
        $PrinterStatus[$FreePair.Photo] = "Free"
        $PrinterStatus[$FreePair.Label] = "Free"
    }

    # Mark printers as Free internally
    $PrinterStatus[$FreePair.Photo] = "Free"
    $PrinterStatus[$FreePair.Label] = "Free"
}

Register-ObjectEvent $Watcher Created -Action $Action
Write-Host "Watching $MasterFolder for new files. Press Ctrl+C to stop."

while ($true) { Start-Sleep -Seconds 1 }



# Currently it is expecting a file to contain a numerical id, no other numbers, and it to either contain the word phot or label. It is expecting the file name to begin with the label/photo word, then the number, then whatever else