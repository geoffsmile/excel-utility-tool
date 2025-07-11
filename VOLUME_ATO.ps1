<#
.SYNOPSIS
    Excel Automation Script for FCC Inventory Processing

.DESCRIPTION
    This script automates several tasks related to FCC Inventory processing:
    1. Creates a daily template file based on the current date
    2. Cleans up existing files in the destination folder
    3. Processes CSV files from downloads folder, converting them to Excel format
    4. Opens the FCC inventory tool file upon completion

.AUTHOR
    Geoff Lu, FCC Reports Analyst

.DATE
    April 11, 2025
#>

# Minimize the console window
$host.UI.RawUI.WindowPosition = New-Object System.Management.Automation.Host.Coordinates(0, -1000)

# Define the source and destination folder paths
$csvFolder = "C:\Users\PH10035990\Downloads\"
$destFolder = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\04 - Raw Data Pull\"
$toolFilePath = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\TOOL - R1Access_Inv_Gen - GEOFF.xlsm"

# Define template file paths
$templateSourcePath = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\03 - Combined\Archive\TEMPLATE.xlsm"
$templateDestFolder = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\03 - Combined\"

# Add necessary assembly for forms
Add-Type -AssemblyName System.Windows.Forms

# Create a form to display progress
$form = New-Object System.Windows.Forms.Form
$form.Text = "Processing Files"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10, 20)
$label.Size = New-Object System.Drawing.Size(380, 40)
$label.Text = "Processing files. Please wait..."
$form.Controls.Add($label)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 70)
$progressBar.Size = New-Object System.Drawing.Size(360, 30)
$form.Controls.Add($progressBar)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 110)
$statusLabel.Size = New-Object System.Drawing.Size(380, 40)
$statusLabel.Text = "Initializing..."
$form.Controls.Add($statusLabel)

# Show the form without blocking
$form.Show()
$form.Refresh()

# Ensure folder paths end with a backslash
if ($csvFolder[-1] -ne '\') { $csvFolder += '\' }
if ($destFolder[-1] -ne '\') { $destFolder += '\' }
if ($templateDestFolder[-1] -ne '\') { $templateDestFolder += '\' }

# Step 1: Create daily template file if it doesn't exist
$statusLabel.Text = "Preparing daily template file..."
$form.Refresh()

try {
    # Get current date in MMDD format
    $currentDate = Get-Date -Format "MMdd"
    $newTemplateFileName = "PH IV - FCC Inventory $currentDate.xlsm"
    $newTemplateFilePath = Join-Path -Path $templateDestFolder -ChildPath $newTemplateFileName
    
    # Check if the file already exists
    if (Test-Path -Path $newTemplateFilePath) {
        # Get file sizes to compare
        $existingFileSize = (Get-Item -Path $newTemplateFilePath).Length
        $templateFileSize = (Get-Item -Path $templateSourcePath).Length
        
        if ($existingFileSize -gt ($templateFileSize * 1.5)) {
            # File exists and is significantly larger than template, skip copying
            $statusLabel.Text = "Daily file already exists and appears to be in use. Skipping template creation."
        } else {
            # File exists but might be a fresh copy, ask user
            $result = [System.Windows.Forms.MessageBox]::Show(
                "A file named '$newTemplateFileName' already exists but appears to be a fresh template. Replace it?",
                "File Exists",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                Copy-Item -Path $templateSourcePath -Destination $newTemplateFilePath -Force
                $statusLabel.Text = "Replaced existing template with a fresh copy."
            } else {
                $statusLabel.Text = "Keeping existing template file. Continuing with process."
            }
        }
    } else {
        # File doesn't exist, create it
        Copy-Item -Path $templateSourcePath -Destination $newTemplateFilePath
        $statusLabel.Text = "Created daily template: $newTemplateFileName"
    }
    
    $form.Refresh()
    Start-Sleep -Seconds 1
} catch {
    $statusLabel.Text = "Error creating template: $_"
    $form.Refresh()
    Start-Sleep -Seconds 2
}

# Check if the source folder exists
if (-Not (Test-Path -Path $csvFolder -PathType Container)) {
    [System.Windows.Forms.MessageBox]::Show("The source folder '$csvFolder' was not found.", "Folder Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $browser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvFolder = $browser.SelectedPath
        if ($csvFolder[-1] -ne '\') { $csvFolder += '\' }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No folder selected. Exiting script.", "Operation Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
        $form.Close()
        exit
    }
}

# Check if the destination folder exists
if (-Not (Test-Path -Path $destFolder -PathType Container)) {
    [System.Windows.Forms.MessageBox]::Show("The destination folder '$destFolder' was not found.", "Folder Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $result = $browser.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $destFolder = $browser.SelectedPath
        if ($destFolder[-1] -ne '\') { $destFolder += '\' }
    } else {
        [System.Windows.Forms.MessageBox]::Show("No folder selected. Exiting script.", "Operation Cancelled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Exclamation)
        $form.Close()
        exit
    }
}

# Clean up existing XLSX files in the destination folder
$statusLabel.Text = "Cleaning up existing files in destination folder..."
$form.Refresh()

try {
    $existingFiles = Get-ChildItem -Path $destFolder -Filter "*.xlsx" | Where-Object { $_.Name -like "*ExportWorklist*" }
    $cleanupCount = $existingFiles.Count
    
    if ($cleanupCount -gt 0) {
        foreach ($file in $existingFiles) {
            Remove-Item -Path $file.FullName -Force
        }
        $statusLabel.Text = "Removed $cleanupCount existing files from destination folder."
    } else {
        $statusLabel.Text = "No existing files to clean up."
    }
    $form.Refresh()
    Start-Sleep -Seconds 1
} catch {
    $statusLabel.Text = "Error cleaning up destination folder: $_"
    $form.Refresh()
    Start-Sleep -Seconds 2
}

# Update status
$statusLabel.Text = "Creating Excel instance..."
$form.Refresh()

# Create Excel application instance
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $excel.Visible = $false  # Keep Excel hidden during processing
    $excel.ScreenUpdating = $false  # Turn off screen updating for better performance
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to create Excel instance: $_", "Excel Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    $form.Close()
    exit
}

# Process all CSV files in the source folder
$csvFiles = Get-ChildItem -Path $csvFolder -Filter "*.csv" | Where-Object { $_.Name -like "*ExportWorklist*" }
$fileCount = $csvFiles.Count

if ($fileCount -eq 0) {
    $statusLabel.Text = "No matching CSV files found."
    $form.Refresh()
    Start-Sleep -Seconds 2
} else {
    $progressBar.Minimum = 0
    $progressBar.Maximum = $fileCount
    $progressBar.Value = 0
    
    $currentFile = 0
    foreach ($file in $csvFiles) {
        $currentFile++
        $progressBar.Value = $currentFile
        $statusLabel.Text = "Processing $($currentFile) of $($fileCount): $($file.Name)"
        $form.Refresh()
        
        try {
            $workbook = $excel.Workbooks.Open($file.FullName)
            $worksheet = $workbook.Sheets.Item(1)
            $newFileName = Join-Path -Path $destFolder -ChildPath ($file.BaseName + ".xlsx")
            $workbook.SaveAs($newFileName, 51) # xlOpenXMLWorkbook
            $workbook.Close($false)
            Remove-Item -Path $file.FullName
        } catch {
            $statusLabel.Text = "Error processing $($file.Name): $_"
            $form.Refresh()
            Start-Sleep -Seconds 2
        }
    }
}

# Re-enable Excel alerts and make it visible for the tool file
$excel.DisplayAlerts = $true
$excel.ScreenUpdating = $true
$excel.Visible = $true

# Update status
$statusLabel.Text = "All files processed. Opening tool file..."
$form.Refresh()

# Close the progress form
$form.Close()

# Clean up COM objects properly
if ($worksheet) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) | Out-Null }
if ($workbook) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null }
if ($excel) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
Remove-Variable -Name worksheet, workbook, excel -ErrorAction SilentlyContinue

# Force garbage collection to release COM objects
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
