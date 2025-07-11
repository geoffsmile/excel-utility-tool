<#
.SYNOPSIS
    Processes various attendance files, converting and organizing them into a structured format.
.DESCRIPTION
    This script searches for specific attendance files, converts text files to Excel,
    archives previous shift data, and organizes everything into appropriate folders.
.NOTES
    Author: GEOFF LU
    Last Updated: April 2025
#>

# Configuration - Edit these paths as needed
$config = @{
    SourceFolder      = "C:\Users\PH10035990\Downloads"
    DestinationFolder = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Attendance\Raw Files"
    ArchiveFolder     = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Attendance\Raw Files\Archive"
    FinalReportPath   = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\01 - FCC Files\Attendance\FCC and CC Comprehensive Attendance Report 2025 Q2.xlsb"
    ExpectedCount     = 6  # Expected number of "Agent Shifts" and "Activities" files
    FilePrefixes      = @("Previous Shift", "Activities", "Agent Shifts", "Attendance Roster", "Shift_Roster")
    LogFile           = "C:\Users\PH10035990\OneDrive - R1\SHARED FOLDER\MIS Sharepoint\04 - MIS Personal Folders\Geoff Lu\Process Log\ProcessLog.txt"
}

# Create a log function
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with color coding
    switch ($Level) {
        "INFO"    { Write-Host $logMessage -ForegroundColor Green }
        "WARNING" { Write-Host $logMessage -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logMessage -ForegroundColor Red }
    }
    
    # Append to log file
    Add-Content -Path $config.LogFile -Value $logMessage
}

# Get last business day function
function Get-LastBusinessDay {
    $currentDate = Get-Date
    $dayOfWeek = $currentDate.DayOfWeek
    
    switch ($dayOfWeek) {
        "Sunday"   { return $currentDate.AddDays(-2) } # Go back to Friday
        "Monday"   { return $currentDate.AddDays(-3) } # Go back to Friday
        default    { return $currentDate.AddDays(-1) } # Go back one day
    }
}

# Function to convert text file to Excel
function Convert-TextToExcel {
    param (
        [string]$SourcePath,
        [string]$DestinationPath
    )
    
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        # Open the text file
        $workbook = $excel.Workbooks.Open($SourcePath)
        
        # Save as Excel file
        $workbook.SaveAs($DestinationPath, 51) # 51 = xlsx format
        
        # Close and clean up
        $workbook.Close($false)
        $excel.Quit()
        
        # Release COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Log "Successfully converted $SourcePath to $DestinationPath"
        return $true
    }
    catch {
        Write-Log "Error converting file: $_" -Level "ERROR"
        return $false
    }
}

# Function to display notification
function Show-Notification {
    param (
        [string]$Message,
        [int]$Duration = 7
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $notification = New-Object System.Windows.Forms.NotifyIcon
    $notification.Icon = [System.Drawing.SystemIcons]::Information
    $notification.BalloonTipTitle = "File Processing Status"
    $notification.BalloonTipText = $Message
    $notification.Visible = $true
    
    # Show balloon tip
    $notification.ShowBalloonTip($Duration * 1000)
    
    # Schedule cleanup
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = $Duration * 1000
    $timer.add_Tick({
        $notification.Dispose()
        $timer.Stop()
        $timer.Dispose()
    })
    $timer.Start()
    
    # Also write to log
    Write-Log $Message
}

# Main script execution starts here
try {
    Write-Log "=== Starting File Processing Job ==="
    
    # Check/create folders
    if (-not (Test-Path $config.SourceFolder)) {
        Write-Log "Source folder does not exist: $($config.SourceFolder)" -Level "ERROR"
        exit
    }
    
    if (-not (Test-Path $config.DestinationFolder)) {
        New-Item -Path $config.DestinationFolder -ItemType Directory | Out-Null
        Write-Log "Created destination folder: $($config.DestinationFolder)"
    }
    
    if (-not (Test-Path $config.ArchiveFolder)) {
        New-Item -Path $config.ArchiveFolder -ItemType Directory | Out-Null
        Write-Log "Created archive folder: $($config.ArchiveFolder)"
    }
    
    # Get the last business day
    $lastBusinessDay = Get-LastBusinessDay
    $dateFormat = @{
        Year = $lastBusinessDay.Year
        Month = $lastBusinessDay.Month.ToString("00")
        Day = $lastBusinessDay.Day.ToString("00")
    }
    
    # Delete old Activities files
    Get-ChildItem -Path $config.DestinationFolder -Filter "Activities*" | ForEach-Object {
        Remove-Item $_.FullName -Force
        Write-Log "Deleted old Activities file: $($_.Name)"
    }
    # Delete old Agent Shifts files
    Get-ChildItem -Path $config.DestinationFolder -Filter "Agent Shifts*" | ForEach-Object {
        Remove-Item $_.FullName -Force
        Write-Log "Deleted old Agent Shifts file: $($_.Name)"
    }	  
    # Count Agent Shifts and Activities files
    $sourceFiles = Get-ChildItem -Path $config.SourceFolder
    $agentShiftsCount = ($sourceFiles | Where-Object { $_.Name -like "Agent Shifts*" }).Count
    $activitiesCount = ($sourceFiles | Where-Object { $_.Name -like "Activities*" }).Count
    
    # Check if counts match expected
    if ($agentShiftsCount -ne $config.ExpectedCount -or $activitiesCount -ne $config.ExpectedCount) {
        # Open explorer to the source folder
        Start-Process "explorer.exe" -ArgumentList $config.SourceFolder
        
        $message = "The number of files does not match what was expected!`n" +
                   "Agent Shifts: $agentShiftsCount (expected $($config.ExpectedCount))`n" +
                   "Activities: $activitiesCount (expected $($config.ExpectedCount))"
        
        Show-Notification $message
        Write-Log $message -Level "WARNING"
        exit
    }
    
    # Track processed files
    $processedFiles = @{}
    $notProcessedFiles = @{}
    
    # Process each file
    foreach ($file in $sourceFiles) {
        $fileName = $file.Name
        $matchesPrefix = $false
        
        # Check if file matches any prefix
        foreach ($prefix in $config.FilePrefixes) {
            if ($fileName.StartsWith($prefix)) {
                $matchesPrefix = $true
                break
            }
        }
        
        if ($matchesPrefix -or $fileName.StartsWith("Activities")) {
            # Process based on file type
            if ($fileName.StartsWith("Previous Shift")) {
                # First, copy to archive with date-based name
                $newFileName = "$($dateFormat.Year)-R1-$($dateFormat.Month)-$($dateFormat.Day).csv"
                $archivePath = Join-Path $config.ArchiveFolder $newFileName
                
                # Delete existing file if it exists in archive
                if (Test-Path $archivePath) { 
                    Remove-Item $archivePath -Force 
                    Write-Log "Deleted existing archive file: $newFileName"
                }
                
                # Copy file to archive with new name
                Copy-Item -Path $file.FullName -Destination $archivePath -Force
                
                # Then also move the original file to destination folder
                $destPath = Join-Path $config.DestinationFolder $fileName
                
                # Delete existing file if it exists in destination
                if (Test-Path $destPath) { 
                    Remove-Item $destPath -Force 
                    Write-Log "Deleted existing destination file: $fileName"
                }
                
                # Move file to destination
                Move-Item -Path $file.FullName -Destination $destPath -Force
                
                $processedFiles[$fileName] = "Moved to destination and archived as $newFileName"
                Write-Log "Processed '$fileName' - Moved and archived as '$newFileName'"
                
            } elseif ($fileName.StartsWith("Agent Shifts")) {
                $destTxtPath = Join-Path $config.DestinationFolder $fileName
                $destXlsxPath = $destTxtPath -replace "\.txt$", ".xlsx"
                
                # Delete existing XLSX if it exists
                if (Test-Path $destXlsxPath) { 
                    Remove-Item $destXlsxPath -Force 
                    Write-Log "Deleted existing Excel file: $(Split-Path $destXlsxPath -Leaf)"
                }
                
                # Move TXT file
                Move-Item -Path $file.FullName -Destination $destTxtPath -Force
                
                # Convert to Excel
                $success = Convert-TextToExcel -SourcePath $destTxtPath -DestinationPath $destXlsxPath
                
                # Delete the original TXT file
                if ($success -and (Test-Path $destTxtPath)) { 
                    Remove-Item $destTxtPath -Force 
                    $processedFiles[$fileName] = "Moved and converted to $(Split-Path $destXlsxPath -Leaf)"
                    Write-Log "Processed '$fileName' - Converted to Excel format"
                }
                
            } elseif ($fileName.StartsWith("Attendance Roster") -or $fileName.StartsWith("Shift_Roster") -or $fileName.StartsWith("Activities")) {
                $destPath = Join-Path $config.DestinationFolder $fileName
                
                # Delete existing file if it exists
                if (Test-Path $destPath) { 
                    Remove-Item $destPath -Force 
                    Write-Log "Deleted existing file: $fileName"
                }
                
                # Move file to destination
                Move-Item -Path $file.FullName -Destination $destPath -Force
                $processedFiles[$fileName] = "Moved to destination folder"
                Write-Log "Moved '$fileName' to destination folder"
            } else {
                $notProcessedFiles[$fileName] = "Did not match specific processing criteria"
            }
        } else {
            $notProcessedFiles[$fileName] = "Did not match any prefix"
        }
    }
    
    # Expected file count calculation (similar logic to original)
    $expectedProcessedCount = $config.FilePrefixes.Count + 2 - 1 - 1
    
    # Check if all expected files were processed
    if ($processedFiles.Count -eq $expectedProcessedCount) {
        Write-Log "All expected files processed successfully. Opening final report..."
        Start-Process $config.FinalReportPath
    } else {
        # Build summary message
        $summary = "File Processing Summary:`n`n"
        $summary += "Processed Files: $($processedFiles.Count) (Expected: $expectedProcessedCount)`n"
        
        if ($processedFiles.Count -gt 0) {
            foreach ($key in $processedFiles.Keys) {
                $summary += "- $key ($($processedFiles[$key]))`n"
            }
        } else {
            $summary += "- None`n"
        }
        
        $summary += "`nNot Processed Files: $($notProcessedFiles.Count)`n"
        if ($notProcessedFiles.Count -gt 0) {
            foreach ($key in $notProcessedFiles.Keys) {
                $summary += "- $key ($($notProcessedFiles[$key]))`n"
            }
        } else {
            $summary += "- None`n"
        }
        
        Show-Notification $summary
    }
    
    Write-Log "=== File Processing Job Completed ==="
}
catch {
    Write-Log "Critical error: $_" -Level "ERROR"
    Show-Notification "An error occurred during file processing. Check the log file for details."
}
finally {
    # Final cleanup of any remaining objects
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
