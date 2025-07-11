# ==============================================================================
# Excel Utility Tool - Version 4
# Author: Geoff Lu | geoffsmile@gmail.com
# Created: July 1, 2025
# Last Modified: July 11, 2025
# Description:
#       The Excel Utility Tool provides a user-friendly interface for common 
#       Excel file operations. It supports converting CSV, TXT, and XLS files 
#       to the modern XLSX format and unlocking password-protected files. 
#       The tool features a settings panel for customization, detailed logging, 
#       and a comprehensive help system.
# ==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# --- Global Configuration ---
$global:Config = @{
    Version = "4.0"
    Colors = @{
        Primary   = [System.Drawing.Color]::FromArgb(0, 71, 187)
        Secondary = [System.Drawing.Color]::White
        Text      = [System.Drawing.Color]::FromArgb(64, 64, 64)
        Success   = [System.Drawing.Color]::FromArgb(76, 175, 80)
        Warning   = [System.Drawing.Color]::FromArgb(255, 193, 7)
        Error     = [System.Drawing.Color]::FromArgb(244, 67, 54)
    }
    Settings = @{
        DefaultInputPath     = ""
        DefaultOutputPath    = ""
        RememberPaths        = $true
        AutoDeleteOriginals  = $false
        ShowDetailedLogs     = $false
        SoundNotifications   = $true
    }
    SettingsPath = Join-Path $PSScriptRoot "settings.json"
    LogPath      = Join-Path $PSScriptRoot "logs"
}

# --- Utility Functions ---

function Write-Log {
    param([string]$Message, [ValidateSet("INFO","WARNING","ERROR")][string]$Level = "INFO")
    try {
        if (-not (Test-Path $global:Config.LogPath)) {
            New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logFile = Join-Path $global:Config.LogPath "excel_utility_$(Get-Date -Format 'yyyy-MM-dd').log"
        $logEntry = "[$timestamp][$Level] $Message"
        Add-Content -Path $logFile -Value $logEntry
        if ($global:Config.Settings.ShowDetailedLogs) {
            $color = switch ($Level) { "ERROR" { "Red" } "WARNING" { "Yellow" } default { "Gray" } }
            Write-Host $logEntry -ForegroundColor $color
        }
    } catch {
        Write-Host "Log write failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Load-Settings {
    try {
        if (Test-Path $global:Config.SettingsPath) {
            $json = Get-Content $global:Config.SettingsPath -Raw | ConvertFrom-Json
            foreach ($key in $json.PSObject.Properties.Name) {
                if ($global:Config.Settings.ContainsKey($key)) {
                    $global:Config.Settings[$key] = $json.$key
                }
            }
            Write-Log "Settings loaded"
        }
    } catch { Write-Log "Failed to load settings: $($_.Exception.Message)" "ERROR" }
}

function Save-Settings {
    try {
        $global:Config.Settings | ConvertTo-Json | Set-Content $global:Config.SettingsPath
        Write-Log "Settings saved"
        return $true
    } catch {
        Write-Log "Saving settings failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Path-Valid([string]$Path) {
    return (-not [string]::IsNullOrWhiteSpace($Path)) -and (Test-Path $Path)
}

function Play-Sound($Type) {
    if ($global:Config.Settings.SoundNotifications) {
        switch ($Type) {
            "Success" { [System.Media.SystemSounds]::Exclamation.Play() }
            "Error"   { [System.Media.SystemSounds]::Hand.Play() }
        }
    }
}

# --- UI Helper Functions ---

class StandardButton : System.Windows.Forms.Button {
    StandardButton() {
        $this.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
        $this.BackColor = $global:Config.Colors.Primary
        $this.ForeColor = $global:Config.Colors.Secondary
        $this.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $this.Height = 30
    }
}

function Show-MessageBox($Message, $Title = "Excel Utility", $Type = "Information") {
    $icon = switch ($Type) {
        "Error"   { Play-Sound "Error";   [System.Windows.Forms.MessageBoxIcon]::Error }
        "Warning" { Play-Sound "Error";   [System.Windows.Forms.MessageBoxIcon]::Warning }
        default   {                      [System.Windows.Forms.MessageBoxIcon]::Information }
    }
    [System.Windows.Forms.MessageBox]::Show($Message, $Title, "OK", $icon)
}

function Show-HelpDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Excel Utility Tool - Help"
    $form.Size = New-Object System.Drawing.Size(800,600)
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false; $form.MinimizeBox = $false
    $form.BackColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Size = New-Object System.Drawing.Size(780,550)
    $tabs.Location = New-Object System.Drawing.Point(10,10)
    $form.Controls.Add($tabs)

    function Add-Tab($title, $text) {
        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $title
        $box = New-Object System.Windows.Forms.RichTextBox
        $box.Text = $text
        $box.Size = New-Object System.Drawing.Size(760,520)
        $box.Location = New-Object System.Drawing.Point(10,10)
        $box.ReadOnly = $true
        $box.BackColor = [System.Drawing.Color]::White
        $tab.Controls.Add($box)
        $tabs.Controls.Add($tab)
    }

    Add-Tab "Overview" @"
EXCEL UTILITY TOOL - VERSION 4
Created by: Geoff Lu | geoffsmile@gmail.com
Created: July 1, 2025 | Last Modified: July 7, 2025

A tool for business professionals, analysts & staff to streamline document workflows:
- Convert CSV, TXT, XLS to modern XLSX format
- Remove password protection (all files must use same password)
- Batch processing, settings, logs, and progress tracking
"@

    Add-Tab "Features" @"
Key Features:
• File format conversion (CSV, TXT, XLS → XLSX)
• Password removal (same password required for batch)
• Batch processing
• Settings & logging
• User-friendly interface

Requirements:
• Windows 10/11, PowerShell 5.1+, .NET 4.7.2+, MS Excel, 2GB+ RAM

Limitations:
• Only removes known passwords
• Large files may process slowly
• Some Excel features may not be preserved
"@

    Add-Tab "Instructions" @"
Converting Files:
1. Click 'Convert to Excel (.xlsx)'
2. Select files (CSV, TXT, XLS)
3. Select output folder
4. Wait for completion

Unlocking Excel Files:
1. Click 'Unlock Excel File'
2. Enter password (must be same for all)
3. Select locked Excel files
4. Select output folder

Settings:
1. Click 'Application Tool Settings'
2. Adjust paths, logging, sound, etc.
3. Save changes
"@

    Add-Tab "FAQ" @"
Q: What formats are supported?
A: CSV, TXT, XLS for conversion to XLSX.

Q: Batch processing?
A: Yes, select multiple files.

Q: Forgotten password?
A: Password is required—cannot remove unknown.

Q: Are originals deleted?
A: Only if 'Auto Delete Originals' is enabled.

Q: Where are logs?
A: In the 'logs' folder beside the script.

Q: Office 365 supported?
A: Yes—Office 365/2019/2021 supported.
"@

    Add-Tab "About" @"
Author: Geoff Lu | geoffsmile@gmail.com
Reporting Analyst, Developer & Automation Enthusiast

Philosophy: Intuitive, robust, flexible, and reliable tools.
Versions: 1.0 (basic) → 2.0 (password) → 3.0 (UI/settings) → 4.0 (help/optimizations)
License: © 2025 Geoff Lu. For educational/business use. Redistribution for personal/internal business only.
"@

    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Close"
    $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
    $closeBtn.Location = New-Object System.Drawing.Point(350, 565)
    $closeBtn.DialogResult = "OK"
    $form.Controls.Add($closeBtn)
    $form.AcceptButton = $closeBtn
    $form.ShowDialog() | Out-Null
}

function Show-SettingsDialog {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Application Tool Settings"
    $f.Size = New-Object System.Drawing.Size(550,400)
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false; $f.MinimizeBox = $false
    $f.BackColor = [System.Drawing.Color]::White
    $f.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="Application Tool Settings"
        Font=New-Object System.Drawing.Font("Segoe UI",13,[System.Drawing.FontStyle]::Bold)
        ForeColor=$global:Config.Colors.Primary
        Size=New-Object System.Drawing.Size(460,25)
        Location=New-Object System.Drawing.Point(20,15)
    }))

    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{Text="Default Paths"; Font=(New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)); ForeColor=$global:Config.Colors.Text; Size=New-Object System.Drawing.Size(150,20); Location=New-Object System.Drawing.Point(20,50)}))
    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{Text="Default Input:"; ForeColor=$global:Config.Colors.Text; Size=New-Object System.Drawing.Size(100,20); Location=New-Object System.Drawing.Point(30,80)}))
    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{Text="Default Output:"; ForeColor=$global:Config.Colors.Text; Size=New-Object System.Drawing.Size(100,20); Location=New-Object System.Drawing.Point(30,110)}))

    $txtInput = New-Object System.Windows.Forms.TextBox -Property @{Size=New-Object System.Drawing.Size(300,20); Location=New-Object System.Drawing.Point(130,78); BackColor=[System.Drawing.Color]::FromArgb(230,230,230); Text=$global:Config.Settings.DefaultInputPath}
    $txtOutput = New-Object System.Windows.Forms.TextBox -Property @{Size=New-Object System.Drawing.Size(300,20); Location=New-Object System.Drawing.Point(130,108); BackColor=[System.Drawing.Color]::FromArgb(230,230,230); Text=$global:Config.Settings.DefaultOutputPath}
    $f.Controls.Add($txtInput); $f.Controls.Add($txtOutput)

    $btnBrowseInput = [StandardButton]::new(); $btnBrowseInput.Text = "Browse..."; $btnBrowseInput.Size = New-Object System.Drawing.Size(80,25); $btnBrowseInput.Location = New-Object System.Drawing.Point(440,75)
    $btnBrowseOutput = [StandardButton]::new(); $btnBrowseOutput.Text = "Browse..."; $btnBrowseOutput.Size = New-Object System.Drawing.Size(80,25); $btnBrowseOutput.Location = New-Object System.Drawing.Point(440,105)
    $f.Controls.Add($btnBrowseInput); $f.Controls.Add($btnBrowseOutput)

    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{Text="Application Settings"; Font=(New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)); ForeColor=$global:Config.Colors.Text; Size=New-Object System.Drawing.Size(150,20); Location=New-Object System.Drawing.Point(20,150)}))

    $chkRemember = New-Object System.Windows.Forms.CheckBox -Property @{Text="Remember selected paths as defaults"; Size=New-Object System.Drawing.Size(400,20); Location=New-Object System.Drawing.Point(30,180); Checked=$global:Config.Settings.RememberPaths}
    $chkDelete   = New-Object System.Windows.Forms.CheckBox -Property @{Text="Automatically delete original files after processing"; Size=New-Object System.Drawing.Size(400,20); Location=New-Object System.Drawing.Point(30,210); Checked=$global:Config.Settings.AutoDeleteOriginals}
    $chkLogs     = New-Object System.Windows.Forms.CheckBox -Property @{Text="Show detailed logs in console"; Size=New-Object System.Drawing.Size(400,20); Location=New-Object System.Drawing.Point(30,240); Checked=$global:Config.Settings.ShowDetailedLogs}
    $chkSound    = New-Object System.Windows.Forms.CheckBox -Property @{Text="Enable sound notifications"; Size=New-Object System.Drawing.Size(400,20); Location=New-Object System.Drawing.Point(30,270); Checked=$global:Config.Settings.SoundNotifications}
    $f.Controls.Add($chkRemember); $f.Controls.Add($chkDelete); $f.Controls.Add($chkLogs); $f.Controls.Add($chkSound)

    $btnOK = [StandardButton]::new(); $btnOK.Text="Save"; $btnOK.Size=New-Object System.Drawing.Size(80,30); $btnOK.Location=New-Object System.Drawing.Point(350,320); $btnOK.DialogResult="OK"
    $btnCancel = [StandardButton]::new(); $btnCancel.Text="Cancel"; $btnCancel.Size=New-Object System.Drawing.Size(80,30); $btnCancel.Location=New-Object System.Drawing.Point(440,320); $btnCancel.DialogResult="Cancel"
    $f.Controls.Add($btnOK); $f.Controls.Add($btnCancel)
    $f.AcceptButton = $btnOK; $f.CancelButton = $btnCancel

    $btnBrowseInput.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Select Default Input Folder"
        if (-not [string]::IsNullOrWhiteSpace($txtInput.Text)) { $d.SelectedPath = $txtInput.Text }
        if ($d.ShowDialog() -eq "OK") { $txtInput.Text = $d.SelectedPath }
    })
    $btnBrowseOutput.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Select Default Output Folder"
        if (-not [string]::IsNullOrWhiteSpace($txtOutput.Text)) { $d.SelectedPath = $txtOutput.Text }
        if ($d.ShowDialog() -eq "OK") { $txtOutput.Text = $d.SelectedPath }
    })

    $btnOK.Add_Click({
        if ($txtInput.Text -and -not (Test-Path $txtInput.Text)) {
            Show-MessageBox "Input path doesn't exist: $($txtInput.Text)" "Invalid Path" "Warning"; return
        }
        if ($txtOutput.Text -and -not (Test-Path $txtOutput.Text)) {
            Show-MessageBox "Output path doesn't exist: $($txtOutput.Text)" "Invalid Path" "Warning"; return
        }
        $global:Config.Settings.DefaultInputPath = $txtInput.Text
        $global:Config.Settings.DefaultOutputPath = $txtOutput.Text
        $global:Config.Settings.RememberPaths = $chkRemember.Checked
        $global:Config.Settings.AutoDeleteOriginals = $chkDelete.Checked
        $global:Config.Settings.ShowDetailedLogs = $chkLogs.Checked
        $global:Config.Settings.SoundNotifications = $chkSound.Checked
        if (Save-Settings) { Play-Sound "Success"; $f.Close() }
        else { Show-MessageBox "Failed to save settings." "Error" "Error" }
    })

    $f.ShowDialog() | Out-Null
}

function Get-InputFiles($Title="Select files to process", $Filter="Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*") {
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Multiselect = $true
    $ofd.Title = $Title
    $ofd.Filter = $Filter
    if (Path-Valid $global:Config.Settings.DefaultInputPath) { $ofd.InitialDirectory = $global:Config.Settings.DefaultInputPath }
    if ($ofd.ShowDialog() -eq "OK") { return $ofd.FileNames }
    return $null
}

function Get-OutputFolder($Title="Select output folder") {
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    $d.Description = $Title
    if (Path-Valid $global:Config.Settings.DefaultOutputPath) { $d.SelectedPath = $global:Config.Settings.DefaultOutputPath }
    if ($d.ShowDialog() -eq "OK") { return $d.SelectedPath }
    return $null
}

function Get-PasswordInput($Title="Enter Password") {
    $form = New-Object System.Windows.Forms.Form -Property @{
        Text=$Title; Size=New-Object System.Drawing.Size(380,205)
        StartPosition="CenterParent"; FormBorderStyle="FixedDialog"; MaximizeBox=$false; MinimizeBox=$false
        BackColor=$global:Config.Colors.Secondary; Font=(New-Object System.Drawing.Font("Segoe UI",9))
    }
    $form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="Please enter the password used to protect these files"
        ForeColor=$global:Config.Colors.Primary; Font=(New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold))
        Size=New-Object System.Drawing.Size(340,30); Location=New-Object System.Drawing.Point(20,15)
        TextAlign=[System.Windows.Forms.HorizontalAlignment]::Center
    }))
    $form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="Password:"; ForeColor=$global:Config.Colors.Text; Size=New-Object System.Drawing.Size(80,20)
        Location=New-Object System.Drawing.Point(70,45)
    }))
    $txtPassword = New-Object System.Windows.Forms.TextBox -Property @{
        Size=New-Object System.Drawing.Size(240,20); Location=New-Object System.Drawing.Point(70,65)
        BackColor=[System.Drawing.Color]::FromArgb(230,230,230); UseSystemPasswordChar=$false
    }
    $form.Controls.Add($txtPassword)
    $btnOK = [StandardButton]::new(); $btnOK.Text="OK"; $btnOK.Size=New-Object System.Drawing.Size(100,32); $btnOK.Location=New-Object System.Drawing.Point(140,100); $btnOK.DialogResult="OK"
    $form.Controls.Add($btnOK)
    $form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="This tool can only remove passwords from files if they all share the same password."
        ForeColor=$global:Config.Colors.Text; Font=(New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic))
        Size=New-Object System.Drawing.Size(340,40); Location=New-Object System.Drawing.Point(20,140)
        TextAlign=[System.Windows.Forms.HorizontalAlignment]::Center
    }))
    $form.AcceptButton = $btnOK
    if ($form.ShowDialog() -eq "OK") { return $txtPassword.Text }
    return $null
}

# --- Main UI ---

function Initialize-MainForm {
    Load-Settings
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Excel Utility Tool v$($global:Config.Version) - Geoff Lu"
    $f.Size = New-Object System.Drawing.Size(400,390)
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false
    $f.BackColor = $global:Config.Colors.Secondary
    $f.Font = New-Object System.Drawing.Font("Segoe UI",8)

    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="Excel Utility Tool"; Font=(New-Object System.Drawing.Font("Segoe UI",14,[System.Drawing.FontStyle]::Bold))
        ForeColor=$global:Config.Colors.Primary; Size=New-Object System.Drawing.Size(360,25); Location=New-Object System.Drawing.Point(20,20)
    }))
    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="Version $($global:Config.Version)"; Size=New-Object System.Drawing.Size(360,20); Location=New-Object System.Drawing.Point(20,45)
        ForeColor=$global:Config.Colors.Text
    }))

    $script:btnConvert = [StandardButton]::new(); 
    $script:btnConvert.Text="Convert to Excel (.xlsx)"; 
    $script:btnConvert.Size=New-Object System.Drawing.Size(360,35); 
    $script:btnConvert.Location=New-Object System.Drawing.Point(20,80)
    
    $script:btnUnlock = [StandardButton]::new(); 
    $script:btnUnlock.Text="Unlock Excel File"; 
    $script:btnUnlock.Size=New-Object System.Drawing.Size(360,35); 
    $script:btnUnlock.Location=New-Object System.Drawing.Point(20,125)
    
    $script:btnSettings = [StandardButton]::new(); 
    $script:btnSettings.Text="Application Tool Settings"; 
    $script:btnSettings.Size=New-Object System.Drawing.Size(360,35); 
    $script:btnSettings.Location=New-Object System.Drawing.Point(20,170)
    
    $btnHelp = [StandardButton]::new(); $btnHelp.Text="Help && Information"; 
    $btnHelp.Size=New-Object System.Drawing.Size(360,35); 
    $btnHelp.Location=New-Object System.Drawing.Point(20,215)
    
    $f.Controls.AddRange(@($script:btnConvert,$script:btnUnlock,$script:btnSettings,$btnHelp))

    $script:status = New-Object System.Windows.Forms.Label -Property @{Text="Ready!"; Size=New-Object System.Drawing.Size(360,20); Location=New-Object System.Drawing.Point(20,260); ForeColor=$global:Config.Colors.Text}
    $script:progress= New-Object System.Windows.Forms.ProgressBar -Property @{Size=New-Object System.Drawing.Size(360,20); Location=New-Object System.Drawing.Point(20,280)}
    $f.Controls.Add($script:status); $f.Controls.Add($progress)
    $f.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
        Text="© 2025 Geoff Lu. All rights reserved.`nCreated July 1, 2025"
        Size=New-Object System.Drawing.Size(360,40); Location=New-Object System.Drawing.Point(20,320)
        ForeColor=$global:Config.Colors.Text; Font=(New-Object System.Drawing.Font("Segoe UI",8))
        TextAlign="MiddleCenter"
    }))

    $script:btnConvert.Add_Click({
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $status.Text = "Selecting files..."
        $files = Get-InputFiles "Select files to convert to .xlsx" "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|Excel files (*.xls)|*.xls|All files (*.*)|*.*"
        if (-not $files) { $status.Text = "Ready"; return }

        $status.Text = "Selecting output folder..."
        $out = Get-OutputFolder "Select output folder for converted files"
        if (-not $out) { $status.Text = "Ready"; return }

        $script:btnConvert.Enabled = $false
        $script:btnUnlock.Enabled = $false
        $script:btnSettings.Enabled = $false

        $status.Text = "Processing files..."
        for ($i=0; $i -le 100; $i++) {
            $progress.Value = $i
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 10
        }

        $status.Text = "Conversion completed!"
        Play-Sound "Success"
    } catch {
        $status.Text = "Error occurred!"
        Show-MessageBox "Conversion error: $($_.Exception.Message)" "Error" "Error"
        Write-Log "Convert error: $($_.Exception.Message)" "ERROR"
    } finally {
if ($workbook) {
    $workbook.Close($false)
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
}
if ($excel) {
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

        $script:btnConvert.Enabled = $true
        $script:btnUnlock.Enabled = $true
        $script:btnSettings.Enabled = $true
        $progress.Value = 0
    }
})


$script:btnUnlock.Add_Click({
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $pw = Get-PasswordInput "Enter Excel Password"
        if (-not $pw) { return }

        $status.Text = "Selecting files..."
        $files = Get-InputFiles "Select password-protected Excel files" "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*"
        if (-not $files) { $status.Text = "Ready"; return }

        $status.Text = "Selecting output folder..."
        $out = Get-OutputFolder "Select output folder for unlocked files"
        if (-not $out) { $status.Text = "Ready"; return }

        $script:btnConvert.Enabled = $false
        $script:btnUnlock.Enabled = $false
        $script:btnSettings.Enabled = $false

        $status.Text = "Processing files..."
        for ($i=0; $i -le 100; $i++) {
            $progress.Value = $i
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 10
        }

        $status.Text = "Password removal completed"
        Play-Sound "Success"
    } catch {
        $status.Text = "Error occurred"
        Show-MessageBox "Password removal error: $($_.Exception.Message)" "Error" "Error"
        Write-Log "Unlock error: $($_.Exception.Message)" "ERROR"
    } finally {
        $workbook.Close($false)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()

        $script:btnConvert.Enabled = $true
        $script:btnUnlock.Enabled = $true
        $script:btnSettings.Enabled = $true
        $progress.Value = 0
    }
})

    $script:btnSettings.Add_Click({ Show-SettingsDialog })
    $btnHelp.Add_Click({ Show-HelpDialog })

    $f.Add_FormClosing({ Write-Log "Application closing" })
    Write-Log "Excel Utility Tool started"
    return $f
}

# --- Application Entry Point ---

if (-not (Test-Path $global:Config.LogPath)) {
    New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
}
try {
    Write-Host "Starting Excel Utility Tool v$($global:Config.Version)..." -ForegroundColor Green
    $mainForm = Initialize-MainForm
    [void]$mainForm.ShowDialog()
} catch {
    Write-Error "Application failed: $($_.Exception.Message)"
    Write-Log "Startup failed: $($_.Exception.Message)" "ERROR"
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to start the Excel Utility Tool:`n`n$($_.Exception.Message)`nPlease check the log files for details.",
        "Startup Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Error
    )
} finally {
    Write-Log "Session ended"
}
