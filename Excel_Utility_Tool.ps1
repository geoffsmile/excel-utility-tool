# ==============================================================================
# Excel Utility Tool - Version 5.7.5
# Author: Geoff Lu | geoffsmile@gmail.com
# Created: June 20, 2025
# Last Modified: July 14, 2025
# Tool Description:
#       The Excel Utility Tool provides a user-friendly interface for common 
#       Excel file operations. It supports converting CSV, TXT, and XLS files 
#       to the modern XLSX format and unlocking password-protected files. 
#       The tool features a settings panel for customization, detailed logging, 
#       and a comprehensive help system.
# ==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# CRITICAL FIX: Validate that the script is being run from a saved file.
# The $PSScriptRoot variable is null if run in an untitled ISE pane, which causes crashes.
if (-not $PSScriptRoot)
{
	[System.Windows.Forms.MessageBox]::Show("This script must be saved to a file before running. Please save the script and try again.", "Startup Error", "OK", "Error")
	exit
}

# --- Global Configuration ---
$global:Config = @{
	Version = "5.7.5" # Updated version number
	Colors  = @{
		Primary = [System.Drawing.Color]::FromArgb(0, 71, 187)
		Secondary = [System.Drawing.Color]::White
		Text    = [System.Drawing.Color]::FromArgb(64, 64, 64)
		Success = [System.Drawing.Color]::FromArgb(76, 175, 80)
		Warning = [System.Drawing.Color]::FromArgb(255, 193, 7)
		Error   = [System.Drawing.Color]::FromArgb(244, 67, 54)
	}
	Settings = @{
		DefaultInputPath    = ""
		DefaultOutputPath   = ""
		RememberPaths	    = $true
		AutoDeleteOriginals = $false
		ShowDetailedLogs    = $false
		SoundNotifications  = $true
	}
	SettingsPath = Join-Path $PSScriptRoot "settings.json"
	LogPath = Join-Path $PSScriptRoot "logs"
}

# --- Utility Functions ---

function Write-Log
{
	param ([string]$Message,
		[ValidateSet("INFO", "WARNING", "ERROR")]
		[string]$Level = "INFO")
	try
	{
		if (-not (Test-Path $global:Config.LogPath))
		{
			New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
		}
		$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
		$logFile = Join-Path $global:Config.LogPath "excel_utility_$(Get-Date -Format 'yyyy-MM-dd').log"
		$logEntry = "[$timestamp][$Level] $Message"
		Add-Content -Path $logFile -Value $logEntry
		if ($global:Config.Settings.ShowDetailedLogs)
		{
			$color = switch ($Level) { "ERROR" { "Red" } "WARNING" { "Yellow" } default { "Gray" } }
			Write-Host $logEntry -ForegroundColor $color
		}
	}
	catch
	{
		Write-Host "Log write failed: $($_.Exception.Message)" -ForegroundColor Red
	}
}

function Load-Settings
{
	try
	{
		if (Test-Path $global:Config.SettingsPath)
		{
			$jsonContent = Get-Content $global:Config.SettingsPath -Raw
			# FIX: Add a null/empty check to handle corrupted settings files gracefully.
			if (-not [string]::IsNullOrWhiteSpace($jsonContent))
			{
				$json = $jsonContent | ConvertFrom-Json
				if ($null -ne $json)
				{
					foreach ($key in $json.PSObject.Properties.Name)
					{
						if ($global:Config.Settings.ContainsKey($key))
						{
							$global:Config.Settings[$key] = $json.$key
						}
					}
					Write-Log "Settings loaded"
				}
				else
				{
					Write-Log "Failed to parse settings.json. The file might be corrupt. Using default settings." "WARNING"
				}
			}
		}
	}
	catch { Write-Log "Failed to load settings: $($_.Exception.Message)" "ERROR" }
}

function Save-Settings
{
	try
	{
		$global:Config.Settings | ConvertTo-Json | Set-Content $global:Config.SettingsPath
		Write-Log "Settings saved"
		return $true
	}
	catch
	{
		Write-Log "Saving settings failed: $($_.Exception.Message)" "ERROR"
		return $false
	}
}

function Path-Valid([string]$Path)
{
	return (-not [string]::IsNullOrWhiteSpace($Path)) -and (Test-Path $Path)
}

function Play-Sound($Type)
{
	if ($global:Config.Settings.SoundNotifications)
	{
		switch ($Type)
		{
			"Success" { [System.Media.SystemSounds]::Exclamation.Play() }
			"Error" { [System.Media.SystemSounds]::Hand.Play() }
		}
	}
}

# --- UI Helper Functions ---

class StandardButton : System.Windows.Forms.Button
{
	StandardButton()
	{
		$this.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
		$this.BackColor = $global:Config.Colors.Primary
		$this.ForeColor = $global:Config.Colors.Secondary
		$this.Font = New-Object System.Drawing.Font("Segoe UI", 9)
		$this.Height = 30
	}
}

function Show-MessageBox($Message, $Title = "Excel Utility", $Type = "Information")
{
	$icon = switch ($Type)
	{
		"Error"   { Play-Sound "Error"; [System.Windows.Forms.MessageBoxIcon]::Error }
		"Warning" { Play-Sound "Error"; [System.Windows.Forms.MessageBoxIcon]::Warning }
		default   { [System.Windows.Forms.MessageBoxIcon]::Information }
	}
	[System.Windows.Forms.MessageBox]::Show($Message, $Title, "OK", $icon)
}

function Show-HelpDialog
{
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Excel Utility Tool - Help"
	$form.Size = New-Object System.Drawing.Size(800, 800)
	$form.MinimumSize = $form.Size
	$form.MaximumSize = $form.Size
	$form.FormBorderStyle = "FixedDialog"
	$form.StartPosition = "CenterParent"
	$form.ForeColor = $global:Config.Colors.Text
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# Header Panel
	$headerPanel = New-Object System.Windows.Forms.Panel
	$headerPanel.Size = New-Object System.Drawing.Size(765, 40)
	$headerPanel.Location = New-Object System.Drawing.Point(10, 10)
	$headerPanel.BackColor = $global:Config.Colors.Primary
	
	$headerLabel = New-Object System.Windows.Forms.Label
	$headerLabel.Text = "HELP CENTER"
	$headerLabel.ForeColor = $global:Config.Colors.Secondary
	$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
	$headerLabel.AutoSize = $true
	$headerLabel.Location = New-Object System.Drawing.Point(20, 10)
	$headerPanel.Controls.Add($headerLabel)
	
	# Tab Control
	$tabControl = New-Object System.Windows.Forms.TabControl
	$tabControl.Size = New-Object System.Drawing.Size(751, 650)
	$tabControl.Location = New-Object System.Drawing.Point(15, 60)
	$tabControl.Appearance = "FlatButtons"
	$tabControl.ItemSize = New-Object System.Drawing.Size(115, 30)
	$tabControl.SizeMode = "Fixed"
	$tabControl.Alignment = "Top"
	$tabControl.SelectedIndex = 0
	
	# Tab Styling
	$tabControl.Padding = New-Object System.Drawing.Point(100, 100)
	$tabControl.DrawMode = "OwnerDrawFixed"
	$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
	$tabControl.Add_DrawItem({
			param ($sender,
				$e)
			$brush = New-Object System.Drawing.SolidBrush($global:Config.Colors.Secondary)
			$bgBrush = if ($e.Index -eq $sender.SelectedIndex)
			{ New-Object System.Drawing.SolidBrush($global:Config.Colors.Primary) }
			else
			{ New-Object System.Drawing.SolidBrush($global:Config.Colors.Text) }
			$e.Graphics.FillRectangle($bgBrush, $e.Bounds)
			$e.Graphics.DrawString($sender.TabPages[$e.Index].Text, $e.Font, $brush, $e.Bounds.X + 8, $e.Bounds.Y + 5)
		})
	
	function Create-HelpTab($title, $content)
	{
		$tabPage = New-Object System.Windows.Forms.TabPage
		$tabPage.Text = $title
		
		$textBox = New-Object System.Windows.Forms.RichTextBox
		$textBox.Text = $content
		$textBox.Dock = "Fill"
		$textBox.ReadOnly = $false
		$textBox.BorderStyle = "FixedSingle"
		$textBox.BackColor = $global:Config.Colors.Secondary
		$textBox.ForeColor = $global:Config.Colors.Text
		$textBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
		$textBox.Margin = New-Object System.Windows.Forms.Padding(1000)
		
		$tabPage.Controls.Add($textBox)
		return $tabPage
	}
	
	# Add Tabs
	$tabControl.Controls.Add((Create-HelpTab "Overview" @"

Excel Utility Tool - Version 5.7.5
• Author: Geoff Lu | geoffsmile@gmail.com
• Created: June 20, 2025
• Last Modified: July 14, 2025

------------------------------------------------------------------------------------------------------------------------------------------------

Tool Description:
The Excel Utility Tool provides a user-friendly interface for common Excel file operations. It supports converting CSV, TXT, and XLS files to the modern XLSX format and unlocking password-protected files. The tool features a settings panel for customization, detailed logging, and a comprehensive help system.

------------------------------------------------------------------------------------------------------------------------------------------------

Development History:
• Version 1 - Initial Launch Introduced the core feature: simple and efficient file conversion from CSV or TXT to XLSX.
• Version 2 - Feature Expansion Integrated password removal capabilities, adding an extra layer of convenience for locked Excel files.
• Version 3 - User Experience Upgrade Rolled out a more refined user interface along with improved settings management for smoother control and customization.
• Version 4 - Help System & Optimization Added a robust Help dialog to guide users, while optimizing performance under the hood for better speed and reliability.
• Version 5 - Accessibility Boost Introduced sound notifications for process completion.

------------------------------------------------------------------------------------------------------------------------------------------------

Included Files:
• ExcelUtilityTool.ps1 - Main application script
• settings.json - User preferences and configuration (auto-generated)
• logs/ - Directory containing operation logs (auto-generated)
• README.txt - Basic usage instructions (if included)

------------------------------------------------------------------------------------------------------------------------------------------------

Tech Specifications & Requirements:
• Built with: PowerShell 5.1+ and .NET Framework
• UI Framework: Windows Forms
• Supported OS: Windows 10/11
• Architecture: Compatible with both x86 and x64 systems
• Dependencies: Microsoft Office Excel (recommended)
• Minimum 2GB RAM (4GB recommended)
• 50MB free disk space

"@))
	
	$tabControl.Controls.Add((Create-HelpTab "Features" @"

Key Features:
• File Format Conversion: Convert CSV, TXT, and legacy XLS files to modern XLSX format
• Password Protection Removal: Unlock password-protected Excel files with ease
• Batch Processing: Handle multiple files simultaneously for improved efficiency
• Customizable Settings: Configure default paths and processing preferences
• Comprehensive Logging: Track all operations with detailed log files
• User-Friendly Interface: Intuitive design suitable for users of all technical levels
• Progress Tracking: Real-time progress indicators for all operations
• Error Handling: Robust error management with clear user feedback

------------------------------------------------------------------------------------------------------------------------------------------------

Known Limitations:
• Password removal requires all files to share the same password
• Large files (>100MB) may require additional processing time
• Some advanced Excel features may not be preserved during conversion
• Network or and Cloud sync drives may cause slower processing speeds
• Corrupted files cannot be processed and will generate error messages

"@))
	
	$tabControl.Controls.Add((Create-HelpTab "Instructions" @"

Converting Files To Excel (.xlsx):

1. Click the "Convert to Excel (.xlsx)" button on the main interface
2. Select the files you want to convert using the file browser dialog
   • Supported formats: CSV, TXT, XLS
   • Multiple files can be selected simultaneously
3. Choose the destination folder where converted files will be saved
4. Wait for the conversion process to complete
5. Check the status bar for completion confirmation and locate your converted files

Unlocking Excel Files:

1. Click the "Unlock Excel File" button on the main interface
2. Enter the password used to protect the Excel files in the password dialog
   • Note: All selected files must share the same password
3. Select the password-protected Excel files you want to unlock
4. Choose the destination folder for the unlocked files
5. Monitor the progress bar and wait for the process to complete

Modifying Tool Settings:

1. Click the "Application Tool Settings" button on the main interface
2. Configure your preferences in the settings dialog:
   • Default Input Path: Set the default folder for selecting input files
   • Default Output Path: Set the default folder for saving processed files
   • Remember Paths: Enable to automatically remember your folder selections
   • Auto Delete Originals: Enable to automatically delete source files after processing
   • Show Detailed Logs: Enable to display detailed processing information
   • Sound Notifications: Enable audio notifications for completed operations
3. Click "Save" to apply your changes or "Cancel" to discard them
4. Settings are automatically saved and will be restored when you restart the tool

------------------------------------------------------------------------------------------------------------------------------------------------

Best Practices:
• Always backup important files before processing
• Use descriptive folder names for better organization
• Regularly check log files for any processing issues
• Keep the tool updated to the latest version for optimal performance
• Test with a small number of files before processing large batches

"@))
	
	$tabControl.Controls.Add((Create-HelpTab "FAQ" @"

Frequently Asked Questions:

Q: Can I process multiple files at once?
A: Yes! The tool is designed for batch processing. You can select multiple files of the same or different formats and process them simultaneously, saving significant time for large operations.

Q: What happens if I forget the password for my Excel files?
A: Unfortunately, the tool can only remove passwords if you know the correct password. This is a security feature to protect your data. If you've forgotten the password, you'll need to use
specialized password recovery software.

Q: Are my original files safe during processing?
A: Yes, the tool creates copies of your files during processing. Your original files remain untouched unless you specifically enable the "Auto Delete Originals" setting in the application
settings.

Q: Where can I find the log files?
A: Log files are automatically created in the logs folder within the same directory as the tool. Each day's activities are logged in a separate file named with the current date.

Q: Can I customize the default folders?
A: Absolutely! Use the "Application Tool Settings" to set default input and output folders. This saves time by automatically opening your preferred locations when browsing for files.

Q: What should I do if the tool crashes or stops responding?
A: First, check the log files for error messages. Restart the tool and try processing fewer files at once.

Q: Is this tool compatible with Office 365?
A: Yes, the tool is compatible with all modern versions of Microsoft Office, including Office 365, Office 2019, and Office 2021.

Q: Can I use this tool on a network drive?
A: While the tool can access network drives, processing may be slower due to network latency. For best performance, consider copying files to your local drive before processing.

"@))
	
	$tabControl.Controls.Add((Create-HelpTab "Troubleshooting" @"

Tool Won't Start:
• Ensure PowerShell execution policy allows script execution
• Verify .NET Framework 4.7.2 or later is installed
• Check that all required assemblies are available
• Make sure the script is saved to a file before running. It cannot be run from an unsaved editor pane.

File Conversion Fails:
• Verify the source file is not corrupted or in use by another application
• Check that you have sufficient disk space in the output folder
• Ensure you have write permissions to the destination folder
• Try converting files one at a time to isolate problematic files

Password Removal Not Working:
• Verify the password is correct (passwords are case-sensitive)
• Ensure all selected files use the same password
• Check that the Excel files are not corrupted
• Try processing files individually if batch processing fails

Slow Performance:
• Close unnecessary applications to free up system resources
• Process files in smaller batches (10-20 files at a time)
• Use local drives instead of network drives when possible

Error Messages:
• Check the log files in the "logs" folder for detailed error information
• Verify file permissions and accessibility
• Try restarting the tool and attempting the operation again

Settings Not Saving:
• Ensure you have write permissions to the tool's directory
• Check that the settings.json file is not read-only
• Verify sufficient disk space is available
• Run the tool as Administrator if permission issues persist

"@))
	
	$tabControl.Controls.Add((Create-HelpTab "About" @"

About Me:
Hi I'm Geoff Lu. I'm former FCC Reports Analyst, a passionate Software Developer and Automation Enthusiast, I specialize in creating practical tools that solve real-world business challenges. With extensive experience in PowerShell & VBScripting, Windows Forms development, and business process automation, I focus on delivering solutions that are both powerful and user-friendly.

My development philosophy centers on creating tools that are:
• Intuitive enough for non-technical users
• Robust enough for enterprise environments
• Flexible enough to adapt to various workflows
• Reliable enough for daily business operations

------------------------------------------------------------------------------------------------------------------------------------------------

Copyright & License:
© 2025 Geoff Lu. All rights & credits reserved.
This tool is provided as-is for educational and business use.
Redistribution and modification are permitted for personal and internal business use only.

"@))
	
	# Close Button
	$closeButton = New-Object System.Windows.Forms.Button
	$closeButton.Text = "Close"
	$closeButton.Size = New-Object System.Drawing.Size(150, 30)
	$closeButton.Location = New-Object System.Drawing.Point(320, 720)
	$closeButton.BackColor = $global:Config.Colors.Primary
	$closeButton.ForeColor = $global:Config.Colors.Secondary
	$closeButton.FlatStyle = "Standard"
	$closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$closeButton.DialogResult = "OK"
	$form.AcceptButton = $closeButton
	
	# Add Controls
	$form.Controls.Add($headerPanel)
	$form.Controls.Add($tabControl)
	$form.Controls.Add($closeButton)
	
	$form.ShowDialog() | Out-Null
}

function Show-SettingsDialog
{
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Application Tool Settings"
	$form.Size = New-Object System.Drawing.Size(550, 370) # Adjusted height
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = [System.Drawing.Color]::White
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$form.StartPosition = "CenterParent"
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Application Tool Settings"
				Font	  = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
				ForeColor = $global:Config.Colors.Primary
				Size	  = New-Object System.Drawing.Size(460, 25)
				Location  = New-Object System.Drawing.Point(20, 15)
			}))
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text = "Default Paths";
				Font = (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold));
				ForeColor = $global:Config.Colors.Text;
				Size = New-Object System.Drawing.Size(150, 20);
				Location = New-Object System.Drawing.Point(20, 50)
			}))
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Default Input:";
				ForeColor = $global:Config.Colors.Text;
				Size	  = New-Object System.Drawing.Size(100, 20);
				Location  = New-Object System.Drawing.Point(30, 80)
			}))
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Default Output:";
				ForeColor = $global:Config.Colors.Text;
				Size	  = New-Object System.Drawing.Size(100, 20);
				Location  = New-Object System.Drawing.Point(30, 110)
			}))
	
	$txtInput = New-Object System.Windows.Forms.TextBox -Property @{
		Size	  = New-Object System.Drawing.Size(300, 20);
		Location  = New-Object System.Drawing.Point(130, 78);
		BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230);
		Text	  = $global:Config.Settings.DefaultInputPath
	}
	
	$txtOutput = New-Object System.Windows.Forms.TextBox -Property @{
		Size	  = New-Object System.Drawing.Size(300, 20);
		Location  = New-Object System.Drawing.Point(130, 108);
		BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230);
		Text	  = $global:Config.Settings.DefaultOutputPath
	}
	
	$form.Controls.Add($txtInput); $form.Controls.Add($txtOutput)
	
	$btnBrowseInput = [StandardButton]::new();
	$btnBrowseInput.Text = "Browse...";
	$btnBrowseInput.Size = New-Object System.Drawing.Size(80, 25);
	$btnBrowseInput.Location = New-Object System.Drawing.Point(440, 75)
	
	$btnBrowseOutput = [StandardButton]::new();
	$btnBrowseOutput.Text = "Browse...";
	$btnBrowseOutput.Size = New-Object System.Drawing.Size(80, 25);
	$btnBrowseOutput.Location = New-Object System.Drawing.Point(440, 105)
	
	$form.Controls.Add($btnBrowseInput);
	$form.Controls.Add($btnBrowseOutput)
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text = "Application Settings";
				Font = (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold));
				ForeColor = $global:Config.Colors.Text;
				Size = New-Object System.Drawing.Size(150, 20);
				Location = New-Object System.Drawing.Point(20, 150)
			}))
	
	$chkRemember = New-Object System.Windows.Forms.CheckBox -Property @{
		Text	 = "Remember selected paths as defaults";
		Size	 = New-Object System.Drawing.Size(400, 20);
		Location = New-Object System.Drawing.Point(30, 180);
		Checked  = $global:Config.Settings.RememberPaths
	}
	
	$chkDelete = New-Object System.Windows.Forms.CheckBox -Property @{
		Text	 = "Automatically delete original files after processing";
		Size	 = New-Object System.Drawing.Size(400, 20);
		Location = New-Object System.Drawing.Point(30, 210);
		Checked  = $global:Config.Settings.AutoDeleteOriginals
	}
	
	$chkLogs = New-Object System.Windows.Forms.CheckBox -Property @{
		Text	 = "Show detailed logs in console";
		Size	 = New-Object System.Drawing.Size(400, 20);
		Location = New-Object System.Drawing.Point(30, 240);
		Checked  = $global:Config.Settings.ShowDetailedLogs
	}
	
	$chkSound = New-Object System.Windows.Forms.CheckBox -Property @{
		Text	 = "Enable sound notifications";
		Size	 = New-Object System.Drawing.Size(400, 20);
		Location = New-Object System.Drawing.Point(30, 270);
		Checked  = $global:Config.Settings.SoundNotifications
	}
	
	$form.Controls.Add($chkRemember);
	$form.Controls.Add($chkDelete);
	$form.Controls.Add($chkLogs);
	$form.Controls.Add($chkSound);
	
	$btnOK = [StandardButton]::new();
	$btnOK.Text = "Save";
	$btnOK.Size = New-Object System.Drawing.Size(80, 30);
	$btnOK.Location = New-Object System.Drawing.Point(350, 300); # Adjusted position
	$btnOK.DialogResult = "OK"
	
	$btnCancel = [StandardButton]::new();
	$btnCancel.Text = "Cancel";
	$btnCancel.Size = New-Object System.Drawing.Size(80, 30);
	$btnCancel.Location = New-Object System.Drawing.Point(440, 300); # Adjusted position
	$btnCancel.DialogResult = "Cancel"
	
	$form.Controls.Add($btnOK);
	$form.Controls.Add($btnCancel)
	$form.AcceptButton = $btnOK;
	$form.CancelButton = $btnCancel
	
	$btnBrowseInput.Add_Click({
			$d = New-Object System.Windows.Forms.FolderBrowserDialog
			$d.Description = "Select Default Input Folder"
			if (-not [string]::IsNullOrWhiteSpace($txtInput.Text)) { $d.SelectedPath = $txtInput.Text }
			if ($d.ShowDialog($form) -eq "OK") { $txtInput.Text = $d.SelectedPath }
		})
	
	$btnBrowseOutput.Add_Click({
			$d = New-Object System.Windows.Forms.FolderBrowserDialog
			$d.Description = "Select Default Output Folder"
			if (-not [string]::IsNullOrWhiteSpace($txtOutput.Text)) { $d.SelectedPath = $txtOutput.Text }
			if ($d.ShowDialog($form) -eq "OK") { $txtOutput.Text = $d.SelectedPath }
		})
	
	$btnOK.Add_Click({
			if ($txtInput.Text -and -not (Test-Path $txtInput.Text))
			{
				Show-MessageBox "Input path doesn't exist: $($txtInput.Text)" "Invalid Path" "Warning"; return
			}
			if ($txtOutput.Text -and -not (Test-Path $txtOutput.Text))
			{
				Show-MessageBox "Output path doesn't exist: $($txtOutput.Text)" "Invalid Path" "Warning"; return
			}
			
			$global:Config.Settings.DefaultInputPath = $txtInput.Text
			$global:Config.Settings.DefaultOutputPath = $txtOutput.Text
			$global:Config.Settings.RememberPaths = $chkRemember.Checked
			$global:Config.Settings.AutoDeleteOriginals = $chkDelete.Checked
			$global:Config.Settings.ShowDetailedLogs = $chkLogs.Checked
			$global:Config.Settings.SoundNotifications = $chkSound.Checked
			if (Save-Settings)
			{
				Play-Sound "Success";
				
				$form.Close()
			}
			else { Show-MessageBox "Failed to save settings." "Error" "Error" }
		})
	
	$form.ShowDialog() | Out-Null
}

function Get-InputFiles($Title = "Select files to process", $Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*")
{
	$ofd = New-Object System.Windows.Forms.OpenFileDialog
	$ofd.Multiselect = $true
	$ofd.Title = $Title
	$ofd.Filter = $Filter
	if (Path-Valid $global:Config.Settings.DefaultInputPath) { $ofd.InitialDirectory = $global:Config.Settings.DefaultInputPath }
	
	$result = $ofd.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileNames }
	return $null
}

function Get-OutputFolder($Title = "Select output folder")
{
	$d = New-Object System.Windows.Forms.FolderBrowserDialog
	$d.Description = $Title
	if (Path-Valid $global:Config.Settings.DefaultOutputPath) { $d.SelectedPath = $global:Config.Settings.DefaultOutputPath }
	
	$result = $d.ShowDialog()
	
	if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $d.SelectedPath }
	return $null
}

function Get-PasswordInput($Title = "Enter Password")
{
	$form = New-Object System.Windows.Forms.Form -Property @{
		Text		    = $Title
		Size		    = New-Object System.Drawing.Size(380, 205)
		StartPosition   = "CenterParent"
		FormBorderStyle = "FixedDialog"
		MaximizeBox	    = $false
		MinimizeBox	    = $false
		BackColor	    = $global:Config.Colors.Secondary
		Font		    = (New-Object System.Drawing.Font("Segoe UI", 9))
	}
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Please enter the password used to protect these files"
				ForeColor = $global:Config.Colors.Primary;
				Font	  = (New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold))
				Size	  = New-Object System.Drawing.Size(340, 30);
				Location  = New-Object System.Drawing.Point(20, 15)
				TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
			}))
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Password:";
				ForeColor = $global:Config.Colors.Text;
				Size	  = New-Object System.Drawing.Size(80, 20)
				Location  = New-Object System.Drawing.Point(70, 45)
			}))
	$txtPassword = New-Object System.Windows.Forms.TextBox -Property @{
		Size	  = New-Object System.Drawing.Size(240, 20);
		Location  = New-Object System.Drawing.Point(70, 65)
		BackColor = [System.Drawing.Color]::FromArgb(230, 230, 230);
		UseSystemPasswordChar = $false
	}
	$form.Controls.Add($txtPassword)
	$btnOK = [StandardButton]::new(); $btnOK.Text = "OK"; $btnOK.Size = New-Object System.Drawing.Size(100, 32); $btnOK.Location = New-Object System.Drawing.Point(140, 100); $btnOK.DialogResult = "OK"
	$form.Controls.Add($btnOK)
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "This tool can only remove passwords from files if they all share the same password."
				ForeColor = $global:Config.Colors.Text;
				Font	  = (New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic))
				Size	  = New-Object System.Drawing.Size(340, 40);
				Location  = New-Object System.Drawing.Point(20, 140)
				TextAlign = [System.Windows.Forms.HorizontalAlignment]::Center
			}))
	$form.AcceptButton = $btnOK
	if ($form.ShowDialog() -eq "OK") { return $txtPassword.Text }
	return $null
}

# --- Main UI ---

function Initialize-MainForm
{
	Load-Settings
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Excel Utility Tool v$($global:Config.Version) - Geoff Lu"
	$form.Size = New-Object System.Drawing.Size(415, 400)
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.BackColor = $global:Config.Colors.Secondary
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 8)
	$form.StartPosition = "CenterScreen"
	
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text = "Excel Utility Tool";
				Font = (New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold))
				ForeColor = $global:Config.Colors.Primary;
				Size = New-Object System.Drawing.Size(360, 25);
				Location = New-Object System.Drawing.Point(20, 20)
			}))
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "Version $($global:Config.Version)";
				Size	  = New-Object System.Drawing.Size(360, 20);
				Location  = New-Object System.Drawing.Point(20, 45)
				ForeColor = $global:Config.Colors.Text
			}))
	
	$script:btnConvert = [StandardButton]::new();
	$script:btnConvert.Text = "Convert to Excel (.xlsx)";
	$script:btnConvert.Size = New-Object System.Drawing.Size(360, 35);
	$script:btnConvert.Location = New-Object System.Drawing.Point(20, 80)
	
	$script:btnUnlock = [StandardButton]::new();
	$script:btnUnlock.Text = "Unlock Excel File";
	$script:btnUnlock.Size = New-Object System.Drawing.Size(360, 35);
	$script:btnUnlock.Location = New-Object System.Drawing.Point(20, 125)
	
	$script:btnSettings = [StandardButton]::new();
	$script:btnSettings.Text = "Application Tool Settings";
	$script:btnSettings.Size = New-Object System.Drawing.Size(360, 35);
	$script:btnSettings.Location = New-Object System.Drawing.Point(20, 170)
	
	# FIX: Add robust try/catch block to the event handler for diagnostics.
	$script:btnSettings.Add_Click({
			try
			{
				Show-SettingsDialog
				$form.CenterToScreen()
			}
			catch
			{
				$errorMessage = "An unexpected error occurred after closing the settings dialog: `n$($_.Exception.Message)"
				Write-Log "$errorMessage `nStack Trace: $($_.ScriptStackTrace)" "ERROR"
				Show-MessageBox $errorMessage "Internal Error" "Error"
			}
		})
	
	$btnHelp = [StandardButton]::new();
	$btnHelp.Text = "Help && Information";
	$btnHelp.Size = New-Object System.Drawing.Size(360, 35);
	$btnHelp.Location = New-Object System.Drawing.Point(20, 215)
	
	$form.Controls.AddRange(@($script:btnConvert, $script:btnUnlock, $script:btnSettings, $btnHelp))
	
	$script:status = New-Object System.Windows.Forms.Label -Property @{
		Text	  = "Ready!";
		Size	  = New-Object System.Drawing.Size(360, 20);
		Location  = New-Object System.Drawing.Point(20, 260);
		ForeColor = $global:Config.Colors.Text
	}
	
	$script:progress = New-Object System.Windows.Forms.ProgressBar -Property @{
		Size	  = New-Object System.Drawing.Size(360, 20)
		Location  = New-Object System.Drawing.Point(20, 280)
		Style	  = "Continuous"
		ForeColor = $global:Config.Colors.Primary
	}
	
	$form.Controls.Add($script:status);
	$form.Controls.Add($script:progress)
	$form.Controls.Add((New-Object System.Windows.Forms.Label -Property @{
				Text	  = "© 2025 Geoff Lu. All rights reserved.`nCreated June 20, 2025"
				Size	  = New-Object System.Drawing.Size(360, 40); Location = New-Object System.Drawing.Point(20, 320)
				ForeColor = $global:Config.Colors.Text; Font = (New-Object System.Drawing.Font("Segoe UI", 8))
				TextAlign = "MiddleCenter"
			}))
	
	$script:btnConvert.Add_Click({
			$excel = $null
			$workbook = $null
			
			try
			{
				# Get input files
				if ($null -ne $status) { $status.Text = "Selecting files..." }
				$files = Get-InputFiles "Select files to convert to .xlsx" "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|Excel files (*.xls)|*.xls|All files (*.*)|*.*"
				if (-not $files) { if ($null -ne $status) { $status.Text = "Ready" }; return }
				
				# Get output folder
				if ($null -ne $status) { $status.Text = "Selecting output folder..." }
				$outFolder = Get-OutputFolder "Select output folder for converted files"
				if (-not $outFolder) { if ($null -ne $status) { $status.Text = "Ready" }; return }
				
				# Disable buttons during processing
				$script:btnConvert.Enabled = $false
				$script:btnUnlock.Enabled = $false
				$script:btnSettings.Enabled = $false
				
				# Create Excel instance
				$excel = New-Object -ComObject Excel.Application
				$excel.Visible = $false
				$excel.DisplayAlerts = $false
				
				$totalFiles = $files.Count
				$currentFile = 0
				
				# Process each file
				foreach ($file in $files)
				{
					$currentFile++
					# FIX: Add defensive null checks for UI elements.
					if ($null -ne $progress) { $progress.Value = ($currentFile / $totalFiles) * 100 }
					if ($null -ne $status) { $status.Text = "Processing file ${currentFile} of ${totalFiles}: $([System.IO.Path]::GetFileName($file))" }
					[System.Windows.Forms.Application]::DoEvents()
					
					try
					{
						$fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
						$outputPath = Join-Path -Path $outFolder -ChildPath "$fileName.xlsx"
						
						if (Test-Path $outputPath)
						{
							Remove-Item $outputPath -Force
						}
						
						$workbook = $excel.Workbooks.Open($file)
						$workbook.SaveAs($outputPath, 51)
						$workbook.Close($false)
						$workbook = $null
						
						Write-Log "Successfully converted $file to $outputPath"
						
						if ($global:Config.Settings.AutoDeleteOriginals)
						{
							Remove-Item $file -Force
							Write-Log "Deleted original file: $file"
						}
					}
					catch
					{
						Write-Log "Error processing $file : $_" -Level "ERROR"
						Show-MessageBox "Failed to convert file: $([System.IO.Path]::GetFileName($file))`n`nError: $_" "Conversion Error" "Error"
						continue
					}
				}
				
				if ($null -ne $status) { $status.Text = "File(s) converted!" }
				Play-Sound "Success"
				Show-MessageBox "Successfully converted $currentFile file(s) to XLSX format." "Conversion Complete"
			}
			catch
			{
				if ($null -ne $status) { $status.Text = "Error occurred!" }
				$errorMessage = $_.Exception.Message
				Write-Log "Conversion error: $errorMessage `nStack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
				if ($errorMessage -like "*Excel*")
				{
					$errorMessage = "Could not start Microsoft Excel. Please ensure it is installed correctly.`n`nOriginal Error: $errorMessage"
				}
				Show-MessageBox "An error occurred during conversion:`n`n$errorMessage" "Error" "Error"
			}
			finally
			{
				if ($workbook)
				{
					$workbook.Close($false)
					[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
					$workbook = $null # FIX: Explicitly null out the variable after release.
				}
				if ($excel)
				{
					$excel.Quit()
					[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
					$excel = $null # FIX: Explicitly null out the variable after release.
				}
				[System.GC]::Collect()
				[System.GC]::WaitForPendingFinalizers()
				
				$script:btnConvert.Enabled = $true
				$script:btnUnlock.Enabled = $true
				$script:btnSettings.Enabled = $true
				if ($null -ne $progress) { $progress.Value = 0 }
			}
		})
	
	$script:btnUnlock.Add_Click({
			$excel = $null
			$workbook = $null
			
			try
			{
				# Get password from user
				$password = Get-PasswordInput "Enter Excel Password"
				if (-not $password)
				{
					if ($null -ne $status) { $status.Text = "Ready" }
					return
				}
				
				# Get input files
				if ($null -ne $status) { $status.Text = "Selecting files..." }
				$files = Get-InputFiles "Select password-protected Excel files" "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*"
				if (-not $files)
				{
					if ($null -ne $status) { $status.Text = "Ready" }
					return
				}
				
				# Get output folder
				if ($null -ne $status) { $status.Text = "Selecting output folder..." }
				$outFolder = Get-OutputFolder "Select output folder for unlocked files"
				if (-not $outFolder)
				{
					if ($null -ne $status) { $status.Text = "Ready" }
					return
				}
				
				# Disable buttons during processing
				$script:btnConvert.Enabled = $false
				$script:btnUnlock.Enabled = $false
				$script:btnSettings.Enabled = $false
				
				# Create Excel instance
				$excel = New-Object -ComObject Excel.Application
				$excel.Visible = $false
				$excel.DisplayAlerts = $false
				
				$totalFiles = $files.Count
				$currentFile = 0
				$successCount = 0
				$failCount = 0
				
				# Process each file
				foreach ($file in $files)
				{
					$currentFile++
					# FIX: Add defensive null checks for UI elements.
					if ($null -ne $progress) { $progress.Value = ($currentFile / $totalFiles) * 100 }
					if ($null -ne $status) { $status.Text = "Processing file ${currentFile} of ${totalFiles}: $([System.IO.Path]::GetFileName($file))" }
					[System.Windows.Forms.Application]::DoEvents()
					
					try
					{
						$fileName = [System.IO.Path]::GetFileNameWithoutExtension($file)
						$extension = [System.IO.Path]::GetExtension($file)
						$outputPath = Join-Path -Path $outFolder -ChildPath "$fileName$extension"
						
						if (Test-Path $outputPath)
						{
							Remove-Item $outputPath -Force
						}
						
						$workbook = $excel.Workbooks.Open($file, $null, $false, $null, $password)
						
						if ($workbook.ProtectStructure)
						{
							$workbook.Unprotect($password)
						}
						
						if ($workbook.ProtectWindows)
						{
							$workbook.UnprotectWindows($password)
						}
						
						foreach ($worksheet in $workbook.Worksheets)
						{
							if ($worksheet.ProtectContents)
							{
								$worksheet.Unprotect($password)
							}
						}
						
						$workbook.SaveAs($outputPath, $null, "", "", $false, $false, $null, $null, $null, $null, $null, $null)
						$workbook.Close($false)
						$workbook = $null
						
						$successCount++
						Write-Log "Successfully unlocked $file and saved to $outputPath"
						
						if ($global:Config.Settings.AutoDeleteOriginals)
						{
							Remove-Item $file -Force
							Write-Log "Deleted original file: $file"
						}
					}
					catch
					{
						$failCount++
						Write-Log "Error unlocking $file : $_" -Level "ERROR"
						
						if ($_.Exception.Message -like "*password*" -or $_.Exception.Message -like "*not valid*")
						{
							Show-MessageBox "Incorrect password for file: $([System.IO.Path]::GetFileName($file))`n`nThis file may have a different password or be protected in a way this tool can't handle." "Password Error" "Error"
						}
						else
						{
							Show-MessageBox "Failed to unlock file: $([System.IO.Path]::GetFileName($file))`n`nError: $_" "Unlock Error" "Error"
						}
						continue
					}
				}
				
				if ($null -ne $status) { $status.Text = "Unlock process completed" }
				Play-Sound "Success"
				
				$resultMessage = "Unlock process completed:`n`n"
				$resultMessage += "Successfully unlocked: $successCount file(s)`n"
				$resultMessage += "Failed to unlock: $failCount file(s)"
				
				if ($failCount -gt 0)
				{
					$resultMessage += "`n`nNote: Some files may have different passwords or stronger protection that this tool cannot remove."
					Show-MessageBox $resultMessage "Unlock Complete" "Warning"
				}
				else
				{
					Show-MessageBox $resultMessage "Unlock Complete" "Information"
				}
			}
			catch
			{
				if ($null -ne $status) { $status.Text = "Error occurred!" }
				$errorMessage = $_.Exception.Message
				Write-Log "Unlock error: $errorMessage `nStack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
				if ($errorMessage -like "*Excel*")
				{
					$errorMessage = "Could not start Microsoft Excel. Please ensure it is installed correctly.`n`nOriginal Error: $errorMessage"
				}
				Show-MessageBox "An error occurred during unlock process:`n`n$errorMessage" "Error" "Error"
			}
			finally
			{
				if ($workbook)
				{
					$workbook.Close($false)
					[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
					$workbook = $null # FIX: Explicitly null out the variable after release.
				}
				if ($excel)
				{
					$excel.Quit()
					[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
					$excel = $null # FIX: Explicitly null out the variable after release.
				}
				[System.GC]::Collect()
				[System.GC]::WaitForPendingFinalizers()
				
				$script:btnConvert.Enabled = $true
				$script:btnUnlock.Enabled = $true
				$script:btnSettings.Enabled = $true
				if ($null -ne $progress) { $progress.Value = 0 }
			}
		})
	
	# FIX: Add robust try/catch block to the event handler for diagnostics.
	$btnHelp.Add_Click({
			try
			{
				Show-HelpDialog
			}
			catch
			{
				$errorMessage = "An unexpected error occurred when trying to show the help dialog: `n$($_.Exception.Message)"
				Write-Log "$errorMessage `nStack Trace: $($_.ScriptStackTrace)" "ERROR"
				Show-MessageBox $errorMessage "Internal Error" "Error"
			}
		})
	
	# Return the form
	return $form
}

# --- Application Entry Point ---

if (-not (Test-Path $global:Config.LogPath))
{
	New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
}
try
{
	Write-Host "Starting Excel Utility Tool v$($global:Config.Version)..." -ForegroundColor Green
	$mainForm = Initialize-MainForm
	[void]$mainForm.ShowDialog()
}
catch
{
	# FIX: Log the full exception details for better startup debugging.
	$errorMessage = "A fatal error occurred on startup: $($_.Exception.ToString())"
	Write-Log $errorMessage "ERROR"
	Show-MessageBox "Failed to start the Excel Utility Tool:`n`n$($_.Exception.Message)`nPlease check the log files for details." "Startup Error" "Error"
}
finally
{
	Write-Log "Session ended"
}
