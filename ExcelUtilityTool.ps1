# ==============================================================================
# Excel Utility Tool - Version 6.0
# Author: Geoff Lu | geoffsmile@gmail.com  
# Rebuilt: July 14, 2025
# Major improvements: COM object management, error handling, reliability
# ==============================================================================

if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne [System.Threading.ApartmentState]::STA)
{
	Write-Log "Switching to STA mode for COM compatibility" "INFO" -Force
	$ps = [PowerShell]::Create()
	$ps.AddScript({
			. $PSScriptRoot\ExcelUtilityTool.ps1
		}) | Out-Null
	$ps.Runspace.ApartmentState = [System.Threading.ApartmentState]::STA
	$ps.Invoke()
	exit
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Validate script is saved to file
if (-not $PSScriptRoot)
{
	[System.Windows.Forms.MessageBox]::Show("This script must be saved to a file before running. Please save the script and try again.", "Startup Error", "OK", "Error")
	exit
}

# --- Enhanced Global Configuration ---
$global:Config = @{
	Version = "6.0"
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
		BatchSize		    = 10
		RetryAttempts	    = 3
		TimeoutSeconds	    = 300
	}
	SettingsPath = Join-Path $PSScriptRoot "settings.json"
	LogPath = Join-Path $PSScriptRoot "logs"
	State   = @{
		IsProcessing	 = $false
		CurrentOperation = ""
		ProcessedCount   = 0
		TotalCount	     = 0
		CancelRequested  = $false
	}
}

# --- Enhanced Logging System ---
function Write-Log
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[ValidateSet("INFO", "WARNING", "ERROR", "DEBUG", "SUCCESS")]
		[string]$Level = "INFO",
		[switch]$Force
	)
	
	try
	{
		# Ensure log directory exists
		if (-not (Test-Path $global:Config.LogPath))
		{
			New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
		}
		
		$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
		$logFile = Join-Path $global:Config.LogPath "excel_utility_$(Get-Date -Format 'yyyy-MM-dd').log"
		$logEntry = "[$timestamp] [$Level] $Message"
		
		# Write to file
		Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
		
		# Console output if enabled or forced
		if ($global:Config.Settings.ShowDetailedLogs -or $Force)
		{
			$color = switch ($Level)
			{
				"ERROR" { "Red" }
				"WARNING" { "Yellow" }
				"SUCCESS" { "Green" }
				"DEBUG" { "Cyan" }
				default { "Gray" }
			}
			Write-Host $logEntry -ForegroundColor $color
		}
	}
	catch
	{
		Write-Host "Log write failed: $($_.Exception.Message)" -ForegroundColor Red
	}
}

# --- Enhanced Settings Management ---
function Load-Settings
{
	try
	{
		if (Test-Path $global:Config.SettingsPath)
		{
			$jsonContent = Get-Content $global:Config.SettingsPath -Raw -ErrorAction Stop
			if (-not [string]::IsNullOrWhiteSpace($jsonContent))
			{
				$settings = $jsonContent | ConvertFrom-Json -ErrorAction Stop
				if ($null -ne $settings)
				{
					foreach ($key in $settings.PSObject.Properties.Name)
					{
						if ($global:Config.Settings.ContainsKey($key))
						{
							$global:Config.Settings[$key] = $settings.$key
						}
					}
					Write-Log "Settings loaded successfully"
					return $true
				}
			}
		}
		Write-Log "Using default settings" "INFO"
		return $true
	}
	catch
	{
		Write-Log "Failed to load settings: $($_.Exception.Message)" "ERROR"
		return $false
	}
}

function Save-Settings
{
	try
	{
		$settingsJson = $global:Config.Settings | ConvertTo-Json -Depth 3
		Set-Content -Path $global:Config.SettingsPath -Value $settingsJson -ErrorAction Stop
		Write-Log "Settings saved successfully"
		return $true
	}
	catch
	{
		Write-Log "Failed to save settings: $($_.Exception.Message)" "ERROR"
		return $false
	}
}

# --- Enhanced Validation Functions ---
function Test-FilePath
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Path,
		[switch]$MustExist
	)
	
	if ([string]::IsNullOrWhiteSpace($Path))
	{
		return $false
	}
	
	if ($MustExist)
	{
		return Test-Path $Path -ErrorAction SilentlyContinue
	}
	
	try
	{
		$parentPath = Split-Path $Path -Parent
		return Test-Path $parentPath -ErrorAction SilentlyContinue
	}
	catch
	{
		return $false
	}
}

function Test-FileAccess
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$FilePath,
		[ValidateSet("Read", "Write")]
		[string]$Access = "Read"
	)
	
	try
	{
		if ($Access -eq "Read")
		{
			$file = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
		}
		else
		{
			$file = [System.IO.File]::Open($FilePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write)
		}
		$file.Close()
		return $true
	}
	catch
	{
		return $false
	}
}

# --- Enhanced COM Object Management ---
class ExcelManager
{
	hidden [System.__ComObject]$Application
	hidden [System.__ComObject]$Workbook
	hidden [System.Collections.ArrayList]$ComObjects
	hidden [bool]$IsDisposed
	hidden [bool]$IsExcelAvailable
	
	ExcelManager()
	{
		$this.ComObjects = [System.Collections.ArrayList]::new()
		$this.IsDisposed = $false
		$this.IsExcelAvailable = $false
		$this.Initialize()
	}
	
	hidden [void]Initialize()
	{
		try
		{
			Write-Log "Attempting to initialize Excel COM object" "DEBUG"
			$this.Application = New-Object -ComObject Excel.Application -ErrorAction Stop
			$this.RegisterComObject($this.Application)
			
			$this.Application.Visible = $false
			$this.Application.DisplayAlerts = $false
			$this.Application.ScreenUpdating = $false
			$this.Application.EnableEvents = $false
			
			$this.IsExcelAvailable = $true
			Write-Log "Excel COM object initialized successfully" "SUCCESS"
		}
		catch
		{
			Write-Log "Failed to initialize Excel COM object: $($_.Exception.Message). Excel may not be installed or accessible." "ERROR"
			$this.IsExcelAvailable = $false
		}
	}
	
	hidden [void]RegisterComObject([System.__ComObject]$ComObject)
	{
		if ($null -ne $ComObject)
		{
			$this.ComObjects.Add($ComObject) | Out-Null
		}
	}
	
	[System.__ComObject]OpenWorkbook([string]$FilePath, [string]$Password = $null)
	{
		if (-not $this.IsExcelAvailable)
		{
			throw "Excel is not available for processing."
		}
		
		try
		{
			Write-Log "Opening workbook: $FilePath" "DEBUG"
			
			# Close any existing workbook first
			if ($null -ne $this.Workbook)
			{
				$this.CloseWorkbook()
			}
			
			# Use the simpler v5 approach for opening workbooks
			if ([string]::IsNullOrEmpty($Password))
			{
				# Open without password
				$this.Workbook = $this.Application.Workbooks.Open($FilePath)
			}
			else
			{
				# Open with password using v5 style parameters
				$this.Workbook = $this.Application.Workbooks.Open($FilePath, $null, $false, $null, $Password)
			}
			
			$this.RegisterComObject($this.Workbook)
			
			if ($null -eq $this.Workbook)
			{
				throw "Failed to open workbook: Workbook object is null"
			}
			
			Write-Log "Workbook opened successfully: $FilePath" "DEBUG"
			return $this.Workbook
		}
		catch
		{
			Write-Log "Failed to open workbook: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" "ERROR"
			throw
		}
	}
	
	[void]SaveWorkbook([string]$OutputPath, [int]$Format = 51)
	{
		if ($null -eq $this.Workbook)
		{
			throw "No workbook is currently open"
		}
		
		try
		{
			Write-Log "Saving workbook to: $OutputPath" "DEBUG"
			
			# Ensure the output directory exists
			$outputDir = [System.IO.Path]::GetDirectoryName($OutputPath)
			if (-not (Test-Path $outputDir))
			{
				New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
			}
			
			# Use v5 style SaveAs for standard conversion
			if ($Format -eq 51)
			{
				$this.Workbook.SaveAs($OutputPath, 51)
			}
			else
			{
				# For unlocked files, save without password
				$this.Workbook.SaveAs($OutputPath, $null, "", "", $false, $false, $null, $null, $null, $null, $null, $null)
			}
			
			Write-Log "Workbook saved successfully" "SUCCESS"
		}
		catch
		{
			Write-Log "Failed to save workbook: $($_.Exception.Message)" "ERROR"
			throw
		}
	}
	
	[void]UnlockWorkbook([string]$Password)
	{
		if ($null -eq $this.Workbook)
		{
			throw "No workbook is currently open"
		}
		
		try
		{
			Write-Log "Unlocking workbook protection" "DEBUG"
			
			# Check and unprotect workbook structure (v5 style)
			if ($this.Workbook.ProtectStructure)
			{
				try
				{
					$this.Workbook.Unprotect($Password)
					Write-Log "Workbook structure protection removed" "DEBUG"
				}
				catch
				{
					Write-Log ("Failed to unprotect workbook structure: " + $_.Exception.Message) "WARNING"
				}
			}
			
			# Check and unprotect windows (v5 style)
			if ($this.Workbook.ProtectWindows)
			{
				try
				{
					$this.Workbook.UnprotectWindows($Password)
					Write-Log "Workbook windows protection removed" "DEBUG"
				}
				catch
				{
					Write-Log ("Failed to unprotect windows: " + $_.Exception.Message) "WARNING"
				}
			}
			
			# Unprotect each worksheet (v5 style iteration)
			foreach ($worksheet in $this.Workbook.Worksheets)
			{
				$this.RegisterComObject($worksheet)
				
				try
				{
					if ($worksheet.ProtectContents)
					{
						$worksheet.Unprotect($Password)
						Write-Log ("Worksheet " + $worksheet.Name + " protection removed") "DEBUG"
					}
				}
				catch
				{
					Write-Log ("Failed to unprotect worksheet " + $worksheet.Name + ": " + $_.Exception.Message) "WARNING"
				}
			}
			
			Write-Log "Workbook unlocked successfully" "SUCCESS"
		}
		catch
		{
			Write-Log "Failed to unlock workbook: $($_.Exception.Message)" "ERROR"
			throw
		}
	}
	
	[void]CloseWorkbook()
	{
		if ($null -ne $this.Workbook)
		{
			try
			{
				Write-Log "Closing workbook" "DEBUG"
				$this.Workbook.Close($false)
			}
			catch
			{
				Write-Log "Error closing workbook: $($_.Exception.Message)" "WARNING"
			}
			finally
			{
				$this.Workbook = $null
			}
		}
	}
	
	[void]Dispose()
	{
		if ($this.IsDisposed)
		{
			return
		}
		
		try
		{
			Write-Log "Disposing Excel COM objects" "DEBUG"
			
			# Close workbook first
			$this.CloseWorkbook()
			
			# Quit Excel application
			if ($null -ne $this.Application)
			{
				try
				{
					$this.Application.Quit()
				}
				catch
				{
					Write-Log "Error quitting Excel application: $($_.Exception.Message)" "WARNING"
				}
			}
			
			# Release COM objects in reverse order
			$this.ComObjects.Reverse()
			foreach ($comObject in $this.ComObjects)
			{
				if ($null -ne $comObject)
				{
					try
					{
						[System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
					}
					catch
					{
						Write-Log "Error releasing COM object: $($_.Exception.Message)" "WARNING"
					}
				}
			}
			
			$this.ComObjects.Clear()
			$this.Application = $null
			$this.Workbook = $null
			
			# Force garbage collection (v5 style)
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
			
			Write-Log "Excel COM objects disposed successfully" "SUCCESS"
		}
		catch
		{
			Write-Log "Error during COM object disposal: $($_.Exception.Message)" "ERROR"
		}
		finally
		{
			$this.IsDisposed = $true
		}
	}
}

# --- Enhanced File Processing ---
function Invoke-FileConversion
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$FilePaths,
		[Parameter(Mandatory = $true)]
		[string]$OutputFolder
	)
	
	$excelManager = $null
	$processedCount = 0
	$failedCount = 0
	$results = @()
	
	try
	{
		Write-Log "Starting file conversion process" "INFO"
		$global:Config.State.IsProcessing = $true
		$global:Config.State.CurrentOperation = "Converting files"
		$global:Config.State.ProcessedCount = 0
		$global:Config.State.TotalCount = $FilePaths.Count
		
		$excelManager = [ExcelManager]::new()
		if (-not $excelManager.IsExcelAvailable)
		{
			throw "Excel is not installed or accessible. Please ensure Microsoft Excel is installed and you have the necessary permissions."
		}
		
		$totalFiles = $FilePaths.Count
		for ($currentFile = 0; $currentFile -lt $totalFiles; $currentFile++)
		{
			if ($global:Config.State.CancelRequested)
			{
				Write-Log "Conversion cancelled by user" "WARNING"
				break
			}
			
			$filePath = $FilePaths[$currentFile]
			$result = Invoke-SingleFileConversion -FilePath $filePath -OutputFolder $OutputFolder -ExcelManager $excelManager
			$results += $result
			
			if ($result.Success)
			{
				$processedCount++
			}
			else
			{
				$failedCount++
			}
			
			$global:Config.State.ProcessedCount = $currentFile + 1
			Update-ProgressDisplay
			[System.Windows.Forms.Application]::DoEvents()
		}
		
		Write-Log "Conversion completed. Processed: $processedCount, Failed: $failedCount" "INFO"
		return $results
	}
	catch
	{
		Write-Log "Critical error in file conversion: $($_.Exception.Message)" "ERROR"
		throw
	}
	finally
	{
		if ($null -ne $excelManager)
		{
			$excelManager.Dispose()
		}
		$global:Config.State.IsProcessing = $false
		Reset-ProgressDisplay
	}
}

function Invoke-SingleFileConversion
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$FilePath,
		[Parameter(Mandatory = $true)]
		[string]$OutputFolder,
		[Parameter(Mandatory = $true)]
		[ExcelManager]$ExcelManager
	)
	
	$result = @{
		FilePath	   = $FilePath
		Success	       = $false
		OutputPath	   = ""
		ErrorMessage   = ""
		ProcessingTime = 0
	}
	
	$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
	
	try
	{
		Write-Log "Converting file: $([System.IO.Path]::GetFileName($FilePath))" "DEBUG"
		
		$fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
		$outputPath = Join-Path $OutputFolder "$fileName.xlsx"
		
		if (Test-Path $outputPath)
		{
			Remove-Item $outputPath -Force -ErrorAction Stop
		}
		
		$openedWorkbook = $ExcelManager.OpenWorkbook($FilePath, $null) # Pass $null for password
		$ExcelManager.SaveWorkbook($outputPath, 51)
		$ExcelManager.CloseWorkbook()
		
		$result.Success = $true
		$result.OutputPath = $outputPath
		
		if ($global:Config.Settings.AutoDeleteOriginals -and $result.Success)
		{
			Remove-Item $FilePath -Force -ErrorAction Stop
			Write-Log "Deleted original file: $FilePath" "DEBUG"
		}
		
		Write-Log "File converted successfully: $([System.IO.Path]::GetFileName($FilePath))" "SUCCESS"
	}
	catch
	{
		$result.ErrorMessage = $_.Exception.Message
		Write-Log "Error processing $FilePath : $($_.Exception.Message)" "ERROR"
		Show-MessageBox "Failed to convert file: $([System.IO.Path]::GetFileName($FilePath))`n`nError: $($_.Exception.Message)" "Conversion Error" "Error"
	}
	finally
	{
		$stopwatch.Stop()
		$result.ProcessingTime = $stopwatch.ElapsedMilliseconds
	}
	
	return $result
}

# --- Enhanced Password Unlock Function ---
function Invoke-FileUnlock
{
	param (
		[Parameter(Mandatory = $true)]
		[string[]]$FilePaths,
		[Parameter(Mandatory = $true)]
		[string]$OutputFolder,
		[Parameter(Mandatory = $true)]
		[string]$Password
	)
	
	$excelManager = $null
	$successCount = 0
	$failCount = 0
	$results = @()
	
	try
	{
		Write-Log "Starting file unlock process" "INFO"
		$global:Config.State.IsProcessing = $true
		$global:Config.State.CurrentOperation = "Unlocking files"
		$global:Config.State.ProcessedCount = 0
		$global:Config.State.TotalCount = $FilePaths.Count
		
		$excelManager = [ExcelManager]::new()
		if (-not $excelManager.IsExcelAvailable)
		{
			throw "Excel is not installed or accessible. Please ensure Microsoft Excel is installed and you have the necessary permissions."
		}
		
		$totalFiles = $FilePaths.Count
		for ($currentFile = 0; $currentFile -lt $totalFiles; $currentFile++)
		{
			if ($global:Config.State.CancelRequested)
			{
				Write-Log "Unlock cancelled by user" "WARNING"
				break
			}
			
			$filePath = $FilePaths[$currentFile]
			$result = Invoke-SingleFileUnlock -FilePath $filePath -OutputFolder $OutputFolder -ExcelManager $excelManager -Password $Password
			$results += $result
			
			if ($result.Success)
			{
				$successCount++
			}
			else
			{
				$failCount++
			}
			
			$global:Config.State.ProcessedCount = $currentFile + 1
			Update-ProgressDisplay
			[System.Windows.Forms.Application]::DoEvents()
		}
		
		Write-Log "Unlock completed. Success: $successCount, Failures: $failCount" "INFO"
		return $results
	}
	catch
	{
		Write-Log "Critical error in file unlock: $($_.Exception.Message)" "ERROR"
		throw
	}
	finally
	{
		if ($null -ne $excelManager)
		{
			$excelManager.Dispose()
		}
		$global:Config.State.IsProcessing = $false
		Reset-ProgressDisplay
	}
}

function Invoke-SingleFileUnlock
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$FilePath,
		[Parameter(Mandatory = $true)]
		[string]$OutputFolder,
		[Parameter(Mandatory = $true)]
		[ExcelManager]$ExcelManager,
		[Parameter(Mandatory = $true)]
		[string]$Password
	)
	
	$result = @{
		FilePath	   = $FilePath
		Success	       = $false
		OutputPath	   = ""
		ErrorMessage   = ""
		ProcessingTime = 0
	}
	
	$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
	
	try
	{
		Write-Log "Unlocking file: $([System.IO.Path]::GetFileName($FilePath))" "DEBUG"
		
		$fileName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
		$extension = [System.IO.Path]::GetExtension($FilePath)
		$outputPath = Join-Path $OutputFolder "$fileName$extension"
		
		if (Test-Path $outputPath)
		{
			Remove-Item $outputPath -Force -ErrorAction Stop
		}
		
		# Open the workbook with password
		$openedWorkbook = $ExcelManager.OpenWorkbook($FilePath, $Password)
		
		# Unlock the workbook
		$ExcelManager.UnlockWorkbook($Password)
		
		# Save the workbook WITHOUT password protection (use 0 to indicate unlocked save)
		$ExcelManager.SaveWorkbook($outputPath, 0)
		
		# Close the workbook
		$ExcelManager.CloseWorkbook()
		
		$result.Success = $true
		$result.OutputPath = $outputPath
		
		if ($global:Config.Settings.AutoDeleteOriginals -and $result.Success)
		{
			Remove-Item $FilePath -Force -ErrorAction Stop
			Write-Log "Deleted original file: $FilePath" "DEBUG"
		}
		
		Write-Log "Successfully unlocked $FilePath and saved to $outputPath" "SUCCESS"
	}
	catch
	{
		$result.ErrorMessage = $_.Exception.Message
		Write-Log "Error unlocking $FilePath : $($_.Exception.Message)" "ERROR"
		
		if ($_.Exception.Message -like "*password*" -or $_.Exception.Message -like "*not valid*")
		{
			$result.ErrorMessage = "Incorrect password or file protection not supported"
			Show-MessageBox "Incorrect password for file: $([System.IO.Path]::GetFileName($FilePath))`n`nThis file may have a different password or be protected in a way this tool can't handle." "Password Error" "Error"
		}
		else
		{
			Show-MessageBox "Failed to unlock file: $([System.IO.Path]::GetFileName($FilePath))`n`nError: $($_.Exception.Message)" "Unlock Error" "Error"
		}
	}
	finally
	{
		$stopwatch.Stop()
		$result.ProcessingTime = $stopwatch.ElapsedMilliseconds
	}
	
	return $result
}

# --- Enhanced UI Helper Functions ---
function Play-NotificationSound
{
	param (
		[ValidateSet("Success", "Error", "Warning", "Info")]
		[string]$Type = "Info"
	)
	
	if (-not $global:Config.Settings.SoundNotifications)
	{
		return
	}
	
	try
	{
		switch ($Type)
		{
			"Success" { [System.Media.SystemSounds]::Asterisk.Play() }
			"Error" { [System.Media.SystemSounds]::Hand.Play() }
			"Warning" { [System.Media.SystemSounds]::Exclamation.Play() }
			"Info" { [System.Media.SystemSounds]::Beep.Play() }
		}
	}
	catch
	{
		Write-Log "Failed to play sound: $($_.Exception.Message)" "WARNING"
	}
}

function Show-MessageBox
{
	param (
		[Parameter(Mandatory = $true)]
		[string]$Message,
		[string]$Title = "Excel Utility Tool",
		[ValidateSet("Information", "Warning", "Error", "Question")]
		[string]$Type = "Information",
		[ValidateSet("OK", "OKCancel", "YesNo", "YesNoCancel")]
		[string]$Buttons = "OK"
	)
	
	$icon = switch ($Type)
	{
		"Error" { Play-NotificationSound "Error"; [System.Windows.Forms.MessageBoxIcon]::Error }
		"Warning" { Play-NotificationSound "Warning"; [System.Windows.Forms.MessageBoxIcon]::Warning }
		"Question" { Play-NotificationSound "Info"; [System.Windows.Forms.MessageBoxIcon]::Question }
		default { Play-NotificationSound "Info"; [System.Windows.Forms.MessageBoxIcon]::Information }
	}
	
	$buttonType = switch ($Buttons)
	{
		"OKCancel" { [System.Windows.Forms.MessageBoxButtons]::OKCancel }
		"YesNo" { [System.Windows.Forms.MessageBoxButtons]::YesNo }
		"YesNoCancel" { [System.Windows.Forms.MessageBoxButtons]::YesNoCancel }
		default { [System.Windows.Forms.MessageBoxButtons]::OK }
	}
	
	return [System.Windows.Forms.MessageBox]::Show($Message, $Title, $buttonType, $icon)
}

function Update-ProgressDisplay
{
	if ($null -ne $script:progress -and $null -ne $script:status)
	{
		$state = $global:Config.State
		
		if ($state.TotalCount -gt 0)
		{
			$percentage = [Math]::Round(($state.ProcessedCount / $state.TotalCount) * 100, 1)
			$script:progress.Value = [Math]::Min($percentage, 100)
			$script:status.Text = "$($state.CurrentOperation): $($state.ProcessedCount) of $($state.TotalCount) ($percentage%)"
		}
		
		[System.Windows.Forms.Application]::DoEvents()
	}
}

function Reset-ProgressDisplay
{
	if ($null -ne $script:progress -and $null -ne $script:status)
	{
		$script:progress.Value = 0
		$script:status.Text = "Ready"
		[System.Windows.Forms.Application]::DoEvents()
	}
}

# --- Enhanced Dialog Functions ---
function Get-InputFiles
{
	param (
		[string]$Title = "Select files to process",
		[string]$Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*"
	)
	
	try
	{
		$ofd = New-Object System.Windows.Forms.OpenFileDialog
		$ofd.Multiselect = $true
		$ofd.Title = $Title
		$ofd.Filter = $Filter
		$ofd.CheckFileExists = $true
		$ofd.CheckPathExists = $true
		
		# Determine initial directory based on default input path
		if ([string]::IsNullOrWhiteSpace($global:Config.Settings.DefaultInputPath) -or
			-not (Test-FilePath $global:Config.Settings.DefaultInputPath -MustExist))
		{
			$ofd.InitialDirectory = [Environment]::GetFolderPath("Desktop")
			Write-Log "No valid default input path found. Prompting for manual selection." "INFO"
		}
		else
		{
			$ofd.InitialDirectory = $global:Config.Settings.DefaultInputPath
			Write-Log "Using default input path: $($global:Config.Settings.DefaultInputPath)" "INFO"
		}
		
		$result = $ofd.ShowDialog()
		if ($result -eq [System.Windows.Forms.DialogResult]::OK)
		{
			return $ofd.FileNames
		}
		
		return $null
	}
	catch
	{
		Write-Log "Error in file selection dialog: $($_.Exception.Message)" "ERROR"
		Show-MessageBox "Error selecting files: $($_.Exception.Message)" -Type "Error"
		return $null
	}
}

function Get-OutputFolder
{
	param (
		[string]$Title = "Select output folder"
	)
	
	try
	{
		$fbd = New-Object System.Windows.Forms.FolderBrowserDialog
		$fbd.Description = $Title
		$fbd.ShowNewFolderButton = $true
		
		# Determine initial directory based on default output path
		if ([string]::IsNullOrWhiteSpace($global:Config.Settings.DefaultOutputPath) -or
			-not (Test-FilePath $global:Config.Settings.DefaultOutputPath -MustExist))
		{
			$fbd.SelectedPath = [Environment]::GetFolderPath("Desktop")
			Write-Log "No valid default output path found. Prompting for manual selection." "INFO"
		}
		else
		{
			$fbd.SelectedPath = $global:Config.Settings.DefaultOutputPath
			Write-Log "Using default output path: $($global:Config.Settings.DefaultOutputPath)" "INFO"
		}
		
		$result = $fbd.ShowDialog()
		if ($result -eq [System.Windows.Forms.DialogResult]::OK)
		{
			return $fbd.SelectedPath
		}
		
		return $null
	}
	catch
	{
		Write-Log "Error in folder selection dialog: $($_.Exception.Message)" "ERROR"
		Show-MessageBox "Error selecting folder: $($_.Exception.Message)" -Type "Error"
		return $null
	}
}

function Get-PasswordInput
{
	param (
		[string]$Title = "Enter Password"
	)
	
	$form = New-Object System.Windows.Forms.Form
	$form.Text = $Title
	$form.Size = New-Object System.Drawing.Size(315, 200)
	$form.StartPosition = "CenterParent"
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = $global:Config.Colors.Secondary
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# Password label
	$lblPassword = New-Object System.Windows.Forms.Label
	$lblPassword.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
	$lblPassword.Text = "Password of protected file(s):"
	$lblPassword.ForeColor = $global:Config.Colors.Primary
	$lblPassword.Size = New-Object System.Drawing.Size(300, 15)
	$lblPassword.Location = New-Object System.Drawing.Point(60, 15)
	$form.Controls.Add($lblPassword)
	
	# Password textbox
	$txtPassword = New-Object System.Windows.Forms.TextBox
	$txtPassword.Size = New-Object System.Drawing.Size(240, 20)
	$txtPassword.Location = New-Object System.Drawing.Point(30, 40)
	$txtPassword.UseSystemPasswordChar = $true
	$txtPassword.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
	$form.Controls.Add($txtPassword)
	
	# OK button
	$btnOK = New-Object System.Windows.Forms.Button
	$btnOK.Text = "OK"
	$btnOK.Size = New-Object System.Drawing.Size(80, 32)
	$btnOK.Location = New-Object System.Drawing.Point(60, 75)
	$btnOK.BackColor = $global:Config.Colors.Primary
	$btnOK.ForeColor = $global:Config.Colors.Secondary
	$btnOK.FlatStyle = "Flat"
	$btnOK.DialogResult = "OK"
	$form.Controls.Add($btnOK)
	
	# Cancel button
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = New-Object System.Drawing.Size(80, 32)
	$btnCancel.Location = New-Object System.Drawing.Point(150, 75)
	$btnCancel.BackColor = $global:Config.Colors.Text
	$btnCancel.ForeColor = $global:Config.Colors.Secondary
	$btnCancel.FlatStyle = "Flat"
	$btnCancel.DialogResult = "Cancel"
	$form.Controls.Add($btnCancel)
	
	# Info label
	$lblInfo = New-Object System.Windows.Forms.Label
	$lblInfo.Text = "Note: This tool can only remove passwords from files if they all share the same password."
	$lblInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
	$lblInfo.ForeColor = $global:Config.Colors.Text
	$lblInfo.Size = New-Object System.Drawing.Size(240, 45)
	$lblInfo.Location = New-Object System.Drawing.Point(30, 115)
	$lblInfo.TextAlign = "MiddleCenter"
	$form.Controls.Add($lblInfo)
	
	$form.AcceptButton = $btnOK
	$form.CancelButton = $btnCancel
	
	# Focus on password field
	$txtPassword.Select()
	
	try
	{
		$result = $form.ShowDialog()
		if ($result -eq "OK")
		{
			return $txtPassword.Text
		}
		return $null
	}
	catch
	{
		Write-Log "Error in password dialog: $($_.Exception.Message)" "ERROR"
		return $null
	}
	finally
	{
		$form.Dispose()
	}
}

# --- Settings Dialog (Enhanced) ---
function Show-SettingsDialog
{
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Excel Utility Tool Settings"
	$form.Size = New-Object System.Drawing.Size(580, 370)
	$form.FormBorderStyle = "FixedDialog"
	$form.MaximizeBox = $false
	$form.MinimizeBox = $false
	$form.BackColor = $global:Config.Colors.Secondary
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	$form.StartPosition = "CenterParent"
	
	# Header
	$lblPanel = New-Object System.Windows.Forms.Panel
	$lblPanel.Size = New-Object System.Drawing.Size(545, 40)
	$lblPanel.Location = New-Object System.Drawing.Point(10, 10)
	$lblPanel.BackColor = $global:Config.Colors.Primary
	
	$lblLabel = New-Object System.Windows.Forms.Label
	$lblLabel.Text = "APPLICATION SETTINGS"
	$lblLabel.ForeColor = $global:Config.Colors.Secondary
	$lblLabel.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
	$lblLabel.Size = New-Object System.Drawing.Size(515, 30)
	$lblLabel.Location = New-Object System.Drawing.Point(10, 5)
	$lblLabel.TextAlign = "MiddleLeft"
	$lblPanel.Controls.Add($lblLabel)
	
	$form.Controls.Add($lblPanel)
	
	# Paths section
	$lblPaths = New-Object System.Windows.Forms.Label
	$lblPaths.Text = "Default Paths:"
	$lblPaths.ForeColor = $global:Config.Colors.Text
	$lblPaths.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$lblPaths.ForeColor = $global:Config.Colors.Text
	$lblPaths.Size = New-Object System.Drawing.Size(200, 20)
	$lblPaths.Location = New-Object System.Drawing.Point(20, 60)
	$form.Controls.Add($lblPaths)
	
	# Input path
	$lblInput = New-Object System.Windows.Forms.Label
	$lblInput.Text = "Source Folder:"
	$lblInput.ForeColor = $global:Config.Colors.Text
	$lblInput.Size = New-Object System.Drawing.Size(85, 15)
	$lblInput.Location = New-Object System.Drawing.Point(30, 90)
	$form.Controls.Add($lblInput)
	
	$txtInput = New-Object System.Windows.Forms.TextBox
	$txtInput.Size = New-Object System.Drawing.Size(355, 20)
	$txtInput.Location = New-Object System.Drawing.Point(115, 85)
	$txtInput.Text = $global:Config.Settings.DefaultInputPath
	$form.Controls.Add($txtInput)
	
	$btnBrowseInput = New-Object System.Windows.Forms.Button
	$btnBrowseInput.Text = "Browse"
	$btnBrowseInput.Size = New-Object System.Drawing.Size(70, 25)
	$btnBrowseInput.Location = New-Object System.Drawing.Point(475, 85)
	$btnBrowseInput.BackColor = $global:Config.Colors.Primary
	$btnBrowseInput.ForeColor = $global:Config.Colors.Secondary
	$btnBrowseInput.FlatStyle = "Flat"
	$form.Controls.Add($btnBrowseInput)
	
	# Output path
	$lblOutput = New-Object System.Windows.Forms.Label
	$lblOutput.Text = "Output Folder:"
	$lblOutput.ForeColor = $global:Config.Colors.Text
	$lblOutput.Size = New-Object System.Drawing.Size(85, 15)
	$lblOutput.Location = New-Object System.Drawing.Point(30, 125)
	$form.Controls.Add($lblOutput)
	
	$txtOutput = New-Object System.Windows.Forms.TextBox
	$txtOutput.Size = New-Object System.Drawing.Size(355, 20)
	$txtOutput.Location = New-Object System.Drawing.Point(115, 120)
	$txtOutput.Text = $global:Config.Settings.DefaultOutputPath
	$form.Controls.Add($txtOutput)
	
	$btnBrowseOutput = New-Object System.Windows.Forms.Button
	$btnBrowseOutput.Text = "Browse"
	$btnBrowseOutput.Size = New-Object System.Drawing.Size(70, 25)
	$btnBrowseOutput.Location = New-Object System.Drawing.Point(475, 121)
	$btnBrowseOutput.BackColor = $global:Config.Colors.Primary
	$btnBrowseOutput.ForeColor = $global:Config.Colors.Secondary
	$btnBrowseOutput.FlatStyle = "Flat"
	$form.Controls.Add($btnBrowseOutput)
	
	# Processing section
	$lblProcessing = New-Object System.Windows.Forms.Label
	$lblProcessing.Text = "Processing Options:"
	$lblProcessing.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
	$lblProcessing.ForeColor = $global:Config.Colors.Text
	$lblProcessing.Size = New-Object System.Drawing.Size(200, 20)
	$lblProcessing.Location = New-Object System.Drawing.Point(20, 160)
	$form.Controls.Add($lblProcessing)
	
	# Batch size
	$lblBatchSize = New-Object System.Windows.Forms.Label
	$lblBatchSize.Text = "Batch Size:"
	$lblBatchSize.ForeColor = $global:Config.Colors.Text
	$lblBatchSize.Size = New-Object System.Drawing.Size(90, 15)
	$lblBatchSize.Location = New-Object System.Drawing.Point(370, 190)
	$form.Controls.Add($lblBatchSize)
	
	$numBatchSize = New-Object System.Windows.Forms.NumericUpDown
	$numBatchSize.Size = New-Object System.Drawing.Size(70, 20)
	$numBatchSize.Location = New-Object System.Drawing.Point(475, 188)
	$numBatchSize.Minimum = 1
	$numBatchSize.Maximum = 100
	$numBatchSize.Value = $global:Config.Settings.BatchSize
	$form.Controls.Add($numBatchSize)
	
	# Retry attempts
	$lblRetryAttempts = New-Object System.Windows.Forms.Label
	$lblRetryAttempts.Text = "Retry Attempts:"
	$lblRetryAttempts.ForeColor = $global:Config.Colors.Text
	$lblRetryAttempts.Size = New-Object System.Drawing.Size(90, 15)
	$lblRetryAttempts.Location = New-Object System.Drawing.Point(370, 220)
	$form.Controls.Add($lblRetryAttempts)
	
	$numRetryAttempts = New-Object System.Windows.Forms.NumericUpDown
	$numRetryAttempts.Size = New-Object System.Drawing.Size(70, 20)
	$numRetryAttempts.Location = New-Object System.Drawing.Point(475, 218)
	$numRetryAttempts.Minimum = 1
	$numRetryAttempts.Maximum = 10
	$numRetryAttempts.Value = $global:Config.Settings.RetryAttempts
	$form.Controls.Add($numRetryAttempts)
	
	# Checkboxes
	$chkRememberPaths = New-Object System.Windows.Forms.CheckBox
	$chkRememberPaths.Text = "Remember selected paths"
	$chkRememberPaths.ForeColor = $global:Config.Colors.Text
	$chkRememberPaths.Size = New-Object System.Drawing.Size(250, 20)
	$chkRememberPaths.Location = New-Object System.Drawing.Point(30, 190)
	$chkRememberPaths.Checked = $global:Config.Settings.RememberPaths
	$form.Controls.Add($chkRememberPaths)
	
	$chkAutoDelete = New-Object System.Windows.Forms.CheckBox
	$chkAutoDelete.Text = "Automatically delete original files after processing"
	$chkAutoDelete.ForeColor = $global:Config.Colors.Text
	$chkAutoDelete.Size = New-Object System.Drawing.Size(350, 20)
	$chkAutoDelete.Location = New-Object System.Drawing.Point(30, 220)
	$chkAutoDelete.Checked = $global:Config.Settings.AutoDeleteOriginals
	$form.Controls.Add($chkAutoDelete)
	
	$chkShowLogs = New-Object System.Windows.Forms.CheckBox
	$chkShowLogs.Text = "Show detailed logs in console"
	$chkShowLogs.ForeColor = $global:Config.Colors.Text
	$chkShowLogs.Size = New-Object System.Drawing.Size(250, 20)
	$chkShowLogs.Location = New-Object System.Drawing.Point(30, 250)
	$chkShowLogs.Checked = $global:Config.Settings.ShowDetailedLogs
	$form.Controls.Add($chkShowLogs)
	
	$chkSoundNotifications = New-Object System.Windows.Forms.CheckBox
	$chkSoundNotifications.Text = "Enable sound notifications"
	$chkSoundNotifications.ForeColor = $global:Config.Colors.Text
	$chkSoundNotifications.Size = New-Object System.Drawing.Size(250, 20)
	$chkSoundNotifications.Location = New-Object System.Drawing.Point(30, 280)
	$chkSoundNotifications.Checked = $global:Config.Settings.SoundNotifications
	$form.Controls.Add($chkSoundNotifications)
	
	# Buttons
	$btnSave = New-Object System.Windows.Forms.Button
	$btnSave.Text = "Save"
	$btnSave.Size = New-Object System.Drawing.Size(70, 30)
	$btnSave.Location = New-Object System.Drawing.Point(395, 280)
	$btnSave.BackColor = $global:Config.Colors.Primary
	$btnSave.ForeColor = $global:Config.Colors.Secondary
	$btnSave.FlatStyle = "Flat"
	$btnSave.DialogResult = "OK"
	$form.Controls.Add($btnSave)
	
	$btnCancel = New-Object System.Windows.Forms.Button
	$btnCancel.Text = "Cancel"
	$btnCancel.Size = New-Object System.Drawing.Size(70, 30)
	$btnCancel.Location = New-Object System.Drawing.Point(475, 280)
	$btnCancel.BackColor = $global:Config.Colors.Text
	$btnCancel.ForeColor = $global:Config.Colors.Secondary
	$btnCancel.FlatStyle = "Flat"
	$btnCancel.DialogResult = "Cancel"
	$form.Controls.Add($btnCancel)
	
	# Event handlers
	$btnBrowseInput.Add_Click({
			$folder = Get-OutputFolder -Title "Select Default Input Folder"
			if ($folder)
			{
				$txtInput.Text = $folder
			}
		})
	
	$btnBrowseOutput.Add_Click({
			$folder = Get-OutputFolder -Title "Select Default Output Folder"
			if ($folder)
			{
				$txtOutput.Text = $folder
			}
		})
	
	$btnSave.Add_Click({
			try
			{
				# Validate paths
				if ($txtInput.Text -and -not (Test-Path $txtInput.Text))
				{
					Show-MessageBox "Input path does not exist: $($txtInput.Text)" -Type "Warning"
					return
				}
				
				if ($txtOutput.Text -and -not (Test-Path $txtOutput.Text))
				{
					Show-MessageBox "Output path does not exist: $($txtOutput.Text)" -Type "Warning"
					return
				}
				
				# Save settings
				$global:Config.Settings.DefaultInputPath = $txtInput.Text
				$global:Config.Settings.DefaultOutputPath = $txtOutput.Text
				$global:Config.Settings.BatchSize = $numBatchSize.Value
				$global:Config.Settings.RetryAttempts = $numRetryAttempts.Value
				$global:Config.Settings.RememberPaths = $chkRememberPaths.Checked
				$global:Config.Settings.AutoDeleteOriginals = $chkAutoDelete.Checked
				$global:Config.Settings.ShowDetailedLogs = $chkShowLogs.Checked
				$global:Config.Settings.SoundNotifications = $chkSoundNotifications.Checked
				
				if (Save-Settings)
				{
					Play-NotificationSound "Success"
					$form.Close()
				}
				else
				{
					Show-MessageBox "Failed to save settings. Please check the log for details." -Type "Error"
				}
			}
			catch
			{
				Write-Log "Error saving settings: $($_.Exception.Message)" "ERROR"
				Show-MessageBox "Error saving settings: $($_.Exception.Message)" -Type "Error"
			}
		})
	
	$form.AcceptButton = $btnSave
	$form.CancelButton = $btnCancel
	
	try
	{
		$form.ShowDialog() | Out-Null
	}
	catch
	{
		Write-Log "Error in settings dialog: $($_.Exception.Message)" "ERROR"
	}
	finally
	{
		$form.Dispose()
	}
}

# --- Enhanced Help Dialog ---
function Show-HelpDialog
{
	$form = New-Object System.Windows.Forms.Form
	$form.Text = "Excel Utility Tool - Help & Information"
	$form.Size = New-Object System.Drawing.Size(900, 700)
	$form.MinimumSize = $form.Size
	$form.MaximumSize = $form.Size
	$form.FormBorderStyle = "FixedDialog"
	$form.StartPosition = "CenterParent"
	$form.BackColor = $global:Config.Colors.Secondary
	$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# Header
	$headerPanel = New-Object System.Windows.Forms.Panel
	$headerPanel.Size = New-Object System.Drawing.Size(865, 50)
	$headerPanel.Location = New-Object System.Drawing.Point(10, 10)
	$headerPanel.BackColor = $global:Config.Colors.Primary
	
	$headerLabel = New-Object System.Windows.Forms.Label
	$headerLabel.Text = "EXCEL UTILITY TOOL - VERSION 6.0"
	$headerLabel.ForeColor = $global:Config.Colors.Secondary
	$headerLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
	$headerLabel.Size = New-Object System.Drawing.Size(825, 30)
	$headerLabel.Location = New-Object System.Drawing.Point(20, 10)
	$headerLabel.TextAlign = "MiddleLeft"
	$headerPanel.Controls.Add($headerLabel)
	
	$form.Controls.Add($headerPanel)
	
	# Tab control
	$tabControl = New-Object System.Windows.Forms.TabControl
	$tabControl.Size = New-Object System.Drawing.Size(865, 580)
	$tabControl.Location = New-Object System.Drawing.Point(10, 70)
	$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9)
	
	# Overview tab
	$tabOverview = New-Object System.Windows.Forms.TabPage
	$tabOverview.Text = "Overview"
	$tabOverview.BackColor = $global:Config.Colors.Secondary
	
	$txtOverview = New-Object System.Windows.Forms.RichTextBox
	$txtOverview.Text = @"
Excel Utility Tool - Version 6.0
Author: Geoff Lu | geoffsmile@gmail.com
Created: June 20, 2025
Last Modified: July 14, 2025

================================================================================
MAJOR VERSION 6.0 IMPROVEMENTS
================================================================================

This version represents a complete rebuild of the Excel Utility Tool with focus on:

✓ RELIABILITY & STABILITY
  • Proper COM object management with automatic cleanup
  • Comprehensive error handling with retry logic
  • Memory leak prevention and resource management
  • Graceful failure recovery

✓ ENHANCED PERFORMANCE
  • Batch processing for large file sets
  • Non-blocking UI during operations
  • Optimized file validation and processing
  • Improved progress reporting

✓ BETTER USER EXPERIENCE
  • Enhanced error messages and logging
  • Configurable processing options
  • Sound notifications for completion
  • Detailed operation feedback

✓ ROBUST ARCHITECTURE
  • Separation of UI and business logic
  • Proper state management
  • Comprehensive input validation
  • Thread-safe operations

================================================================================
CORE FEATURES
================================================================================

• Convert CSV, TXT, and XLS files to modern XLSX format
• Unlock password-protected Excel files
• Batch processing with configurable batch sizes
• Automatic retry logic for failed operations
• Comprehensive logging with multiple levels
• Customizable settings with validation
• Progress tracking and cancellation support
• Sound notifications for operation completion
• Automatic cleanup of temporary resources
"@
	$txtOverview.Dock = "Fill"
	$txtOverview.ReadOnly = $true
	$txtOverview.Font = New-Object System.Drawing.Font("Consolas", 9)
	$txtOverview.BackColor = $global:Config.Colors.Secondary
	$txtOverview.ForeColor = $global:Config.Colors.Text
	$tabOverview.Controls.Add($txtOverview)
	
	# Instructions tab
	$tabInstructions = New-Object System.Windows.Forms.TabPage
	$tabInstructions.Text = "Instructions"
	$tabInstructions.BackColor = $global:Config.Colors.Secondary
	
	$txtInstructions = New-Object System.Windows.Forms.RichTextBox
	$txtInstructions.Text = @"
EXCEL UTILITY TOOL - USAGE INSTRUCTIONS
================================================================================

CONVERTING FILES TO EXCEL (.xlsx)
================================================================================

1. Click "Convert to Excel (.xlsx)" button
2. Select files using the file browser:
   • Supported formats: CSV, TXT, XLS
   • Multiple files can be selected simultaneously
   • Files are validated before processing
3. Choose destination folder for converted files
4. Monitor progress through the progress bar and status display
5. Review completion message for processing summary

Features:
• Batch processing for improved performance
• Automatic retry on temporary failures
• File validation before processing
• Progress tracking with cancellation support
• Optional deletion of original files after conversion

================================================================================
UNLOCKING PROTECTED EXCEL FILES
================================================================================

1. Click "Unlock Excel File" button
2. Enter the password in the password dialog
   • Password is case-sensitive
   • All selected files must use the same password
3. Select the password-protected Excel files
4. Choose destination folder for unlocked files
5. Monitor progress and review completion summary

Features:
• Supports both workbook and worksheet protection
• Batch processing for multiple files
• Automatic retry on authentication failures
• Preservation of original file format
• Comprehensive error reporting

================================================================================
CONFIGURING SETTINGS
================================================================================

1. Click "Application Tool Settings" button
2. Configure options in the settings dialog:

   DEFAULT PATHS:
   • Default Input Folder: Starting location for file selection
   • Default Output Folder: Default save location

   PROCESSING OPTIONS:
   • Batch Size: Number of files processed simultaneously (1-100)
   • Retry Attempts: Number of retry attempts for failed operations (1-10)

   PREFERENCES:
   • Remember selected paths as defaults
   • Automatically delete original files after processing
   • Show detailed logs in console
   • Enable sound notifications

3. Click "Save" to apply changes
4. Settings are automatically loaded on startup

================================================================================
TROUBLESHOOTING & TIPS
================================================================================

• Check log files in the 'logs' folder for detailed error information
• Ensure you have sufficient permissions for file operations
• For large batches, consider reducing batch size if memory issues occur
• Use the retry settings to handle temporary network or file access issues
• Monitor the progress bar for operation status and completion
• Sound notifications can be disabled in settings if needed

================================================================================
BEST PRACTICES
================================================================================

• Always backup important files before processing
• Test with a small batch before processing large numbers of files
• Ensure adequate disk space in the output folder
• Keep the tool updated for optimal performance and reliability
• Use descriptive folder names for better organization
• Check log files regularly for any processing issues
"@
	$txtInstructions.Dock = "Fill"
	$txtInstructions.ReadOnly = $true
	$txtInstructions.Font = New-Object System.Drawing.Font("Consolas", 9)
	$txtInstructions.BackColor = $global:Config.Colors.Secondary
	$txtInstructions.ForeColor = $global:Config.Colors.Text
	$tabInstructions.Controls.Add($txtInstructions)
	
	# About tab
	$tabAbout = New-Object System.Windows.Forms.TabPage
	$tabAbout.Text = "About"
	$tabAbout.BackColor = $global:Config.Colors.Secondary
	
	$txtAbout = New-Object System.Windows.Forms.RichTextBox
	$txtAbout.Text = @"
EXCEL UTILITY TOOL - VERSION 6.0
================================================================================

ABOUT THE DEVELOPER
================================================================================

Geoff Lu - Entry Software Developer & Automation Specialist
Email: geoffsmile@gmail.com

A former FCC Reports Analyst with extensive experience in:
• PowerShell scripting and automation
• Windows Forms application development
• Business process optimization
• Enterprise software solutions

Development Philosophy:
• User-friendly interfaces for non-technical users
• Robust architecture for enterprise environments
• Flexible design for various workflows
• Reliable performance for daily operations

================================================================================
VERSION HISTORY
================================================================================

Version 6.0 - Major Rebuild
• Complete architecture overhaul
• Enhanced COM object management
• Improved error handling and logging
• Better performance and reliability
• Enhanced user experience

Version 5.0
• Sound notifications
• Always-on-top functionality

Version 4.0
• Help system introduction
• Performance optimizations

Version 3.0
• Enhanced user interface
• Settings management

Version 2.0
• Password removal capabilities

Version 1.0
• Initial release
• Basic file conversion functionality

================================================================================
TECHNICAL SPECIFICATIONS
================================================================================

• Built with: PowerShell 5.1+ and .NET Framework
• UI Framework: Windows Forms
• Supported OS: Windows 10/11
• Architecture: x86 and x64 compatible
• Dependencies: Microsoft Office Excel (recommended)
• Memory: 4GB RAM recommended
• Storage: 100MB free disk space

================================================================================
COPYRIGHT & LICENSE
================================================================================

© 2025 Geoff Lu. All rights reserved.

This software is provided as-is for educational and business use.
Redistribution and modification are permitted for personal and 
internal business use only.

No warranties are provided. Use at your own risk.

================================================================================
SUPPORT & FEEDBACK
================================================================================

For support, questions, or feedback:
• Email: geoffsmile@gmail.com
• Check log files for detailed error information
• Review the troubleshooting section for common issues

Thank you for using the Excel Utility Tool!
"@
	$txtAbout.Dock = "Fill"
	$txtAbout.ReadOnly = $true
	$txtAbout.Font = New-Object System.Drawing.Font("Consolas", 9)
	$txtAbout.BackColor = $global:Config.Colors.Secondary
	$txtAbout.ForeColor = $global:Config.Colors.Text
	$tabAbout.Controls.Add($txtAbout)
	
	# Add tabs to control
	$tabControl.Controls.AddRange(@($tabOverview, $tabInstructions, $tabAbout))
	$form.Controls.Add($tabControl)
	
	# Close button
	$btnClose = New-Object System.Windows.Forms.Button
	$btnClose.Text = "Close"
	$btnClose.Size = New-Object System.Drawing.Size(100, 35)
	$btnClose.Location = New-Object System.Drawing.Point(400, 660)
	$btnClose.BackColor = $global:Config.Colors.Primary
	$btnClose.ForeColor = $global:Config.Colors.Secondary
	$btnClose.FlatStyle = "Flat"
	$btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
	$btnClose.DialogResult = "OK"
	$form.Controls.Add($btnClose)
	
	$form.AcceptButton = $btnClose
	
	try
	{
		$form.ShowDialog() | Out-Null
	}
	catch
	{
		Write-Log "Error in help dialog: $($_.Exception.Message)" "ERROR"
	}
	finally
	{
		$form.Dispose()
	}
}

# --- Enhanced Main Form ---
function Initialize-MainForm
{
	try
	{
		# Load settings
		if (-not (Load-Settings))
		{
			Write-Log "Failed to load settings, using defaults" "WARNING"
		}
		
		# Create main form
		$form = New-Object System.Windows.Forms.Form
		$form.Text = "Excel Utility Tool v$($global:Config.Version) - Geoff Lu"
		$form.Size = New-Object System.Drawing.Size(450, 400)
		$form.FormBorderStyle = "FixedDialog"
		$form.MaximizeBox = $false
		$form.BackColor = $global:Config.Colors.Secondary
		$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
		$form.StartPosition = "CenterScreen"
		
		# Application title
		$lblTitle = New-Object System.Windows.Forms.Label
		$lblTitle.Text = "Excel Utility Tool"
		$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
		$lblTitle.ForeColor = $global:Config.Colors.Primary
		$lblTitle.Size = New-Object System.Drawing.Size(400, 30)
		$lblTitle.Location = New-Object System.Drawing.Point(25, 20)
		$lblTitle.TextAlign = "MiddleCenter"
		$form.Controls.Add($lblTitle)
		
		# Version label
		$lblVersion = New-Object System.Windows.Forms.Label
		$lblVersion.Text = "Version $($global:Config.Version) by Geoff Lu"
		$lblVersion.Font = New-Object System.Drawing.Font("Segoe UI", 9)
		$lblVersion.ForeColor = $global:Config.Colors.Text
		$lblVersion.Size = New-Object System.Drawing.Size(400, 20)
		$lblVersion.Location = New-Object System.Drawing.Point(25, 50)
		$lblVersion.TextAlign = "MiddleCenter"
		$form.Controls.Add($lblVersion)
		
		# Convert button
		$script:btnConvert = New-Object System.Windows.Forms.Button
		$script:btnConvert.Text = "Convert to Excel (.xlsx)"
		$script:btnConvert.Size = New-Object System.Drawing.Size(380, 35)
		$script:btnConvert.Location = New-Object System.Drawing.Point(25, 90)
		$script:btnConvert.BackColor = $global:Config.Colors.Primary
		$script:btnConvert.ForeColor = $global:Config.Colors.Secondary
		$script:btnConvert.FlatStyle = "Flat"
		$script:btnConvert.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
		$form.Controls.Add($script:btnConvert)
		
		# Unlock button
		$script:btnUnlock = New-Object System.Windows.Forms.Button
		$script:btnUnlock.Text = "Unlock Excel Files"
		$script:btnUnlock.Size = New-Object System.Drawing.Size(380, 35)
		$script:btnUnlock.Location = New-Object System.Drawing.Point(25, 135)
		$script:btnUnlock.BackColor = $global:Config.Colors.Primary
		$script:btnUnlock.ForeColor = $global:Config.Colors.Secondary
		$script:btnUnlock.FlatStyle = "Flat"
		$script:btnUnlock.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
		$form.Controls.Add($script:btnUnlock)
		
		# Settings button
		$script:btnSettings = New-Object System.Windows.Forms.Button
		$script:btnSettings.Text = "Application Settings"
		$script:btnSettings.Size = New-Object System.Drawing.Size(380, 35)
		$script:btnSettings.Location = New-Object System.Drawing.Point(25, 180)
		$script:btnSettings.BackColor = $global:Config.Colors.Primary
		$script:btnSettings.ForeColor = $global:Config.Colors.Secondary
		$script:btnSettings.FlatStyle = "Flat"
		$script:btnSettings.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
		$form.Controls.Add($script:btnSettings)
		
		# Help button
		$btnHelp = New-Object System.Windows.Forms.Button
		$btnHelp.Text = "Help & Information"
		$btnHelp.Size = New-Object System.Drawing.Size(380, 35)
		$btnHelp.Location = New-Object System.Drawing.Point(25, 225)
		$btnHelp.BackColor = $global:Config.Colors.Primary
		$btnHelp.ForeColor = $global:Config.Colors.Secondary
		$btnHelp.FlatStyle = "Flat"
		$btnHelp.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
		$form.Controls.Add($btnHelp)
		
		# Status label
		$script:status = New-Object System.Windows.Forms.Label
		$script:status.Text = "Ready"
		$script:status.Size = New-Object System.Drawing.Size(380, 20)
		$script:status.Location = New-Object System.Drawing.Point(25, 270)
		$script:status.ForeColor = $global:Config.Colors.Text
		$script:status.Font = New-Object System.Drawing.Font("Segoe UI", 9)
		$form.Controls.Add($script:status)
		
		# Progress bar
		$script:progress = New-Object System.Windows.Forms.ProgressBar
		$script:progress.Size = New-Object System.Drawing.Size(380, 20)
		$script:progress.Location = New-Object System.Drawing.Point(25, 290)
		$script:progress.Style = "Continuous"
		$script:progress.ForeColor = $global:Config.Colors.Primary
		$form.Controls.Add($script:progress)
		
		# Copyright label
		$lblCopyright = New-Object System.Windows.Forms.Label
		$lblCopyright.Text = "© 2025 Geoff Lu. All rights reserved.`nVersion 6.0 - Created on June 20, 2025"
		$lblCopyright.Size = New-Object System.Drawing.Size(380, 40)
		$lblCopyright.Location = New-Object System.Drawing.Point(25, 320)
		$lblCopyright.ForeColor = $global:Config.Colors.Text
		$lblCopyright.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
		$lblCopyright.TextAlign = "MiddleCenter"
		$form.Controls.Add($lblCopyright)
		
		# Event handlers
		$script:btnConvert.Add_Click({
				try
				{
					if ($global:Config.State.IsProcessing)
					{
						Show-MessageBox "Another operation is currently in progress. Please wait for it to complete." -Type "Warning"
						return
					}
					
					Write-Log "Starting file conversion process" "INFO"
					
					$script:status.Text = "Selecting files to convert..."
					[System.Windows.Forms.Application]::DoEvents()
					
					$files = Get-InputFiles -Title "Select files to convert to Excel (.xlsx)" -Filter "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|Excel files (*.xls)|*.xls|All files (*.*)|*.*"
					if (-not $files)
					{
						$script:status.Text = "Ready"
						return
					}
					
					$script:status.Text = "Selecting output folder..."
					[System.Windows.Forms.Application]::DoEvents()
					
					$outputFolder = Get-OutputFolder -Title "Select output folder for converted files"
					if (-not $outputFolder)
					{
						$script:status.Text = "Ready"
						return
					}
					
					$script:btnConvert.Enabled = $false
					$script:btnUnlock.Enabled = $false
					$script:btnSettings.Enabled = $false
					
					$results = Invoke-FileConversion -FilePaths $files -OutputFolder $outputFolder
					
					$successCount = ($results | Where-Object { $_.Success }).Count
					$failureCount = $results.Count - $successCount
					
					if ($failureCount -eq 0)
					{
						$message = "Successfully converted $successCount file(s) to Excel format."
						$script:status.Text = "File(s) converted!"
						Show-MessageBox $message -Type "Information"
						Play-NotificationSound "Success"
					}
					else
					{
						$message = "Conversion completed with mixed results:`n`nSuccessfully converted: $successCount file(s)`nFailed to convert: $failureCount file(s)`n`nCheck the log files for detailed error information."
						$script:status.Text = "Conversion completed with errors!"
						Show-MessageBox $message -Type "Warning"
						Play-NotificationSound "Warning"
					}
					
					Write-Log "File conversion completed. Success: $successCount, Failures: $failureCount" "INFO"
				}
				catch
				{
					$errorMessage = "An error occurred during file conversion: $($_.Exception.Message)"
					Write-Log $errorMessage "ERROR"
					$script:status.Text = "Error occurred!"
					Show-MessageBox $errorMessage -Type "Error"
					Play-NotificationSound "Error"
				}
				finally
				{
					$script:btnConvert.Enabled = $true
					$script:btnUnlock.Enabled = $true
					$script:btnSettings.Enabled = $true
					Reset-ProgressDisplay
				}
			})
		
		$script:btnUnlock.Add_Click({
				try
				{
					if ($global:Config.State.IsProcessing)
					{
						Show-MessageBox "Another operation is currently in progress. Please wait for it to complete." -Type "Warning"
						return
					}
					
					Write-Log "Starting file unlock process" "INFO"
					
					$password = Get-PasswordInput -Title "Enter Excel Password"
					if (-not $password)
					{
						$script:status.Text = "Ready"
						return
					}
					
					$script:status.Text = "Selecting files..."
					[System.Windows.Forms.Application]::DoEvents()
					
					$files = Get-InputFiles -Title "Select password-protected Excel files" -Filter "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx|All files (*.*)|*.*"
					if (-not $files)
					{
						$script:status.Text = "Ready"
						return
					}
					
					$script:status.Text = "Selecting output folder..."
					[System.Windows.Forms.Application]::DoEvents()
					
					$outputFolder = Get-OutputFolder -Title "Select output folder for unlocked files"
					if (-not $outputFolder)
					{
						$script:status.Text = "Ready"
						return
					}
					
					$script:btnConvert.Enabled = $false
					$script:btnUnlock.Enabled = $false
					$script:btnSettings.Enabled = $false
					
					$results = Invoke-FileUnlock -FilePaths $files -OutputFolder $outputFolder -Password $password
					
					$successCount = ($results | Where-Object { $_.Success }).Count
					$failureCount = $results.Count - $successCount
					
					$script:status.Text = "Unlock process completed"
					$message = "Unlock process completed:`n`nSuccessfully unlocked: $successCount file(s)`nFailed to unlock: $failureCount file(s)"
					if ($failureCount -gt 0)
					{
						$message += "`n`nNote: Some files may have different passwords or stronger protection that this tool cannot remove."
						Show-MessageBox $message -Type "Warning"
					}
					else
					{
						Show-MessageBox $message -Type "Information"
					}
					Play-NotificationSound "Success"
					
					Write-Log "File unlock completed. Success: $successCount, Failures: $failureCount" "INFO"
				}
				catch
				{
					$errorMessage = "An error occurred during file unlock: $($_.Exception.Message)"
					Write-Log $errorMessage "ERROR"
					$script:status.Text = "Error occurred!"
					Show-MessageBox $errorMessage -Type "Error"
					Play-NotificationSound "Error"
				}
				finally
				{
					$script:btnConvert.Enabled = $true
					$script:btnUnlock.Enabled = $true
					$script:btnSettings.Enabled = $true
					Reset-ProgressDisplay
				}
			})
		
		$script:btnSettings.Add_Click({
				try
				{
					if ($global:Config.State.IsProcessing)
					{
						Show-MessageBox "Another operation is currently in progress. Please wait for it to complete." -Type "Warning"
						return
					}
					Write-Log "Opening settings dialog" "DEBUG"
					Show-SettingsDialog
					Write-Log "Settings dialog closed" "DEBUG"
				}
				catch
				{
					Write-Log "Error opening settings dialog: $($_.Exception.Message)" "ERROR"
					Show-MessageBox "Error opening settings: $($_.Exception.Message)" -Type "Error"
				}
			})
		
		$btnHelp.Add_Click({
				try
				{
					if ($global:Config.State.IsProcessing)
					{
						Show-MessageBox "Another operation is currently in progress. Please wait for it to complete." -Type "Warning"
						return
					}
					Write-Log "Opening help dialog" "DEBUG"
					Show-HelpDialog
					Write-Log "Help dialog closed" "DEBUG"
				}
				catch
				{
					Write-Log "Error opening help dialog: $($_.Exception.Message)" "ERROR"
					Show-MessageBox "Error opening help: $($_.Exception.Message)" -Type "Error"
				}
			})
		
		# Return the form
		return $form
	}
	catch
	{
		$errorMessage = "A fatal error occurred while initializing the main form: $($_.Exception.Message)"
		Write-Log $errorMessage "ERROR"
		Show-MessageBox "Failed to initialize the main form:`n`n$($_.Exception.Message)`nPlease check the log files for details." "Startup Error" "Error"
		return $null
	}
}

# --- Application Entry Point ---
try
{
	# Ensure log directory exists
	if (-not (Test-Path $global:Config.LogPath))
	{
		New-Item -Path $global:Config.LogPath -ItemType Directory -Force | Out-Null
	}
	
	Write-Log "Starting Excel Utility Tool v$($global:Config.Version)" "INFO" -Force
	Write-Host "Excel Utility Tool v$($global:Config.Version) - Enhanced Edition" -ForegroundColor Green
	Write-Host "Initializing application..." -ForegroundColor Gray
	
	# Initialize and show main form
	$mainForm = Initialize-MainForm
	
	Write-Log "Application initialized successfully" "SUCCESS"
	Write-Host "Application ready!" -ForegroundColor Green
	
	# Show the main form
	[void]$mainForm.ShowDialog()
}
catch
{
	$errorMessage = "Critical error during application startup: $($_.Exception.Message)"
	Write-Log $errorMessage "ERROR" -Force
	Write-Host $errorMessage -ForegroundColor Red
	
	# Show error dialog
	try
	{
		[System.Windows.Forms.MessageBox]::Show(
			"Failed to start Excel Utility Tool:`n`n$($_.Exception.Message)`n`nPlease check the log files for detailed information.",
			"Application Startup Error",
			"OK",
			"Error"
		)
	}
	catch
	{
		Write-Host "Failed to show error dialog: $($_.Exception.Message)" -ForegroundColor Red
	}
}
finally
{
	Write-Log "Application session ended" "INFO"
}