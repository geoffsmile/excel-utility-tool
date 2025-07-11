Dim objExcel, objFSO, objFile, objWorkbook, objWorksheet
Dim csvFolder, destFolder, newFileName, fileCount, toolFilePath
Dim objShell

' Set the source and destination folder paths
csvFolder = "C:\Users\PH10035990\Downloads\"
destFolder = "C:\Users\PH10035990\OneDrive - R1\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\04 - Raw Data Pull\"
toolFilePath = "C:\Users\PH10035990\OneDrive - R1\MIS Sharepoint\01 - FCC Files\Volume Pull - Hourly Inventory\TOOL - R1Access_Inv_Gen - GEOFF.xlsm"

' Ensure folder paths end with a backslash
If Right(csvFolder, 1) <> "\" Then csvFolder = csvFolder & "\"
If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"

' Create FileSystemObject and Shell instances
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")

' Check if the source folder exists
If Not objFSO.FolderExists(csvFolder) Then
    MsgBox "The source folder '" & csvFolder & "' was not found.", 48, "Folder Not Found"
    ' Prompt the user to select the source folder
    Set objFolder = objShell.BrowseForFolder(0, "Select Source Folder Containing CSV Files", 0)
    If objFolder Is Nothing Then
        MsgBox "No folder selected. Exiting script.", 48, "Operation Cancelled"
        WScript.Quit
    End If
    csvFolder = objFolder.Self.Path
    If Right(csvFolder, 1) <> "\" Then csvFolder = csvFolder & "\"
End If

' Check if the destination folder exists
If Not objFSO.FolderExists(destFolder) Then
    MsgBox "The destination folder '" & destFolder & "' was not found.", 48, "Folder Not Found"
    ' Prompt the user to select the destination folder
    Set objFolder = objShell.BrowseForFolder(0, "Select Destination Folder for XLSX Files", 0)
    If objFolder Is Nothing Then
        MsgBox "No folder selected. Exiting script.", 48, "Operation Cancelled"
        WScript.Quit
    End If
    destFolder = objFolder.Self.Path
    If Right(destFolder, 1) <> "\" Then destFolder = destFolder & "\"
End If

' Initialize file counter
fileCount = 0

' Create Excel instance
Set objExcel = CreateObject("Excel.Application")

' Disable Excel alerts (e.g., overwrite prompts)
objExcel.DisplayAlerts = False

' Process all CSV files in the source folder
For Each objFile In objFSO.GetFolder(csvFolder).Files
   ' Check if the file is a CSV and contains "ExportWorklist" in its name
   If LCase(objFSO.GetExtensionName(objFile.Name)) = "csv" And InStr(objFile.Name, "ExportWorklist") > 0 Then
       fileCount = fileCount + 1
       
       ' Open the CSV file
       Set objWorkbook = objExcel.Workbooks.Open(objFile.Path)
       Set objWorksheet = objWorkbook.Sheets(1)
       
       ' Define the new file name in the destination folder
       newFileName = objFSO.BuildPath(destFolder, Replace(objFile.Name, ".csv", ".xlsx"))
       
       ' Save as XLSX with the same file name in the destination folder
       ' The SaveAs method will automatically overwrite the file if it already exists
       objWorkbook.SaveAs newFileName, 51 ' xlOpenXMLWorkbook
       
       ' Close the workbook
       objWorkbook.Close False
       
       ' Delete the original CSV file
       objFile.Delete
   End If
Next

' Cleanup
objExcel.DisplayAlerts = True ' Re-enable alerts
' Check if exactly 9 files were processed
If fileCount = 9 Then
    ' Open the tool file after successful conversion
    If objFSO.FileExists(toolFilePath) Then
        Set objWorkbook = objExcel.Workbooks.Open(toolFilePath)
        objExcel.Visible = True ' Make Excel visible to the user
        objExcel.ActiveWindow.WindowState = -4137 ' Maximize the Excel window
        
        ' Run the macro "Macro5"
        On Error Resume Next ' Ignore errors if the macro does not exist
        objExcel.Run "Macro5"
        On Error GoTo 0 ' Reset error handling
    Else
        MsgBox "The tool file '" & toolFilePath & "' was not found.", 48, "Tool File Not Found"
        objExcel.Quit ' Quit Excel if the tool file is not found
    End If
ElseIf fileCount > 0 Then
    ' Display a message if fewer than 9 files were processed
    MsgBox "Only " & fileCount & " file(s) containing 'ExportWorklist' were found and processed. Expected 9 files.", 48, "Incomplete Conversion"
    objExcel.Quit ' Quit Excel if fewer than 9 files were processed
Else
    ' Display a message if no files were processed
    MsgBox "No CSV files containing 'ExportWorklist' found in the specified folder.", 48, "No Files Found"
    objExcel.Quit ' Quit Excel if no files were processed
End If

' Release objects
Set objExcel = Nothing
Set objFSO = Nothing
Set objFile = Nothing
Set objWorkbook = Nothing
Set objWorksheet = Nothing
Set objShell = Nothing
Set objFolder = Nothing