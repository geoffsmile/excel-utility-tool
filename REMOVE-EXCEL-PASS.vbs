' Script to Remove Password Protection from Excel Files using Excel Interop
'
' Author: GEOFF LU 
' LAST UPDATE: APRIL 17, 2025
' VERSION No: v1.9 (Silent VBS Version - Same Directory)

Option Explicit

' Configuration
Dim password
password = "1"

' Constants for Excel
Const xlOpenXMLWorkbook = 51 ' Excel Open XML Workbook format

' Create objects
Dim fso, shell, currentPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Get the directory where the script is running
currentPath = fso.GetParentFolderName(WScript.ScriptFullName)

' Kill Excel processes function
Sub KillExcelProcesses()
    On Error Resume Next
    shell.Run "taskkill /f /im excel.exe", 0, True
    WScript.Sleep 2000 ' Wait 2 seconds
    On Error Goto 0
End Sub

' Start with clean slate
Call KillExcelProcesses()

' Get all Excel files in current folder
Dim folder, file
Dim processedFiles, totalFiles
Set folder = fso.GetFolder(currentPath)
processedFiles = 0
totalFiles = 0

' Count Excel files first
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Path)) = "xlsx" Then
        totalFiles = totalFiles + 1
    End If
Next

' Get all Excel files in current folder
Dim folder, file
Set folder = fso.GetFolder(currentPath)

' Process each Excel file in the current directory
For Each file In folder.Files
    If LCase(fso.GetExtensionName(file.Path)) = "xlsx" Then
        Dim excel, workbook
        Dim sourcePath, unprotectedPath
        
        sourcePath = file.Path
        unprotectedPath = currentPath & "\unprotected_" & fso.GetFileName(sourcePath)
        
        ' Create new Excel instance
        On Error Resume Next
        Set excel = CreateObject("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        ' Try to open with password
        Set workbook = excel.Workbooks.Open(sourcePath, 0, False, 5, password)
        
        If Err.Number = 0 Then
            ' Save without password protection
            workbook.SaveAs unprotectedPath, xlOpenXMLWorkbook, "", "", False
            workbook.Close True
            
            ' Delete the original password-protected file
            fso.DeleteFile sourcePath
            
            ' Rename the unprotected file to the original name
            fso.MoveFile unprotectedPath, sourcePath
        End If
        
        ' Clean up
        If Not workbook Is Nothing Then
            On Error Resume Next
            workbook.Close False
            Set workbook = Nothing
            On Error Goto 0
        End If
        
        If Not excel Is Nothing Then
            On Error Resume Next
            excel.Quit
            Set excel = Nothing
            On Error Goto 0
        End If
        
        ' Kill any Excel processes that might be hanging
        Call KillExcelProcesses()
    End If
Next

' Clean up objects
Set fso = Nothing
Set shell = Nothing

' Display completion message
MsgBox "Password removal complete!" & vbCrLf & _
       "Successfully processed " & processedFiles & " of " & totalFiles & " Excel files.", _
       vbInformation, "Password Removal Complete"