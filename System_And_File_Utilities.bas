Attribute VB_Name = "System_And_File_Utilities"
Option Explicit

' --- API Declarations ---
' Refactored from: modSleep / modSleep2
#If VBA7 Then
    Private Declare PtrSafe Sub kernel32_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub kernel32_Sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
#End If

' -------------------------------------------------------------------------
' Subroutine: SafeSleep
' Purpose: Pauses execution securely, keeping Excel responsive.
' -------------------------------------------------------------------------
Public Sub SafeSleep(Optional ByVal milliseconds As Long = 1000)
    DoEvents
    kernel32_Sleep milliseconds
    DoEvents
End Sub

' -------------------------------------------------------------------------
' Subroutine: ListFormattedSubfolders
' Refactored from: ListarNomesDePastas
' Purpose: Uses the FileSystemObject to scan a directory and list its 
'          subfolders in Excel, formatting them to a specific pattern.
' -------------------------------------------------------------------------
Public Sub ListFormattedSubfolders()
    Dim fso As Object
    Dim folderObj As Object, subfolderObj As Object
    Dim targetSheet As Worksheet
    Dim mainFolderPath As String
    Dim formattedName As String
    Dim nextRow As Long
    
    ' Original hardcoded path
    mainFolderPath = "D:\Documents\rnd1\dev\"
    Set targetSheet = ThisWorkbook.Sheets("resumo")
    nextRow = 2
    
    On Error GoTo ErrorHandler
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(mainFolderPath)
    
    Helper_Core.TogglePerformance False
    
    For Each subfolderObj In folderObj.Subfolders
        ' Format the folder name (e.g., to "0000.0000")
        formattedName = Format(subfolderObj.Name, "@@@@.@@@@")
        
        If Len(formattedName) < 9 Then
            formattedName = "0" & formattedName
        End If
        
        targetSheet.Cells(nextRow, 1).Value = formattedName
        nextRow = nextRow + 1
    Next subfolderObj
    
CleanUp:
    Set fso = Nothing
    Set folderObj = Nothing
    Helper_Core.TogglePerformance True
    MsgBox "Subfolders listed successfully.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error accessing folders: " & Err.Description, vbCritical
    Resume CleanUp
End Sub