Attribute VB_Name = "Terminal_Scheduled_Bots"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: RunTimeLimitedBatch
' Refactored from: execucaoAgendada
' Purpose: Runs a terminal loop (e.g., Screen 05) but safely aborts execution
'          when the system clock reaches a user-defined time limit.
' -------------------------------------------------------------------------
Public Sub RunTimeLimitedBatch()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim timeLimitStr As String, timeLimit As Date
    Dim customerId As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    ' Prompt for the cutoff time
    timeLimitStr = InputBox("Enter the stop time (HH:MM format):" & vbCrLf & "Example: 11:59", "Set Time Limit")
    If timeLimitStr = "" Then Exit Sub
    
    On Error Resume Next
    timeLimit = TimeValue(timeLimitStr)
    If Err.Number <> 0 Then
        MsgBox "Invalid time format! Use HH:MM.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3
    
    For Each currentCell In dataRange
        ' Check if the time limit has been reached
        If Time >= timeLimit Then
            targetSheet.Cells(currentCell.Row, 2).Value = "Stopped at time limit"
            MsgBox "Process interrupted! Time limit (" & timeLimitStr & ") reached.", vbExclamation
            Exit For
        End If
        
        customerId = Format(currentCell.Value, "00000000")
        
        If customerId <> "00000000" Then
            ' --- INSERT TERMINAL LOGIC HERE (E.g., Screen 05) ---
            term.PutString 22, 11, "05" & customerId
            term.Enter
            ' ... (Data extraction logic) ...
            term.SendPFKey 3
        End If
        
        ' Update progress with current time
        Application.StatusBar = "Processing: " & customerId & " | Current Time: " & Format(Time, "HH:MM:SS")
        DoEvents ' Ensure the system clock updates
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Scheduled Batch Finished!", vbInformation
End Sub