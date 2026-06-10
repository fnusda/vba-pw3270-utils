Attribute VB_Name = "Terminal_Misc_Queries"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: VerifyEmployeeProductivity
' Refactored from: verificaFuncional (Screen 77)
' Purpose: Checks Screen 77 for employee productivity logs based on date.
' -------------------------------------------------------------------------
Public Sub VerifyEmployeeProductivity()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim termRow As Integer
    Dim foundRecord As Boolean
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 4 Then Exit Sub ' Original started at row 4

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A4:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 3, dataRange.Rows.Count
        
        term.PutString 22, 11, "77"
        term.Enter
        term.PutString 22, 61, "07"
        term.Enter
        
        ' Input Employee ID and Date
        term.PutString 3, 10, Format(currentCell.Value, "00000000")
        term.PutString 4, 16, targetSheet.Cells(2, "A").Value
        term.Enter
        
        If term.GetString(15, 3, 5) = "TOTAL" Then
            targetSheet.Cells(currentCell.Row, 2).Value = "No records for this date"
            term.SendPFKey 3
            term.SendPFKey 3
            GoTo NextIter
        End If
        
        foundRecord = False
        Do
            term.SendPFKey 8 ' F8 Next Page
            For termRow = 7 To 22
                If term.GetString(termRow, 6, 26) = "TOTAL DE CLIENTES ATENDIDO" Then
                    targetSheet.Cells(currentCell.Row, 2).Value = "Records Found"
                    foundRecord = True
                    Exit Do
                End If
            Next termRow
        Loop Until foundRecord
        
        term.SendPFKey 3
        term.SendPFKey 3

NextIter:
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Productivity Verification Completed.", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: LookupHDDetails
' Refactored from: procurarHDT71 (Screen 71)
' Purpose: Looks up hardware/meter details on Screen 71.
' -------------------------------------------------------------------------
Public Sub LookupHDDetails()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "V")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("V2:V" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "71"
        term.Enter
        term.PutString 16, 29, "03"
        term.Enter
        term.PutString 13, 44, currentCell.Value ' Protocol
        term.Enter
        
        targetSheet.Cells(currentCell.Row, "W").Value = Replace(term.GetString(5, 13, 9), ".", "")
        targetSheet.Cells(currentCell.Row, "X").Value = Replace(term.GetString(13, 2, 9), ".", "")
        
        term.SendPFKey 3
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "HD Lookup Completed.", vbInformation, "Completed"
End Sub