Attribute VB_Name = "Terminal_Tariff_Simulation"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: SimulateTariffs
' Refactored from: simulacao / consultaValoresTela14
' Purpose: Navigates to Screen 14 to run tariff simulations based on 
'          historical dates and consumption categories.
' -------------------------------------------------------------------------
Public Sub SimulateTariffs()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "F") ' Assuming trigger is in F
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("F2:F" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "14"
        term.Enter
        
        ' Bypass initial prompt
        term.PutString 15, 64, "s"
        term.Enter
        
        ' Input Simulation Parameters
        term.PutString 11, 37, "207" ' Example base parameter
        term.PutString 13, 37, Format(targetSheet.Cells(currentCell.Row, "G").Value, "00000000") ' Customer ID
        term.PutString 14, 37, "011"
        term.PutString 15, 37, currentCell.Value ' Dynamic Value
        term.PutString 17, 37, "001"
        term.Enter
        
        ' Extract Result
        targetSheet.Cells(currentCell.Row, "H").Value = Trim(term.GetString(22, 39, 10))
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Tariff Simulations Completed.", vbInformation, "Completed"
End Sub