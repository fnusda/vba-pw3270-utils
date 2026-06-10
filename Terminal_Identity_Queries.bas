Attribute VB_Name = "Terminal_Identity_Queries"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: SearchByCPF_CNPJ
' Refactored from: consultaPorCPF / consultaPorCNPJ
' Purpose: Navigates to Screen 02, inputs a CPF (1) or CNPJ (2), and 
'          extracts up to 5 associated property records.
' -------------------------------------------------------------------------
Public Sub SearchByCPF_CNPJ(ByVal isCNPJ As Boolean)
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim docType As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    docType = IIf(isCNPJ, "2", "1")
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "02"
        term.Enter
        
        ' Clear screen fields
        term.PutString 10, 36, "   "
        term.PutString 12, 36, "     "
        term.PutString 13, 36, "     "
        
        ' Input Document
        term.PutString 15, 36, docType
        term.PutString 15, 38, currentCell.Value
        term.Enter
        
        ' Extract up to 5 records horizontally
        targetSheet.Cells(currentCell.Row, "B").Value = term.GetString(5, 13, 55)
        targetSheet.Cells(currentCell.Row, "D").Value = term.GetString(9, 13, 55)
        targetSheet.Cells(currentCell.Row, "F").Value = term.GetString(13, 13, 55)
        targetSheet.Cells(currentCell.Row, "H").Value = term.GetString(17, 13, 55)
        targetSheet.Cells(currentCell.Row, "J").Value = term.GetString(21, 13, 55)
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Identity Search Completed.", vbInformation
End Sub