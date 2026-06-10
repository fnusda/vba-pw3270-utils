Attribute VB_Name = "Terminal_Service_Orders"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: SearchServiceOrders
' Refactored from: a_procura8135T11 / t11D46
' Purpose: Navigates to Screen 11, queries service 8135 in district 46, 
'          extracts grid data, and injects Excel formulas for analysis.
' -------------------------------------------------------------------------
Public Sub SearchServiceOrders()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim termRow As Integer
    Dim emptyCount As Integer
    Dim startTime As Double, endTime As Double
    
    startTime = Timer
    Set targetSheet = ActiveSheet
    
    Helper_Core.TogglePerformance False
    term.Connect
    term.SendPFKey 3
    
    ' Navigate to Screen 11
    term.PutString 22, 11, "11"
    term.Enter
    term.PutString 14, 40, "8135" ' Service
    term.PutString 18, 40, "46"   ' District
    term.SetCursor 19, 40
    term.Enter
    
    Do
        emptyCount = 0
        lastRow = Helper_Core.GetLastRow(targetSheet, "A")
        If lastRow < 2 Then lastRow = 1
        
        ' Loop through terminal grid (Rows 9 to 22)
        For termRow = 9 To 22
            If Trim(term.GetString(termRow, 5, 5)) = "" Then
                emptyCount = emptyCount + 1
            Else
                emptyCount = 0
            End If
            
            ' Stop if two consecutive empty rows are found
            If emptyCount = 2 Then Exit Do
            
            targetRow = lastRow + (termRow - 8)
            
            ' Extract Data
            targetSheet.Cells(targetRow, "A").Value = Trim(term.GetString(termRow, 76, 3))
            targetSheet.Cells(targetRow, "C").Value = Trim(term.GetString(termRow, 45, 5))
            targetSheet.Cells(targetRow, "D").Value = Trim(term.GetString(termRow, 51, 5))
            targetSheet.Cells(targetRow, "E").Value = Trim(term.GetString(termRow, 5, 8))
            targetSheet.Cells(targetRow, "F").Value = Format(Trim(term.GetString(termRow, 14, 4)), "0000")
            targetSheet.Cells(targetRow, "G").Value = Trim(term.GetString(termRow, 19, 5))
            
            If Trim(term.GetString(termRow, 25, 9)) = "" Then
                targetSheet.Cells(targetRow, "H").Value = "NL"
            Else
                targetSheet.Cells(targetRow, "H").Value = Format(Replace(Trim(term.GetString(termRow, 25, 9)), ".", ""), "00000000")
            End If
            
            targetSheet.Cells(targetRow, "I").Value = Now
            
            ' Inject Excel Formulas
            targetSheet.Cells(targetRow, "J").FormulaR1C1 = "=RC[-4]-1"
            targetSheet.Cells(targetRow, "K").FormulaR1C1 = "=IF(RC[2]="""",CONCAT(RC[-6],TEXT(RC[-1],""0000""),TEXT(RC[-4],""00000"")),CONCAT(RC[-6],TEXT(RC[2],""0000""),TEXT(RC[-4],""00000"")))"
            targetSheet.Cells(targetRow, "M").FormulaR1C1 = "=IF(RC[-1]=""Verificar Protocolo"",TEXT(ABS(RC[-7]-2400),""0000""),"""")"
            targetSheet.Cells(targetRow, "Q").FormulaR1C1 = "=CONCAT(RC[-12],TEXT(RC[-11],""0000""),TEXT(RC[-10],""00000""))"
        Next termRow
        
        ' Next Page
        term.SetCursor 3, 9
        term.Enter
    Loop
    
    term.SendPFKey 3
    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Screen 11 Extraction Complete! Time: " & Format(endTime - startTime, "0.00") & "s", vbInformation
End Sub