Attribute VB_Name = "Terminal_Additional_Queries"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractBilledServices
' Refactored from: Tela_9 / importaDados (Screen 09)
' Purpose: Navigates to Screen 09 to extract billed service history, 
'          paginating through the terminal grid.
' -------------------------------------------------------------------------
Public Sub ExtractBilledServices()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim termRow As Integer
    Dim excelRow As Long
    
    Set targetSheet = ActiveSheet
    excelRow = 3 ' Starts pasting on row 3
    
    Helper_Core.TogglePerformance False
    term.Connect
    term.SendPFKey 3
    
    term.PutString 22, 11, "09"
    term.Enter
    
    term.PutString 17, 40, targetSheet.Cells(1, 2).Value ' Uses param from B1
    term.Enter
    
    If term.GetString(14, 22, 4) = "NADA" Then
        term.SendPFKey 3
        GoTo CleanUp
    End If
    
    Do While True
        For termRow = 8 To 24
            targetSheet.Cells(excelRow, "A").Value = term.GetString(termRow, 2, 4)
            targetSheet.Cells(excelRow, "B").Value = term.GetString(termRow, 7, 2)
            targetSheet.Cells(excelRow, "C").Value = term.GetString(termRow, 10, 2)
            targetSheet.Cells(excelRow, "D").Value = term.GetString(termRow, 23, 9)
            targetSheet.Cells(excelRow, "E").Value = term.GetString(termRow, 34, 7)
            targetSheet.Cells(excelRow, "I").Value = term.GetString(termRow, 54, 8)
            targetSheet.Cells(excelRow, "J").Value = term.GetString(termRow, 64, 17)
            excelRow = excelRow + 1
        Next termRow
        
        term.SetCursor 3, 9
        term.SendPFKey 2 ' Next page
        
        If term.GetString(1, 2, 5) = "OPCAO" Then Exit Do
    Loop

CleanUp:
    term.SendPFKey 3
    Helper_Core.TogglePerformance True
    MsgBox "Billed Services Extraction Complete.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ExtractCustomerDetails
' Refactored from: t39ColetaCPF (Screen 39)
' Purpose: Extracts detailed customer identification data (Name, CPF, DOB).
' -------------------------------------------------------------------------
Public Sub ExtractCustomerDetails()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        term.PutString 22, 11, "39"
        term.Enter
        term.PutString 12, 40, Replace(Format(currentCell.Value, "00000000"), ".", "")
        term.Enter
        
        targetSheet.Cells(currentCell.Row, "B").Value = term.GetString(18, 39, 1)
        targetSheet.Cells(currentCell.Row, "C").Value = Replace(Trim(term.GetString(18, 41, 16)), "_", "")
        targetSheet.Cells(currentCell.Row, "D").Value = Replace(Trim(term.GetString(19, 39, 35)), "_", "")
        
        ' Format Date of Birth
        Dim d As String, m As String, y As String
        d = Format(Replace(term.GetString(20, 43, 2), "_", ""), "00")
        m = Format(Replace(term.GetString(20, 48, 2), "_", ""), "00")
        y = Replace(term.GetString(20, 53, 4), "_", "")
        targetSheet.Cells(currentCell.Row, "E").Value = d & "/" & m & "/" & y
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Customer Details Extraction Complete.", vbInformation
End Sub