Attribute VB_Name = "Terminal_Ultra_Niche"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractLogCode
' Refactored from: coletaCodLog (Screen 70 > 320)
' Purpose: Navigates to Screen 70, Option 320, and extracts a 5-char code.
' -------------------------------------------------------------------------
Public Sub ExtractLogCode()
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
        term.PutString 22, 11, "70"
        term.Enter
        term.PutString 22, 9, "320"
        term.Enter
        
        ' Specific sequence from original macro
        term.PutString 4, 17, "C"
        term.PutString 13, 19, "2"
        term.Enter
        term.PutString 6, 18, "120"
        term.PutString 8, 18, currentCell.Value
        term.Enter

        ' Extract Log Code
        targetSheet.Cells(currentCell.Row, "B").Value = Trim(term.GetString(6, 2, 5))
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Log codes extracted.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: MassUpdateCreditControl
' Refactored from: insereS (Screen 47 > 15)
' Purpose: Inputs values (like "s") directly from Excel Col F into the 
'          terminal grid (Rows 8 to 22), handling pagination automatically.
' -------------------------------------------------------------------------
Public Sub MassUpdateCreditControl()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim unitCode As String
    Dim termRow As Integer
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "F")
    If lastRow < 2 Then Exit Sub
    
    unitCode = InputBox("Enter the Unit Code:", "Unit Input")
    If unitCode = "" Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("F2:F" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    term.PutString 22, 11, "47"
    term.Enter
    term.PutString 22, 57, "15"
    term.Enter
    term.SetCursor 9, 28
    term.Enter
    term.PutString 15, 60, unitCode
    term.Enter
    
    termRow = 8 ' Grid starts at row 8
    
    For Each currentCell In dataRange
        term.SetCursor termRow, 5
        ' Original used SendKeys, PutString is safer and faster for HLLAPI
        term.PutString termRow, 5, currentCell.Value 
        
        termRow = termRow + 1
        
        ' Handle Pagination
        If termRow > 22 Then
            term.SendPFKey 8 ' F8 Next Page
            termRow = 8      ' Reset grid row
        End If
        
        ' Stop if end of records is reached
        If Trim(term.GetString(2, 4, 11)) = "ULTIMA TELA" Then
            term.Enter
            term.SendPFKey 3
            MsgBox "Process finished. Reached the last screen.", vbInformation
            Exit For
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Mass update completed.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ExtractScreen07Date
' Refactored from: tela7 / dataT7
' Purpose: Navigates to Screen 07 and extracts a specific 8-char string.
' -------------------------------------------------------------------------
Public Sub ExtractScreen07Date()
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
        ' Only process if Col B is blank
        If Trim(targetSheet.Cells(currentCell.Row, "B").Value) = "" Then
            term.PutString 22, 11, "07"
            term.Enter
            
            term.PutString 16, 43, Format(currentCell.Value, "00000000")
            term.Enter
            
            If term.GetString(15, 26, 3) = "ATE" Then
                term.PutString 16, 72, "n"
                term.Enter
            End If
            
            targetSheet.Cells(currentCell.Row, "B").Value = Trim(term.GetString(7, 8, 8))
            term.SendPFKey 3
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Screen 07 Dates Extracted.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ExtractCustomOffsets
' Refactored from: v_hdEImplantacao, za_nomeRequisicao, z_tela3_NomeComplemento
' Purpose: Handles the specific, non-standard offset scraping from Screens 03/08.
' -------------------------------------------------------------------------
Public Sub ExtractCustomOffsets()
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
        ' --- Example: Screen 08 specific scrape (za_nomeRequisicao) ---
        term.PutString 22, 11, "08"
        term.Enter
        term.PutString 17, 39, "        "
        term.PutString 20, 39, Format(currentCell.Value, "00000000")
        term.Enter
        
        If term.GetString(14, 22, 4) <> "PROT" Then
            If term.GetString(15, 26, 3) = "ATE" Then
                term.PutString 16, 72, "n"
                term.Enter
            End If
            
            targetSheet.Cells(currentCell.Row, "B").Value = Trim(term.GetString(5, 45, 34))
            targetSheet.Cells(currentCell.Row, "C").Value = Trim(term.GetString(4, 45, 34))
        Else
            targetSheet.Cells(currentCell.Row, "D").Value = "Check Protocol"
        End If
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Custom Offsets Extracted.", vbInformation
End Sub