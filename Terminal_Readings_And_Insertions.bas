Attribute VB_Name = "Terminal_Readings_And_Insertions"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractMeterReadings
' Refactored from: infosTela4 (Screen 04)
' Purpose: Extracts a block of historical meter readings horizontally.
' -------------------------------------------------------------------------
Public Sub ExtractMeterReadings()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim termRow As Integer, excelCol As Integer
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "04"
        term.PutString 22, 24, Format(currentCell.Value, "00000000")
        term.Enter
        
        targetSheet.Cells(currentCell.Row, "B").Value = term.GetString(18, 26, 12)
        
        excelCol = 3 ' Start at Column C
        For termRow = 7 To 19
            targetSheet.Cells(currentCell.Row, excelCol).Value = term.GetString(termRow, 5, 7)
            targetSheet.Cells(currentCell.Row, excelCol + 1).Value = Trim(term.GetString(termRow, 46, 4))
            targetSheet.Cells(currentCell.Row, excelCol + 2).Value = Trim(term.GetString(termRow, 61, 3))
            targetSheet.Cells(currentCell.Row, excelCol + 3).Value = Trim(term.GetString(termRow, 70, 5))
            targetSheet.Cells(currentCell.Row, excelCol + 4).Value = Trim(term.GetString(termRow, 76, 5))
            excelCol = excelCol + 5
        Next termRow
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Meter Readings Extracted.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: InsertNeighborhoodsAndBlocks
' Refactored from: incluiBairro / incluiQuadras (Screens 6904 / 6926)
' Purpose: Injects new neighborhood or block data and verifies registration.
' -------------------------------------------------------------------------
Public Sub InsertNeighborhoodsAndBlocks(ByVal isBlock As Boolean)
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim screenCode As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "D")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("D2:D" & lastRow)
    screenCode = IIf(isBlock, "6926", "6904")
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        term.PutString 22, 11, screenCode
        term.Enter
        term.PutString 4, 16, "i"
        term.Enter
        
        If isBlock Then
            term.PutString 6, 22, currentCell.Value
            term.Enter
            term.Enter
        Else
            term.PutString 5, 11, Left(currentCell.Value, 3)
            term.Enter
            term.PutString 6, 11, Mid(currentCell.Value, 4)
            term.Enter
        End If
        
        If term.GetString(1, 2, 22) = "ELEMENTO JA CADASTRADO" Then
            targetSheet.Cells(currentCell.Row, "E").Value = "Already Registered"
        Else
            term.PutString 18, 73, "s"
            term.Enter
            targetSheet.Cells(currentCell.Row, "E").Value = "Inserted"
            term.Enter
        End If
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Insertion Complete.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: RegisterPreventiveReplacement
' Refactored from: d_registraTela24 (Screen 24)
' Purpose: Registers a meter replacement and saves the protocol to a new sheet.
' -------------------------------------------------------------------------
Public Sub RegisterPreventiveReplacement()
    Dim term As New clsTerminalMacros
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long, nextRowDest As Long
    Dim protocol As String
    
    Set wsOrigem = ThisWorkbook.Worksheets("matriculas")
    Set wsDestino = ThisWorkbook.Worksheets("solicitarTroca")
    
    lastRow = Helper_Core.GetLastRow(wsOrigem, "A")
    If lastRow < 2 Then Exit Sub

    wsDestino.Rows("2:" & wsDestino.Rows.Count).ClearContents
    nextRowDest = 2

    Helper_Core.TogglePerformance False
    Set dataRange = wsOrigem.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        If currentCell.Value > 0 Then
            term.PutString 22, 11, "24"
            term.Enter
            term.PutString 9, 28, Format(currentCell.Value, "00000000")
            term.PutString 10, 28, "0541"
            term.PutString 13, 25, "preferenc ligar cli"
            
            If wsOrigem.Cells(currentCell.Row, "B").Value <> "0" Then
                term.PutString 14, 25, wsOrigem.Cells(currentCell.Row, "B").Value
            End If
            
            If wsOrigem.Cells(currentCell.Row, "E").Value <> "" Then
                term.PutString 14, 48, wsOrigem.Cells(currentCell.Row, "E").Value
            End If
            
            term.PutString 15, 25, "troca preventiva ccpg"
            term.Enter
            term.SendPFKey 2
            term.PutString 24, 56, "s"
            term.Enter
            
            protocol = Trim(term.GetString(14, 39, 17))
            wsOrigem.Cells(currentCell.Row, "F").Value = protocol
            
            wsDestino.Cells(nextRowDest, "A").Value = currentCell.Value
            wsDestino.Cells(nextRowDest, "B").Value = protocol
            nextRowDest = nextRowDest + 1
            term.Enter
        Else
            wsOrigem.Cells(currentCell.Row, "G").Value = "No ID"
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Replacements Registered.", vbInformation
End Sub