Attribute VB_Name = "Terminal_Property_Lookup"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: LookupNeighboringProperties
' Refactored from: procuraMatriculaFrente / procuraMatriculaLado / tela02_Aprimorada
' Purpose: Navigates to Screen 02 to look up neighboring property IDs 
'          (front, side) based on logradouro and property numbers.
' -------------------------------------------------------------------------
Public Sub LookupNeighboringProperties()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim startTime As Double, endTime As Double
    Dim propertyNumber As String, sideNumber As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A") ' Assuming base ID is in A
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    startTime = Timer
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' Assumes Logradouro is in Col A, Property No. in Col C, Side No. in Col D
        propertyNumber = Format(targetSheet.Cells(currentCell.Row, "C").Value, "00000")
        sideNumber = Format(targetSheet.Cells(currentCell.Row, "D").Value, "00000")
        
        term.PutString 22, 11, "02"
        term.Enter
        
        term.PutString 10, 36, Format(currentCell.Value, "000") ' Logradouro Code
        term.PutString 12, 36, propertyNumber
        term.PutString 13, 36, sideNumber
        term.PutString 13, 46, "     " ' Clear unused fields
        term.PutString 13, 56, "     "
        term.Enter
        
        ' Handle Terminal Responses
        If term.GetString(7, 14, 5) = "ENTRE" Or term.GetString(7, 18, 3) = "MAT" Then
            targetSheet.Cells(currentCell.Row, "G").Value = "ID Not Found"
            targetSheet.Cells(currentCell.Row, "H").Value = "Verify Manually"
        ElseIf term.GetString(4, 7, 5) = "LOCAL" Then
            targetSheet.Cells(currentCell.Row, "G").Value = Replace(term.GetString(5, 13, 9), ".", "")
            term.SetCursor 5, 13
            term.Enter
            targetSheet.Cells(currentCell.Row, "H").Value = Trim(term.GetString(6, 58, 3))
        Else
            targetSheet.Cells(currentCell.Row, "G").Value = Replace(term.GetString(4, 10, 9), ".", "")
            targetSheet.Cells(currentCell.Row, "H").Value = Trim(term.GetString(6, 58, 3))
        End If
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Neighbor Lookup Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub