Attribute VB_Name = "Data_Custom_Filters"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: FilterAndMoveLocalities
' Refactored from: migrarLocalidades / b_migrarLocalidades
' Purpose: Moves rows from sheet '8135' to 'outrasLocalidades' if they DO NOT 
'          contain specific locality codes (207, 453, 565).
' -------------------------------------------------------------------------
Public Sub FilterAndMoveLocalities()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long, i As Long
    
    Set wsSource = ThisWorkbook.Sheets("8135")
    Set wsDest = ThisWorkbook.Sheets("outrasLocalidades")
    
    lastRow = Helper_Core.GetLastRow(wsSource, "A")
    Helper_Core.TogglePerformance False
    
    ' Loop backwards when moving/deleting rows
    For i = lastRow To 2 Step -1
        If InStr(wsSource.Cells(i, 1).Value, "207") = 0 And _
           InStr(wsSource.Cells(i, 1).Value, "453") = 0 And _
           InStr(wsSource.Cells(i, 1).Value, "565") = 0 Then
               
            ' Cut and paste to the next available row in destination
            wsSource.Rows(i).Cut Destination:=wsDest.Rows(Helper_Core.GetLastRow(wsDest, "A") + 1)
            wsSource.Rows(i).Delete
        End If
    Next i
    
    wsDest.Columns.AutoFit
    Helper_Core.TogglePerformance True
    MsgBox "Localities filtered and moved successfully.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: CalculateDecaHex
' Refactored from: geraLote
' Purpose: Injects an Excel formula to convert decimal sums to Hexadecimal.
' -------------------------------------------------------------------------
Public Sub CalculateDecaHex()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("fixit")
    Set dataRange = ws.Range("E2:H" & Helper_Core.GetLastRow(ws, "E"))

    Helper_Core.TogglePerformance False

    For i = 2 To dataRange.Rows.Count
        If Application.WorksheetFunction.Sum(dataRange.Rows(i)) = 0 Then
            If ws.Cells(i + 1, "D").Value = ws.Cells(i, "D").Value Then
                dataRange.Rows(i).Value = ws.Range("E" & i + 2 & ":H" & i + 2).Value
                ws.Cells(i + 1, "I").FormulaR1C1 = "=DECAHEX((SUM(RC[-7]/128,RC[-5]:RC[-1])),6)"
                ws.Cells(i + 1, "L").Value = "II"
            End If
        End If
    Next i

    Helper_Core.TogglePerformance True
    MsgBox "DecaHex calculations applied.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: CopyDataBetweenWorkbooks
' Refactored from: copiaValor
' Purpose: Copies specific columns from an external workbook directly into 
'          the master tracker, mapping them to specific locations.
' -------------------------------------------------------------------------
Public Sub CopyDataBetweenWorkbooks()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, nextRowDest As Long
    Dim i As Long

    ' Assumes both workbooks are open in the current Excel instance
    On Error Resume Next
    Set wsDest = Workbooks("207").Sheets("ligacoes")
    On Error GoTo 0
    
    If wsDest Is Nothing Then
        MsgBox "Target workbook '207' or sheet 'ligacoes' not found.", vbExclamation
        Exit Sub
    End If

    Helper_Core.TogglePerformance False

    ' Loop through all sheets in the analysis workbook
    For Each wsSource In Workbooks("analiseV1").Sheets
        lastRowSource = Helper_Core.GetLastRow(wsSource, "A")
    
        For i = 2 To lastRowSource
            nextRowDest = Helper_Core.GetLastRow(wsDest, "A") + 1

            wsDest.Cells(nextRowDest, "A").Value = wsSource.Cells(i, "A").Value
            wsDest.Cells(nextRowDest, "B").Value = wsSource.Cells(i, "B").Value
            wsDest.Cells(nextRowDest, "C").Value = wsSource.Cells(i, "I").Value
            wsDest.Cells(nextRowDest, "E").Value = wsSource.Cells(i, "C").Value
            wsDest.Cells(nextRowDest, "G").Value = wsSource.Cells(i, "D").Value
            wsDest.Cells(nextRowDest, "H").Value = wsSource.Cells(i, "E").Value
            wsDest.Cells(nextRowDest, "I").Value = wsSource.Cells(i, "H").Value
            wsDest.Cells(nextRowDest, "J").Value = wsSource.Cells(i, "F").Value
            wsDest.Cells(nextRowDest, "K").Value = wsSource.Cells(i, "G").Value
        Next i
    Next wsSource

    Helper_Core.TogglePerformance True
    MsgBox "Cross-workbook data copy completed.", vbInformation
End Sub