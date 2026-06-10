Attribute VB_Name = "Data_Management"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ArchiveProcessedRecords
' Refactored from: mudarDeAba / arquivarProtocolos8135
' Purpose: Moves completed records from an active queue to a history tab,
'          stamping them with the current date.
' -------------------------------------------------------------------------
Public Sub ArchiveProcessedRecords()
    Dim wsSource As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRowSrc As Long
    Dim nextRowHist As Long
    Dim numRowsToMove As Long
    
    Set wsSource = ThisWorkbook.Worksheets("a_solicitar")
    Set wsHistory = ThisWorkbook.Worksheets("ja_solicitado")
    
    lastRowSrc = Helper_Core.GetLastRow(wsSource, "A")
    If lastRowSrc < 3 Then
        MsgBox "No data found in 'a_solicitar' to archive.", vbExclamation, "Notice"
        Exit Sub
    End If
    
    Helper_Core.TogglePerformance False
    
    numRowsToMove = lastRowSrc - 2
    nextRowHist = Helper_Core.GetLastRow(wsHistory, "A") + 1
    
    ' 1. Bulk copy columns A and B
    wsHistory.Range("A" & nextRowHist & ":B" & nextRowHist + numRowsToMove - 1).Value = _
        wsSource.Range("A3:B" & lastRowSrc).Value
        
    ' 2. Stamp the current date in Column C
    With wsHistory.Range("C" & nextRowHist & ":C" & nextRowHist + numRowsToMove - 1)
        .Value = Date
        .NumberFormat = "dd/mm/yyyy"
    End With
    
    ' 3. Clear source data
    wsSource.Range("A3:B" & lastRowSrc).ClearContents
    
    Helper_Core.TogglePerformance True
    MsgBox numRowsToMove & " records successfully archived.", vbInformation, "Migration Complete"
End Sub

' -------------------------------------------------------------------------
' Subroutine: RemoveDuplicateRequests
' Refactored from: verificarSolicitacao
' Purpose: Compares the active queue against the history tab and deletes
'          any requests that have already been processed.
' -------------------------------------------------------------------------
Public Sub RemoveDuplicateRequests()
    Dim wsSource As Worksheet
    Dim wsHistory As Worksheet
    Dim lastRowSrc As Long, lastRowHist As Long
    Dim i As Long
    Dim searchVal As Variant
    Dim removalCount As Long
    Dim historyRange As Range
    
    Set wsSource = ThisWorkbook.Worksheets("a_solicitar")
    Set wsHistory = ThisWorkbook.Worksheets("ja_solicitado")
    
    lastRowSrc = Helper_Core.GetLastRow(wsSource, "A")
    lastRowHist = Helper_Core.GetLastRow(wsHistory, "A")
    
    If lastRowSrc < 3 Then Exit Sub
    
    Set historyRange = wsHistory.Range("A2:A" & lastRowHist)
    removalCount = 0
    
    Helper_Core.TogglePerformance False
    
    ' Loop backwards when deleting rows
    For i = lastRowSrc To 3 Step -1
        searchVal = wsSource.Cells(i, 1).Value
        If searchVal <> "" Then
            If Application.WorksheetFunction.CountIf(historyRange, searchVal) > 0 Then
                wsSource.Rows(i).Delete
                removalCount = removalCount + 1
            End If
        End If
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Duplicate check complete! " & removalCount & " existing records were removed.", vbInformation, "Done"
End Sub

' -------------------------------------------------------------------------
' Subroutine: StackColumnsIntoA
' Refactored from: u_ajeitaLinhas
' Purpose: Takes a wide dataset and stacks all columns sequentially underneath
'          Column A. Useful for flattening data arrays.
' -------------------------------------------------------------------------
Public Sub StackColumnsIntoA()
    Dim ws As Worksheet
    Dim lastCol As Long, lastRowCol As Long, nextRowA As Long
    Dim col As Long
    
    Set ws = ActiveSheet
    Helper_Core.TogglePerformance False
    
    ' Find absolute last column
    lastCol = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    For col = 2 To lastCol
        lastRowCol = Helper_Core.GetLastRow(ws, ws.Cells(1, col).Address(False, False))
        
        If lastRowCol > 1 Then
            nextRowA = Helper_Core.GetLastRow(ws, "A") + 1
            
            ' Move data
            ws.Range(ws.Cells(2, col), ws.Cells(lastRowCol, col)).Copy Destination:=ws.Range("A" & nextRowA)
            ws.Range(ws.Cells(2, col), ws.Cells(lastRowCol, col)).ClearContents
        End If
    Next col
    
    Helper_Core.TogglePerformance True
    MsgBox "All columns stacked into Column A.", vbInformation
End Sub