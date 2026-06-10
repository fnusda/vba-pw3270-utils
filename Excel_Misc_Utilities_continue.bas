' -------------------------------------------------------------------------
' Subroutine: DeleteAllOtherSheets
' Refactored from: ExcluirColunasEPlanilhas / limparDesnecessarios
' Purpose: Deletes columns B:Z on the "Ajustada" sheet and completely 
'          deletes every other worksheet in the workbook.
' -------------------------------------------------------------------------
Public Sub DeleteAllOtherSheets()
    Dim ws As Worksheet
    Dim targetSheet As Worksheet
    
    On Error Resume Next
    Set targetSheet = ThisWorkbook.Worksheets("Ajustada")
    On Error GoTo 0
    
    If targetSheet Is Nothing Then
        MsgBox "The sheet 'Ajustada' was not found.", vbExclamation
        Exit Sub
    End If
    
    Helper_Core.TogglePerformance False
    
    ' Clear columns B through Z
    targetSheet.Columns("B:Z").Delete
    
    ' Suppress deletion warnings and loop through sheets
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Ajustada" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    Helper_Core.TogglePerformance True
    MsgBox "Cleanup complete: Columns deleted and other sheets removed.", vbInformation
End Sub