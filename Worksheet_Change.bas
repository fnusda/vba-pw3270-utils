' -------------------------------------------------------------------------
' Event: Worksheet_Change
' Refactored from: Worksheet_Change (Módulo1)
' Purpose: Triggers automatically when a cell is edited. Specifically monitors 
'          cell F2 and alerts the user if its contents are deleted.
' -------------------------------------------------------------------------
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Check if the cell that changed intersects with F2
    If Not Intersect(Target, Me.Range("F2")) Is Nothing Then
        ' If the user cleared the cell, trigger a warning
        If Target.Value = "" Then
            MsgBox "The contents of cell " & Target.Address & " were deleted.", vbInformation, "Cell Cleared"
        End If
    End If
End Sub