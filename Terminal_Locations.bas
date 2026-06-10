Attribute VB_Name = "Terminal_Locations"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractStreetCrossings
' Refactored from: coletaCruzamentos / coletaTrechos
' Purpose: Navigates to Screen 70 (Location/Crossings) using a specific 
'          city/locality code. Extracts all listed street intersections.
' -------------------------------------------------------------------------
Public Sub ExtractStreetCrossings()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long, nextEmptyRow As Long
    Dim localityCode As String
    Dim termRow As Integer
    Dim isFinished As Boolean
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    ' Prompt user for the 3-digit locality code
    localityCode = InputBox("Enter the 3-digit locality code (e.g., 209):", "Locality Code")
    If Len(localityCode) <> 3 Then
        MsgBox "Invalid code. Must be 3 digits.", vbExclamation
        Exit Sub
    End If

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    nextEmptyRow = lastRow + 1
    
    term.Connect
    term.SendPFKey 3
    isFinished = False

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' Navigate to Screen 70 > Sub-screen 620
        term.PutString 22, 11, "70"
        term.Enter
        term.PutString 22, 9, "620"
        term.Enter
        
        ' Input search parameters (c2 = Consult)
        term.PutString 4, 17, "c2"
        term.Enter
        term.PutString 4, 18, localityCode
        term.Enter
        
        ' Loop through pagination
        Do
            ' Terminal rows 10 to 23 contain the crossing data
            For termRow = 10 To 23
                Dim mainStreet As String, cornerStreet As String
                
                mainStreet = Trim(term.GetString(6, 19, 5))
                cornerStreet = Trim(term.GetString(termRow, 12, 5))
                
                If mainStreet <> "" Then
                    targetSheet.Cells(nextEmptyRow, "A").Value = mainStreet
                    targetSheet.Cells(nextEmptyRow, "B").Value = Trim(term.GetString(termRow, 6, 6)) ' Corner ID
                    targetSheet.Cells(nextEmptyRow, "C").Value = cornerStreet
                    targetSheet.Cells(nextEmptyRow, "D").Value = Trim(term.GetString(termRow, 64, 1)) ' Side
                    targetSheet.Cells(nextEmptyRow, "E").Value = Trim(term.GetString(termRow, 70, 5)) ' Crossing ID
                    nextEmptyRow = nextEmptyRow + 1
                End If
            Next termRow
            
            term.Enter ' Next page
            
            ' If the locality code changes, we reached the end of this search
            If Trim(term.GetString(4, 18, 3)) <> localityCode Then
                isFinished = True
                Exit Do
            End If
        Loop Until isFinished
        
        If isFinished Then Exit For
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Street crossings extracted successfully!", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: DeleteStreetCrossings
' Refactored from: excluirCruzamentos / removeCruzamentos
' Purpose: Removes street crossings based on locality and crossing codes.
' -------------------------------------------------------------------------
Public Sub DeleteStreetCrossings()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim localityCode As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    localityCode = InputBox("Enter the 3-digit locality code:", "Locality Code")
    If localityCode = "" Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "70"
        term.Enter
        term.PutString 22, 9, "620"
        term.Enter
        
        ' 'e1' command for Exclusion
        term.PutString 4, 17, "e1"
        term.Enter
        term.PutString 4, 18, localityCode
        term.PutString 6, 17, currentCell.Value
        term.Enter
        
        ' Handle system responses
        If term.GetString(1, 2, 3) = "REG" Then
            targetSheet.Cells(currentCell.Row, "G").Value = "Record not found"
        Else
            ' Confirm deletion
            term.PutString 21, 46, "s"
            term.Enter
            targetSheet.Cells(currentCell.Row, "G").Value = "Removed successfully"
        End If
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Deletion process completed.", vbInformation, "Completed"
End Sub