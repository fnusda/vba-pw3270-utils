' -------------------------------------------------------------------------
' Subroutine: QueryPendingDebts
' Refactored from: consultaPendencias
' Purpose: Queries Screen 05 (Pending Debts) for each customer ID in Column A.
'          Extracts debt references from rows 11-20 and pastes them 
'          horizontally in Excel starting from Column F.
' -------------------------------------------------------------------------
Public Sub QueryPendingDebts()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
    Dim extractedDebt As String
    Dim terminalRow As Integer
    Dim targetColumn As Integer
    Dim startTime As Double, endTime As Double

    startTime = Timer
    Set targetSheet = ActiveSheet
    
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False

    term.Connect
    term.SendPFKey 3 ' Ensure terminal is at the home screen

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        customerId = Format(currentCell.Value, "00000000")

        ' Navigate to Screen 05
        term.PutString 22, 11, "05" & customerId
        term.Enter
        
        targetColumn = 6 ' Start pasting data in Column F (Index 6)
        
        ' Loop through the terminal rows (11 to 20) where debts are listed
        For terminalRow = 11 To 20
            extractedDebt = Trim(term.GetString(terminalRow, 9, 7))
            
            If extractedDebt <> "" Then
                targetSheet.Cells(currentCell.Row, targetColumn).Value = extractedDebt
                targetColumn = targetColumn + 1 ' Move to the next Excel column
            End If
        Next terminalRow
        
        term.SendPFKey 3 ' Return to home screen for the next iteration
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Pending Debts Query Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: QueryMobilePhones
' Refactored from: b_t8Celular / t8ConsultaCelular
' Purpose: Queries Screen 08 for mobile phones. Evaluates specific rows 
'          and safely handles terminal pagination (F8) to find the data.
' -------------------------------------------------------------------------
Public Sub QueryMobilePhones()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
    Dim safetyCounter As Integer
    Dim startTime As Double, endTime As Double

    startTime = Timer
    Set targetSheet = ActiveSheet
    
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False

    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        customerId = Format(currentCell.Value, "00000000")

        ' Navigate to Screen 08
        term.PutString 22, 11, "08"
        term.Enter
        
        ' Enter the Customer ID in the specific search field
        term.PutString 17, 39, customerId
        term.Enter
        
        ' Check if the expected "MATR" label is present
        If term.GetString(14, 22, 4) = "MATR" Then
            targetSheet.Cells(currentCell.Row, "C").Value = Trim(term.GetString(14, 42, 11))
        Else
            ' Evaluate multiple rows to find a mobile number (starts with 9)
            If term.GetString(11, 64, 1) = "9" Then
                targetSheet.Cells(currentCell.Row, "C").Value = Trim(term.GetString(11, 62, 11))
            
            ElseIf term.GetString(16, 64, 1) = "9" Then
                targetSheet.Cells(currentCell.Row, "C").Value = Trim(term.GetString(16, 62, 11))
            
            ElseIf term.GetString(21, 64, 1) = "9" Then
                targetSheet.Cells(currentCell.Row, "C").Value = Trim(term.GetString(21, 62, 11))
            
            Else
                ' Paginate through the terminal if the phone is not on the first page
                safetyCounter = 0
                
                Do
                    term.SendPFKey 8 ' F8 = Next Page
                    safetyCounter = safetyCounter + 1
                    
                    ' Prevent infinite loops if "ULTIMA" is never reached
                    If safetyCounter > 20 Then Exit Do 
                Loop Until Trim(term.GetString(7, 3, 6)) = "ULTIMA"
            End If
        End If
        
        term.SendPFKey 3 ' Return to home screen
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Mobile Phones Query Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub