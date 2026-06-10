Attribute VB_Name = "Terminal_Updates"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: UpdateCustomerContactInfo
' Refactored from: t74AtualizacaoCelular / numCelular
' Purpose: Navigates to Screen 74 to update or clean up customer phone 
'          numbers, emails, and bypasses common registry errors.
' -------------------------------------------------------------------------
Public Sub UpdateCustomerContactInfo()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
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
        ' Only process rows that haven't been altered yet
        If targetSheet.Cells(currentCell.Row, "B").Value <> "Altered" Then
            Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
            customerId = Format(currentCell.Value, "00000000")

            ' Navigate to Screen 74 (Customer Info Update)
            term.PutString 22, 11, "74"
            term.Enter
            term.PutString 19, 57, "01" ' Operation code for Update
            term.Enter
            term.PutString 6, 15, customerId
            term.Enter
            
            ' Clean up malformed landline fields
            If term.GetString(11, 20, 1) <> "_" Then
                term.PutString 11, 15, String(10, " ")
            ElseIf term.GetString(11, 20, 1) = "_" Then
                term.PutString 11, 15, "  "
            End If
            
            ' Commit changes and navigate to the next page of the registry (F12)
            term.PutString 18, 21, "N"
            term.SendPFKey 12
            
            ' Handle "Missing Name" Error
            If term.GetString(1, 2, 4) = "RAZA" Then
                term.PutString 8, 15, "Sem nome de cliente"
                term.PutString 15, 15, "Missing data bypass"
                term.SendPFKey 12
            End If
            
            ' Handle "Missing DOB" Error
            If term.GetString(1, 2, 4) = "INFO" Then
                term.PutString 16, 66, "01011960" ' Placeholder DOB
                term.PutString 15, 15, "Placeholder DOB applied"
                term.SendPFKey 12
            End If
            
            ' Final Confirmation Phase
            If term.GetString(17, 24, 4) = "ATEN" Then
                term.Enter
                term.PutString 17, 52, "s"
                term.Enter
            End If
            
            ' Save the record
            term.PutString 13, 58, "s"
            term.Enter
            
            ' Bypass print requests
            If term.GetString(10, 25, 4) = "EMIT" Then
                term.PutString 10, 47, "n"
                term.Enter
            End If
            
            ' Exit the update loop
            term.PutString 4, 19, "n"
            targetSheet.Cells(currentCell.Row, "B").Value = "Altered"
            term.Enter
            term.SendPFKey 3
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Customer Contact Updates Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub