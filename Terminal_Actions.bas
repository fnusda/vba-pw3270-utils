Attribute VB_Name = "Terminal_Actions"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: CloseAdministrativeService
' Refactored from: baixaAdmin / baixaAdminT19 / baixaT19
' Purpose: Navigates to Screen 19 to close/finalize administrative services.
'          Reads the service number from Column A and details from B, C, D, E.
' -------------------------------------------------------------------------
Public Sub CloseAdministrativeService()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim startTime As Double, endTime As Double

    startTime = Timer
    Set targetSheet = ActiveSheet
    
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False

    term.Connect
    term.SendPFKey 3 ' Home screen

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' Only process if the cell has a valid service number (> 0)
        If Val(currentCell.Value) > 0 Then
            ' Navigate to Screen 19 (Service Closure)
            term.PutString 22, 11, "19"
            term.Enter
            
            ' Input closure details
            term.PutString 16, 31, currentCell.Value
            term.PutString 17, 31, "8135" ' Specific service code (adjust as needed)
            
            ' Build the observation string from adjacent columns
            term.PutString 18, 31, targetSheet.Cells(currentCell.Row, "B").Value & _
                                   " Fat em " & targetSheet.Cells(currentCell.Row, "C").Value & " vezes"
            term.PutString 19, 31, targetSheet.Cells(currentCell.Row, "D").Value
            term.PutString 20, 31, targetSheet.Cells(currentCell.Row, "E").Value
            term.Enter
            
            ' Check if there is an error prompt blocking the closure
            If term.GetString(14, 17, 1) <> " " Then
                term.SendPFKey 3
                targetSheet.Cells(currentCell.Row, "R").Value = "Check input fields (Error)"
            Else
                ' Confirm closure
                term.Enter
                targetSheet.Cells(currentCell.Row, "R").Value = "Successfully Closed"
                term.SendPFKey 3
            End If
        Else
            targetSheet.Cells(currentCell.Row, "R").Value = "Not Deployed"
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Service Closure Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: RegisterServiceOrder
' Refactored from: registra8125 / registra8152 / registraTela26
' Purpose: Navigates to Screen 26 to open/register a new service order.
'          Fetches the generated protocol number and saves it back to Excel.
' -------------------------------------------------------------------------
Public Sub RegisterServiceOrder()
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
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        customerId = Format(currentCell.Value, "00000000")
        
        ' Navigate to Screen 26 (Service Registration)
        term.PutString 22, 11, "26"
        term.Enter
        
        ' Input Customer ID and specific service code
        term.PutString 9, 28, customerId
        term.Enter
        term.PutString 10, 28, "8125" ' Service code (e.g., Expansion Plan)
        
        ' Input observations
        term.PutString 13, 25, "usar o survey"
        term.PutString 15, 25, "plano expansao esgoto"
        term.Enter
        
        ' Commit the registration using F2
        term.SendPFKey 2
        term.PutString 24, 56, "s" ' Confirm 'Yes'
        term.Enter
        
        ' Extract the newly generated protocol number and save it to Column B
        targetSheet.Cells(currentCell.Row, "B").Value = Trim(term.GetString(14, 39, 17))
        
        term.Enter
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Service Registration Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: CancelInvoiceReference
' Refactored from: cancelamento43 / cancelarReferencias
' Purpose: Prompts for a Customer ID, then loops through invoice references 
'          in Column A and cancels them on Screen 43.
' -------------------------------------------------------------------------
Public Sub CancelInvoiceReference()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
    Dim invoiceRef As String
    Dim startTime As Double, endTime As Double

    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    ' Ask user for the Customer ID to apply cancellations against
    customerId = InputBox("Enter the 8-digit Customer ID (Matrícula):", "Customer ID")
    If customerId = "" Then Exit Sub ' User cancelled
    
    customerId = Format(customerId, "00000000")
    
    startTime = Timer
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False

    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' Format the invoice reference (e.g., 012023 for Jan 2023)
        invoiceRef = Format(currentCell.Value, "000000")
        
        ' Navigate to Screen 43 (Cancellations)
        term.PutString 22, 11, "43"
        term.Enter
        
        ' Navigate past initial prompt
        term.PutString 15, 64, "s"
        term.Enter
        
        ' Enter Customer ID and Invoice Reference
        term.PutString 11, 39, customerId
        term.PutString 12, 39, invoiceRef
        term.Enter
        term.Enter ' Bypass potential warnings
        
        ' Enter cancellation reason code (e.g., 'as' for Administrative)
        term.PutString 10, 26, "as"
        term.Enter
        
        ' Check if cancellation was successful or if an error blocked it
        If term.GetString(1, 2, 6) = "FATURA" Then
            targetSheet.Cells(currentCell.Row, "K").Value = "Cancelled"
            term.SendPFKey 3
        Else
            ' Confirm the cancellation
            term.PutString 23, 71, "s"
            term.Enter
            targetSheet.Cells(currentCell.Row, "K").Value = "Confirmed"
        End If
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Invoice Cancellations Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub