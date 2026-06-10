Attribute VB_Name = "Terminal_Invoicing_And_Retentions"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ManageInvoiceRetention
' Refactored from: retemCorte / removeRetCorte (Screen 40)
' Purpose: Navigates to Screen 40 to apply or remove an invoice retention 
'          (corte) based on a specific reason code.
' -------------------------------------------------------------------------
Public Sub ManageInvoiceRetention(ByVal isRemoving As Boolean)
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String, reasonCode As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    ' Prompt for parameters
    customerId = InputBox("Enter the 8-digit Customer ID:", "Customer ID")
    If customerId = "" Then Exit Sub
    
    If Not isRemoving Then
        reasonCode = InputBox("Enter the Retention Reason Code:", "Reason")
        If reasonCode = "" Then Exit Sub
    Else
        reasonCode = "s" ' Default release code from original macro
    End If

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "40"
        term.Enter
        
        term.PutString 17, 40, Format(customerId, "00000000")
        term.PutString 18, 40, Format(currentCell.Value, "000000") ' Reference Date
        term.Enter
        
        If isRemoving Then
            term.PutString 18, 50, reasonCode
            term.Enter
            
            If term.GetString(14, 22, 6) = "Fatura" Then term.SendPFKey 3
        Else
            ' Applying retention
            If term.GetString(14, 22, 6) = "FATURA" Then
                targetSheet.Cells(currentCell.Row, "F").Value = "Invoice Already Paid"
            ElseIf term.GetString(18, 22, 8) = "CONFIRME" Then
                ' Skip
            Else
                term.PutString 13, 38, reasonCode
                term.Enter
                term.PutString 17, 57, targetSheet.Cells(currentCell.Row, "E").Value
                term.Enter
                term.PutString 22, 78, "s"
                term.Enter
                targetSheet.Cells(currentCell.Row, "F").Value = "Retained for: " & reasonCode
            End If
        End If
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox IIf(isRemoving, "Retentions removed.", "Retentions applied."), vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ReopenInvoice
' Refactored from: reabre44 (Screen 44)
' Purpose: Navigates to Screen 44 to reopen/rebill an invoice reference.
' -------------------------------------------------------------------------
Public Sub ReopenInvoice()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    customerId = InputBox("Enter the 8-digit Customer ID:", "Customer ID")
    If customerId = "" Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "44"
        term.Enter
        
        term.PutString 10, 37, Format(customerId, "00000000")
        term.PutString 11, 37, Format(currentCell.Value, "000000")
        term.Enter
        term.Enter ' Bypass warnings
        
        ' Input rebilling parameters (assumes action code 'ac' + Col C value)
        term.PutString 10, 22, "ac" & targetSheet.Cells(currentCell.Row, "C").Value
        term.PutString 13, 15, targetSheet.Cells(currentCell.Row, "H").Value
        term.Enter
        
        term.PutString 23, 67, "s" ' Confirm
        term.Enter
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Invoices Reopened successfully.", vbInformation
End Sub