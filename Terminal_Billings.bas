Attribute VB_Name = "Terminal_Billing"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ProcessInstallmentBilling
' Refactored from: k_faturarTela41
' Purpose: Navigates to Screen 41 to apply billing installments to accounts.
' -------------------------------------------------------------------------
Public Sub ProcessInstallmentBilling()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "T")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("T2:T" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' Validate business rules: Service must be > 3000 and < 4000
        Dim serviceCode As Long
        serviceCode = Val(targetSheet.Cells(currentCell.Row, "L").Value)
        
        If serviceCode > 3000 And serviceCode < 4000 And Val(currentCell.Value) <> 0 Then
            term.PutString 22, 11, "41"
            term.Enter
            
            ' Input billing parameters
            term.PutString 12, 45, Format(currentCell.Value, "00000000") ' Customer ID
            term.PutString 13, 45, Format(serviceCode, "0000")          ' Service Code
            term.PutString 14, 45, Format(targetSheet.Cells(currentCell.Row, "N").Value, "00000000") ' Execution Date
            term.PutString 16, 45, "01"                                 ' Initial Installment
            term.PutString 17, 45, Format(targetSheet.Cells(currentCell.Row, "P").Value, "00")       ' Final Installment
            term.Enter
            
            ' Check for existing installments
            If term.GetString(10, 27, 4) = "PARC" Then
                targetSheet.Cells(currentCell.Row, "AD").Value = "Already billed"
            Else
                term.Enter
                term.PutString 19, 51, "s" ' Confirm
                term.Enter
                targetSheet.Cells(currentCell.Row, "AD").Value = Trim(term.GetString(13, 29, 23)) ' Success message
            End If
            
            term.SendPFKey 3
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Installment Billing Completed.", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: ProcessRefunds
' Refactored from: t35_01_dev
' Purpose: Navigates to Screen 35 to process financial refunds (devoluções).
' -------------------------------------------------------------------------
Public Sub ProcessRefunds()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "35"
        term.Enter
        term.PutString 22, 57, "01" ' Operation: Refund
        term.Enter
        
        ' Input Customer ID and Target Reference Month
        term.PutString 12, 47, Format(currentCell.Value, "00000000")
        term.PutString 13, 47, "022025" ' Hardcoded in original, consider making dynamic
        term.Enter
        
        ' Check if a refund already exists for this reference
        If term.GetString(12, 2, 4) = "****" Then
            targetSheet.Cells(currentCell.Row, "F").Value = "Refund already exists for this reference"
        Else
            ' Proceed with refund entry
            Dim targetCol As Integer
            targetCol = Val(targetSheet.Cells(currentCell.Row, "E").Value)
            
            ' Apply refund values
            If targetCol > 0 Then
                term.PutString 11, targetCol, targetSheet.Cells(currentCell.Row, "C").Value
            End If
            
            term.PutString 16, 48, "n084" ' Refund reason code
            term.Enter
            term.PutString 12, 47, "n"
            term.Enter
            
            targetSheet.Cells(currentCell.Row, "F").Value = "Refund Processed"
            
            ' Final confirmation
            term.PutString 13, 55, "s"
            term.Enter
        End If
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Refund Processing Completed.", vbInformation, "Completed"
End Sub