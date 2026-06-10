Attribute VB_Name = "Terminal_Queries"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: QueryCustomerRegistration
' Refactored from: a_t3ConsultaCad / am_t3ConsultaCad
' Purpose: Queries Screen 03 (Registration) for a list of customer IDs in 
'          Column A. Extracts specific registration data into Column B 
'          while handling potential terminal warning prompts.
' -------------------------------------------------------------------------
Public Sub QueryCustomerRegistration()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range
    Dim currentCell As Range
    Dim lastRow As Long
    Dim customerId As String
    Dim extractedValue As String
    Dim startTime As Double
    Dim endTime As Double

    ' 1. Initialization and Data Setup
    startTime = Timer
    Set targetSheet = ActiveSheet
    
    ' Utilize the Helper_Core function to find the last row dynamically
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")

    ' Prevent execution if the sheet is empty
    If lastRow < 2 Then
        MsgBox "No customer IDs found in Column A.", vbExclamation, "No Data"
        Exit Sub
    End If

    Set dataRange = targetSheet.Range("A2:A" & lastRow)

    ' Optimize Excel speed by turning off screen updates during execution
    Helper_Core.TogglePerformance False

    ' 2. Terminal Connection
    term.Connect
    term.SendPFKey 3 ' Ensure the terminal is at the home screen before starting

    ' 3. Main Processing Loop
    For Each currentCell In dataRange
        ' Update the Excel Status Bar using the Helper_Core function
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count

        ' Ensure the customer ID is strictly 8 digits long to avoid terminal errors
        customerId = Format(currentCell.Value, "00000000")

        ' Navigate to Screen 03 (Registration) using the specific customer ID
        term.PutString 22, 11, "03" & customerId
        term.Enter

        ' Check for an "ATEN" (Attention/Warning) message at row 15, column 26
        If term.GetString(15, 26, 4) = "ATEN" Then
            ' Bypass the warning by sending an "n" (No) response
            term.PutString 16, 72, "n"
            term.Enter
        End If

        ' Extract the target data from row 7, column 55 (length 11)
        extractedValue = Trim(term.GetString(7, 55, 11))

        ' Clean the extracted string: remove spaces, hyphens, and replace underscores with 'X'
        extractedValue = Replace(extractedValue, " ", "")
        extractedValue = Replace(extractedValue, "-", "")
        extractedValue = Replace(extractedValue, "_", "X")

        ' Write the cleaned value to Column B
        targetSheet.Cells(currentCell.Row, "B").Value = extractedValue

        ' Return to the previous terminal screen to prepare for the next iteration (F3)
        term.SendPFKey 3
    Next currentCell

    ' 4. Clean up and Exit
    Helper_Core.TogglePerformance True ' Restore Excel screen updates
    endTime = Timer

    ' Display execution metrics
    MsgBox "Process Finished! Execution time: " & _
           Int((endTime - startTime) / 60) & " min " & _
           Format((endTime - startTime) Mod 60, "0.00") & " seconds", _
           vbInformation, "Completed"
End Sub