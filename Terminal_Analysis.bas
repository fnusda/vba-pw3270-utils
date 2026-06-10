Attribute VB_Name = "Terminal_Analysis"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: AnalyzeZeroConsumption
' Refactored from: analiseCons0
' Purpose: Analyzes water meters with zero consumption. Navigates to Screen 08, 
'          extracts forecasted vs. executed dates, and calculates service delays.
' -------------------------------------------------------------------------
Public Sub AnalyzeZeroConsumption()
    Dim term As New clsTerminalMacros
    Dim targetBook As Workbook
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim forecastDate As Date, executionDate As Date
    Dim daysDelayed As Long
    Dim serviceCode As String, protocol As String
    
    ' Note: Original macro hardcoded a specific path. 
    ' Recommendation: Use Application.GetOpenFilename for flexibility in the future.
    On Error Resume Next
    Set targetBook = Workbooks.Open("D:\Documents\randomizer\consumo_Zero.xlsm")
    On Error GoTo ErrorHandler
    
    Set targetSheet = targetBook.Worksheets("ocorrencias")
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False
    
    term.Connect
    term.SendPFKey 3
    
    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        ' --- 1. Screen 03: Basic Info ---
        term.PutString 22, 11, "03"
        term.PutString 22, 24, Format(currentCell.Value, "00000000")
        term.Enter
        
        targetSheet.Cells(currentCell.Row, "C").Value = term.GetString(13, 10, 3) ' Local Value
        targetSheet.Cells(currentCell.Row, "D").Value = term.GetString(10, 57, 2) ' Group Value
        term.Enter
        
        ' --- 2. Screen 08: Service Analysis ---
        term.PutString 22, 11, "08"
        term.PutString 22, 24, Format(currentCell.Value, "00000000")
        term.Enter
        
        ' Check Row 19 for Service 8125 or 8004 (We can loop rows 9, 14, 19 to make this dynamic)
        Dim checkRows As Variant, r As Variant
        checkRows = Array(19, 14, 9) ' Check from bottom up as in original macro
        
        For Each r In checkRows
            serviceCode = term.GetString(r, 3, 4)
            
            If serviceCode = "8125" Or serviceCode = "8004" Then
                targetSheet.Cells(currentCell.Row, "G").Value = serviceCode
                protocol = term.GetString(r, 8, 17)
                targetSheet.Cells(currentCell.Row, "H").Value = Format(protocol, "00000000000000000")
                
                ' Parse Dates (Uses Private Helper Function below)
                ' The column offset for dates varies slightly by service code in the original script
                Dim forecastCol As Integer
                forecastCol = IIf(serviceCode = "8125", 26, 10) 
                
                forecastDate = ParseTerminalDate(term, CInt(r), forecastCol)
                executionDate = ParseTerminalDate(term, CInt(r), 33)
                
                targetSheet.Cells(currentCell.Row, "I").Value = Format(forecastDate, "dd/mm/yy")
                
                ' Business Logic: Calculate Delays
                If term.GetString(r, 33, 6) = "000000" Then
                    daysDelayed = DateDiff("d", forecastDate, Date)
                    targetSheet.Cells(currentCell.Row, "J").Value = "Service delayed by " & daysDelayed & " days"
                    targetSheet.Cells(currentCell.Row, "K").Value = "No"
                Else
                    daysDelayed = DateDiff("d", forecastDate, executionDate)
                    
                    targetSheet.Cells(currentCell.Row, "J").Value = term.GetString(r + 1, 49, 30) ' Result text
                    
                    If executionDate > forecastDate Then
                        targetSheet.Cells(currentCell.Row, "K").Value = "Delayed by " & daysDelayed & " days"
                    ElseIf executionDate = forecastDate Then
                        targetSheet.Cells(currentCell.Row, "K").Value = "On Time"
                    Else
                        targetSheet.Cells(currentCell.Row, "K").Value = "Completed " & Abs(daysDelayed) & " days early"
                    End If
                End If
                Exit For ' Stop checking rows once we find the relevant service
            End If
        Next r
        
        term.SendPFKey 3
    Next currentCell

CleanUp:
    Helper_Core.TogglePerformance True
    MsgBox "Zero Consumption Analysis Complete.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' -------------------------------------------------------------------------
' Function: ParseTerminalDate (Helper)
' Purpose: Replaces repetitive date extraction math (Year/Month/Day splitting).
' -------------------------------------------------------------------------
Private Function ParseTerminalDate(ByRef term As clsTerminalMacros, ByVal rowIdx As Integer, ByVal colIdx As Integer) As Date
    Dim d As Integer, m As Integer, y As Integer
    ' Terminal format assumed: DD at colIdx, MM at colIdx+2, YY at colIdx+4
    d = CInt(term.GetString(rowIdx, colIdx, 2))
    m = CInt(term.GetString(rowIdx, colIdx + 2, 2))
    y = CInt(term.GetString(rowIdx, colIdx + 4, 2))
    
    ' Handle blank/zero dates to prevent Type Mismatch
    If d = 0 Or m = 0 Then
        ParseTerminalDate = DateSerial(1900, 1, 1)
    Else
        ' Assumes 2000s for 2-digit years
        ParseTerminalDate = DateSerial(2000 + y, m, d)
    End If
End Function

' -------------------------------------------------------------------------
' Subroutine: AdjustRouting
' Refactored from: ajustaRota
' Purpose: Navigates to Screen 36 and updates the meter reading route.
' -------------------------------------------------------------------------
Public Sub AdjustRouting()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    Helper_Core.TogglePerformance False
    
    term.Connect
    term.SendPFKey 3
    
    For Each currentCell In dataRange
        term.PutString 22, 11, "36"
        term.Enter
        
        term.PutString 13, 40, Format(currentCell.Value, "00000000")
        term.SetCursor 16, 37
        term.Enter
        
        ' Update routing code (original had hardcoded "04880")
        term.PutString 19, 30, "04880"
        term.Enter
        term.PutString 23, 78, "s" ' Confirm
        term.Enter
        term.SendPFKey 3
        
        ' Clear alerts if present
        If term.GetString(1, 2, 3) = "MUD" Then
            term.SendPFKey 1
            term.SendPFKey 3
        End If
        
        targetSheet.Cells(currentCell.Row, "G").Value = "Route Updated"
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Routing adjustments finished.", vbInformation
End Sub