Attribute VB_Name = "Terminal_Special_Operations"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractIPCAData
' Refactored from: dadosIPCAmensal
' Purpose: Navigates to Screen 69 > 14 (Economic Indices) to extract IPCA 
'          history values, handling pagination properly.
' -------------------------------------------------------------------------
Public Sub ExtractIPCAData()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim excelRow As Integer, termRow As Integer
    Dim stopFlag As String, extractedValue As String
    Dim startTime As Double, endTime As Double
    
    startTime = Timer
    Set targetSheet = ActiveSheet
    
    Helper_Core.TogglePerformance False
    term.Connect
    
    ' Setup Screen 69 > 14
    term.PutString 22, 11, "69"
    term.Enter
    term.PutString 24, 11, "14"
    term.Enter
    
    ' Parameter setup (c1 = consult)
    term.PutString 4, 16, "c1"
    term.Enter
    term.PutString 13, 62, "ipca"
    term.PutString 14, 62, targetSheet.Cells(1, 2).Value ' Uses start date from cell B1
    term.Enter
    
    excelRow = 2
    termRow = 6
    
    ' Extraction Loop
    Do
        stopFlag = Trim(term.GetString(termRow, 21, 3))
        If stopFlag = "OTN" Then Exit Do ' End of IPCA records
        
        extractedValue = Trim(term.GetString(termRow, 48, 7))
        
        If extractedValue <> "" Then
            targetSheet.Cells(excelRow, 1).Value = "'" & extractedValue ' Prevent Excel auto-formatting
            excelRow = excelRow + 1
        End If
        
        termRow = termRow + 1
        
        ' Handle Pagination (Screen limit is row 21)
        If termRow > 21 Then
            term.Enter
            termRow = 6
        End If
    Loop
    
    term.SendPFKey 3
    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "IPCA Data Extracted! Time: " & Format((endTime - startTime), "0.00") & " seconds", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: CheckWaterOutages
' Refactored from: faltaDagua
' Purpose: Checks Screen 12 (Water Outages) for specific registrations 
'          and extracts the start, end, and normalization dates.
' -------------------------------------------------------------------------
Public Sub CheckWaterOutages()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim dateStart As String, dateEnd As String, dateNorm As String, reason As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub
    
    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3
    
    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        If targetSheet.Cells(currentCell.Row, "E").Value = "Sem registro de manobra" Then
            term.PutString 22, 11, "12" & Format(currentCell.Value, "00000000")
            term.Enter
            
            If term.GetString(19, 29, 4) = "NADA" Then
                targetSheet.Cells(currentCell.Row, "E").Value = "Sem registro de manobra"
            Else
                dateStart = Trim(term.GetString(13, 29, 16))
                dateEnd = Trim(term.GetString(13, 64, 16))
                dateNorm = Trim(term.GetString(14, 29, 16))
                reason = Trim(term.GetString(17, 16, 65))
                
                targetSheet.Cells(currentCell.Row, "E").Value = "Manobra Registrada"
                targetSheet.Cells(currentCell.Row, "F").Value = dateStart & " | " & dateEnd & " | " & dateNorm & " | " & reason
            End If
            
            term.SendPFKey 3
        End If
    Next currentCell
    
    Helper_Core.TogglePerformance True
    MsgBox "Water Outage Check Complete.", vbInformation
End Sub