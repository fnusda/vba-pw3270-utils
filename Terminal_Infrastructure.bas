Attribute VB_Name = "Terminal_Infrastructure"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: DeploySewerConnection
' Refactored from: t36_implantar / t37_implEsg
' Purpose: Navigates Screens 36 and 37 to implant a new sewer/water 
'          connection based on protocols and categories defined in Excel.
' -------------------------------------------------------------------------
Public Sub DeploySewerConnection()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim customerId As String, protocolId As String
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
        
        customerId = Format(currentCell.Value, "00000000")
        protocolId = Trim(targetSheet.Cells(currentCell.Row, "E").Value) ' Assume Protocol is in E
        
        ' Skip empty or invalid rows
        If customerId <> "00000000" And targetSheet.Cells(currentCell.Row, "G").Value = "" Then
            
            ' --- Phase 1: Screen 36 (Property / Routing Details) ---
            term.PutString 22, 11, "36"
            term.Enter
            
            term.PutString 13, 40, customerId
            term.Enter
            
            ' Validate the protocol
            term.PutString 11, 30, protocolId
            term.Enter
            
            If term.GetString(9, 24, 4) = "PROT" Then
                targetSheet.Cells(currentCell.Row, "G").Value = "Invalid Protocol"
                term.SendPFKey 3
                GoTo NextIteration
            End If
            
            ' Save and proceed
            term.PutString 23, 78, "s"
            term.Enter
            
            ' --- Phase 2: Screen 37 (Connection Specifications) ---
            term.Enter ' Advance to 37 natively from 36's save
            
            ' Check if already connected to sewer
            If term.GetString(9, 45, 6) = "LIGADA" Then
                targetSheet.Cells(currentCell.Row, "G").Value = "Already Connected"
                term.SendPFKey 3
                GoTo NextIteration
            End If
            
            ' Update Connection Status (e.g., Code 1 for Active, Code 3 for Sewer)
            term.PutString 19, 37, "1"
            term.PutString 20, 37, targetSheet.Cells(currentCell.Row, "C").Value ' Economies count
            term.PutString 21, 37, "3" ' Sewer identifier
            term.PutString 22, 42, Format(Date, "ddmmyyyy") ' Execution Date
            term.Enter
            
            ' --- Validation Bypass Loop ---
            BypassTerminalWarnings term
            
            ' Confirm and Save Deployment
            term.PutString 23, 70, "s"
            term.Enter
            
            targetSheet.Cells(currentCell.Row, "G").Value = "Deployed Successfully"
            term.SendPFKey 3
        End If

NextIteration:
    Next currentCell

    Helper_Core.TogglePerformance True
    endTime = Timer
    MsgBox "Sewer Deployments Finished! Time: " & Int((endTime - startTime) / 60) & " min " & Format((endTime - startTime) Mod 60, "0.00") & " seconds", vbInformation, "Completed"
End Sub

' -------------------------------------------------------------------------
' Subroutine: BypassTerminalWarnings (Helper)
' Purpose: Checks for standard system warnings (ATEN, DATA, MUDA) on save 
'          screens and resolves them cleanly.
' -------------------------------------------------------------------------
Private Sub BypassTerminalWarnings(ByRef term As clsTerminalMacros)
    ' 1. Standard "Attention" confirmation
    If term.GetString(17, 24, 4) = "ATEN" Then
        term.Enter
        term.PutString 17, 52, "s"
        term.Enter
    End If
    
    ' 2. "Print Term?" prompt
    If term.GetString(10, 25, 6) = "EMITIR" Then
        term.PutString 10, 63, "n"
        term.Enter
    End If
    
    ' 3. Category Change prompt (Requires clearing a specific field)
    If term.GetString(1, 2, 4) = "MUDA" Then
        term.PutString 20, 65, "        "
        term.Enter
    End If
End Sub