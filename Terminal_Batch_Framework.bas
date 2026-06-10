Attribute VB_Name = "Terminal_Batch_Framework"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: RunResumableBatch
' Refactored from: buscaAnoEBairro / ConfereCliente
' Purpose: A universal wrapper for running long terminal queries in batches. 
'          It asks the user where to start, calculates the remaining lines, 
'          and restricts the loop. You can insert ANY terminal logic inside.
' -------------------------------------------------------------------------
Public Sub RunResumableBatch()
    Dim ws As Worksheet
    Dim lastRow As Long, startRow As Long, endRow As Long
    Dim maxPossible As Long, batchSize As Long
    Dim userInput As String
    Dim i As Long
    Dim term As New clsTerminalMacros
    
    Set ws = ActiveSheet
    lastRow = Helper_Core.GetLastRow(ws, "A")
    If lastRow < 2 Then Exit Sub
    
    ' 1. Determine Starting Point
    If MsgBox("Do you want to resume from the first empty row in Column Z?", vbYesNo + vbQuestion, "Resume Point") = vbYes Then
        startRow = Helper_Core.GetLastRow(ws, "Z")
        If startRow < 2 Then
            startRow = 2
        ElseIf Trim(ws.Cells(startRow, "Z").Value) <> "" Then
            startRow = startRow + 1
        End If
    Else
        startRow = 2
    End If
    
    If startRow > lastRow Then
        MsgBox "All rows have already been processed.", vbInformation
        Exit Sub
    End If
    
    ' 2. Determine Batch Size
    maxPossible = lastRow - startRow + 1
    userInput = InputBox("How many lines do you want to process in this batch?" & vbCrLf & _
                         "Remaining lines: " & maxPossible, "Batch Limit", maxPossible)
                         
    If userInput = "" Or Not IsNumeric(userInput) Then Exit Sub
    batchSize = CLng(userInput)
    If batchSize <= 0 Then Exit Sub
    
    endRow = startRow + batchSize - 1
    If endRow > lastRow Then endRow = lastRow

    ' 3. Execution Setup
    Helper_Core.TogglePerformance False
    term.Connect
    term.SendPFKey 3
    
    ' 4. The Processing Loop
    For i = startRow To endRow
        Helper_Core.UpdateProgress i - startRow, endRow - startRow + 1
        
        If Trim(ws.Cells(i, "A").Value) <> "" Then
            
            ' =========================================================
            ' INSERT YOUR SPECIFIC TERMINAL LOGIC HERE
            ' Example: Checking CODOPE on Screen 03
            ' =========================================================
            term.PutString 22, 11, "03"
            term.PutString 22, 24, Format(ws.Cells(i, "A").Value, "00000000")
            term.Enter
            
            If LCase(Trim(term.GetString(16, 26, 16))) = "deseja atualizar" Then
                term.PutString 16, 72, "n"
                term.Enter
            End If
            
            ' Extract data to Z, AA, AB
            ws.Cells(i, "Z").Value = term.GetString(13, 61, 4)
            ws.Cells(i, "AA").Value = term.GetString(13, 59, 2)
            ws.Cells(i, "AB").Value = IIf(term.GetString(12, 10, 2) = "  ", "Sem CODOPE", Trim(term.GetString(12, 10, 34)))
            
            term.SendPFKey 3
            ' =========================================================
            
        End If
    Next i

    Helper_Core.TogglePerformance True
    MsgBox "Batch of " & (endRow - startRow + 1) & " lines processed successfully!", vbInformation
End Sub