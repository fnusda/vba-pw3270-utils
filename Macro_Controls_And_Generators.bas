Attribute VB_Name = "Macro_Controls_And_Generators"
Option Explicit

' --- Global Execution State Variables ---
Public IsCancellationRequested As Boolean
Public IsPaused As Boolean

' -------------------------------------------------------------------------
' Subroutine: RequestCancel / TogglePause
' Refactored from: BotaoCancelar / BotaoPausarContinuar
' Purpose: Used by UI buttons to safely interrupt or pause long loops.
' Note: You must add 'If IsCancellationRequested Then Exit Sub' inside your loops.
' -------------------------------------------------------------------------
Public Sub RequestCancel()
    IsCancellationRequested = True
End Sub

Public Sub TogglePause()
    IsPaused = Not IsPaused
    If IsPaused Then MsgBox "Macro Paused. Click again to resume.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: GenerateUniqueHexCodes
' Refactored from: i_idACad / GerarCodigosHexa
' Purpose: Generates unique 6-character hex codes using a Dictionary.
' -------------------------------------------------------------------------
Public Sub GenerateUniqueHexCodes()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim hexCode As String
    Dim hexDict As Object
    
    Set ws = ActiveSheet
    lastRow = Helper_Core.GetLastRow(ws, "K") ' Based on original code logic
    If lastRow < 2 Then Exit Sub
    
    Set hexDict = CreateObject("Scripting.Dictionary")
    Helper_Core.TogglePerformance False
    
    For i = 2 To lastRow
        Do
            ' Generate random number between 0 and FFFFFF (16777215)
            hexCode = Right("000000" & Hex(Int((16777215) * Rnd)), 6)
        Loop While hexDict.Exists(hexCode)
        
        hexDict.Add hexCode, 1
        ws.Cells(i, "J").Value = hexCode
    Next i
    
    Set hexDict = Nothing
    Helper_Core.TogglePerformance True
    MsgBox "Hex codes generated in Column J.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: GenerateTimeSequence
' Refactored from: PreencherHorasMinutos_Texto_Eficiente
' Purpose: Creates a full 24-hour array (0000 to 2359) and dumps it to Col D.
' -------------------------------------------------------------------------
Public Sub GenerateTimeSequence()
    Dim hourLoop As Integer, minuteLoop As Integer
    Dim arrayIndex As Long
    Dim timeArray(1 To 1440, 1 To 1) As String
    
    Helper_Core.TogglePerformance False
    arrayIndex = 0
    
    For hourLoop = 0 To 23
        For minuteLoop = 0 To 59
            arrayIndex = arrayIndex + 1
            timeArray(arrayIndex, 1) = Format(hourLoop, "00") & Format(minuteLoop, "00")
        Next minuteLoop
    Next hourLoop
    
    ' Apply to Column D
    With ActiveSheet
        .Columns("D").ClearContents
        .Columns("D").NumberFormat = "@"
        .Range("D1").Value = "Horas (HHMM)"
        .Range("D1").Font.Bold = True
        .Range("D2").Resize(1440, 1).Value = timeArray
    End With
    
    Helper_Core.TogglePerformance True
    MsgBox "24-Hour Time Sequence generated in Column D.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Function: ExtractLettersOnly
' Refactored from: ExtrairLetras
' Purpose: Cleans a string, returning only alphabetical characters.
' -------------------------------------------------------------------------
Public Function ExtractLettersOnly(ByVal inputTxt As String) As String
    Dim i As Integer
    Dim resultStr As String
    
    For i = 1 To Len(inputTxt)
        If Mid(inputTxt, i, 1) Like "[A-Za-z]" Then
            resultStr = resultStr & Mid(inputTxt, i, 1)
        End If
    Next i
    
    ExtractLettersOnly = resultStr
End Function