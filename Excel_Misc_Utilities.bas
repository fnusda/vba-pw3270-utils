Attribute VB_Name = "Excel_Misc_Utilities"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: StopAllMacros
' Refactored from: paraMacros / PararExecucao
' Purpose: Forces an immediate hard stop of VBA execution using an error raise.
' -------------------------------------------------------------------------
Public Sub StopAllMacros()
    On Error Resume Next
    Err.Raise vbObjectError + 9999, "User-defined error", "Execution forcefully stopped by user."
    On Error GoTo 0
End Sub

' -------------------------------------------------------------------------
' Function: NumberToText
' Refactored from: Num2Text
' Purpose: Safely converts doubles to formatted strings.
' -------------------------------------------------------------------------
Public Function NumberToText(ByVal val As Double) As String
    If val = 0 Then
        NumberToText = "0,00"
    Else
        NumberToText = Format(val, "##,##0.00")
    End If
End Function

' -------------------------------------------------------------------------
' Subroutine: RemoveDotsColC
' Refactored from: apaga_pto
' Purpose: Strips all periods (.) from Column C.
' -------------------------------------------------------------------------
Public Sub RemoveDotsColC()
    ActiveSheet.Columns("C:C").Replace What:=".", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End Sub

' -------------------------------------------------------------------------
' Subroutine: Route8135Data
' Refactored from: h_importarDe8135 / enviarParaBaixa
' Purpose: Moves specific data blocks from the '8135' sheet to 'preparaImplantacao'
'          and 'tela19' sheets, enforcing text formats.
' -------------------------------------------------------------------------
Public Sub Route8135Data()
    Dim wsSource As Worksheet, wsImpl As Worksheet, wsT19 As Worksheet
    Dim lastRow As Long, i As Long
    Dim targetRowT19 As Long
    Dim valR As String, valS As String

    Set wsSource = ThisWorkbook.Sheets("8135")
    Set wsImpl = ThisWorkbook.Sheets("preparaImplantacao")
    Set wsT19 = ThisWorkbook.Sheets("tela19")
    
    lastRow = Helper_Core.GetLastRow(wsSource, "L")
    targetRowT19 = Helper_Core.GetLastRow(wsT19, "A") + 1
    
    Helper_Core.TogglePerformance False

    For i = 2 To lastRow
        ' --- 1. Route to preparaImplantacao ---
        wsImpl.Cells(i, "A").Value = Format(wsSource.Cells(i, "L").Value, "0000")
        wsImpl.Cells(i, "B").Value = Format(wsSource.Cells(i, "K").Value, "00")
        wsImpl.Cells(i, "G").Value = Format(wsSource.Cells(i, "AC").Value, "00000000")

        valR = CStr(wsSource.Cells(i, "R").Value)
        wsImpl.Cells(i, "P").Value = Format(Left(valR, 3), "000")
        wsImpl.Cells(i, "Q").Value = Format(Mid(valR, 4, 3), "000")
        wsImpl.Cells(i, "R").Value = Format(Mid(valR, 7, 3), "000")
        wsImpl.Cells(i, "S").Value = Format(Mid(valR, 10, 4), "0000")
        wsImpl.Cells(i, "T").Value = Right("000000" & Right(valR, 5), 6)

        valS = CStr(wsSource.Cells(i, "S").Value)
        wsImpl.Cells(i, "K").Value = Format(Left(valS, 2), "00")
        wsImpl.Cells(i, "L").Value = Format(Mid(valS, 3, 2), "00")
        wsImpl.Cells(i, "M").Value = Format(Mid(valS, 5, 3), "000")
        wsImpl.Cells(i, "N").Value = Format(Right(valS, 5), "00000")

        wsImpl.Cells(i, "C").Value = wsImpl.Cells(i, "P").Value & wsImpl.Cells(i, "Q").Value & _
                                     wsImpl.Cells(i, "R").Value & wsImpl.Cells(i, "S").Value & wsImpl.Cells(i, "T").Value
                                     
        wsImpl.Cells(i, "F").Value = wsImpl.Cells(i, "K").Value & wsImpl.Cells(i, "L").Value & _
                                     wsImpl.Cells(i, "M").Value & wsImpl.Cells(i, "N").Value
                                     
        ' --- 2. Route to tela19 (if Col T is not empty) ---
        If wsSource.Cells(i, "T").Value <> "" Then
            wsT19.Cells(targetRowT19, "A").Value = wsSource.Cells(i, "Q").Value
            wsT19.Cells(targetRowT19, "B").Value = wsSource.Cells(i, "T").Value
            wsT19.Cells(targetRowT19, "C").Value = wsSource.Cells(i, "P").Value
            wsT19.Cells(targetRowT19, "D").Value = wsSource.Cells(i, "N").Value
            wsT19.Cells(targetRowT19, "E").Value = wsSource.Cells(i, "O").Value
            targetRowT19 = targetRowT19 + 1
        End If
    Next i

    Helper_Core.TogglePerformance True
    MsgBox "Data routing complete.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: MergeValidSheets
' Refactored from: concatenar
' Purpose: Merges rows across all sheets into 'Ajustada' if Col A matches regex.
' -------------------------------------------------------------------------
Public Sub MergeValidSheets()
    Dim ws As Worksheet, wsAjustada As Worksheet
    Dim lastRow As Long, targetRow As Long
    Dim cell As Range
    Dim regex As Object
    
    On Error Resume Next
    Set wsAjustada = ThisWorkbook.Worksheets("Ajustada")
    If wsAjustada Is Nothing Then
        Set wsAjustada = ThisWorkbook.Worksheets.Add
        wsAjustada.Name = "Ajustada"
    Else
        wsAjustada.Cells.Clear
    End If
    On Error GoTo 0
    
    targetRow = 1
    wsAjustada.Range("A1:K1").Value = Array("MATRÍCULA", "CATEGORIA", "TOTAL ECON. ÁGUA", "ANORM. ANTERIOR", "ANORMALIDADE", "DATA LEITURA", "MÉDIA", "MEDIDO", "MOTIVO", "FISCALIZAÇÃO", "ANOTAÇÕES")
    targetRow = targetRow + 1
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d{4}\.\d{4}$"
    regex.IgnoreCase = True
    
    Helper_Core.TogglePerformance False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Ajustada" Then
            lastRow = Helper_Core.GetLastRow(ws, "A")
            For Each cell In ws.Range("A1:A" & lastRow)
                If regex.Test(cell.Value) Then
                    wsAjustada.Range("A" & targetRow & ":K" & targetRow).Value = ws.Range("A" & cell.Row & ":K" & cell.Row).Value
                    wsAjustada.Cells(targetRow, 1).Value = Replace(cell.Value, ".", "")
                    targetRow = targetRow + 1
                End If
            Next cell
        End If
    Next ws
    
    Set regex = Nothing
    Helper_Core.TogglePerformance True
    MsgBox "Sheets concatenated successfully!", vbInformation
End Sub