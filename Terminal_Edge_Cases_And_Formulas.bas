Attribute VB_Name = "Terminal_Edge_Cases_And_Formulas"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractPropertyRoutingID
' Refactored from: recuperaII / tela36
' Purpose: Navigates to Screen 36, extracts routing quadrants, cleans spaces, 
'          and concatenates them into a single property registration ID.
' -------------------------------------------------------------------------
Public Sub ExtractPropertyRoutingID()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim quadrant As String, quadricula As String, sector As String
    Dim block As String, lot As String, fullID As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "D")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("D2:D" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        Helper_Core.UpdateProgress currentCell.Row - 1, dataRange.Rows.Count
        
        term.PutString 22, 11, "36"
        term.Enter
        term.PutString 13, 40, Format(currentCell.Value, "00000000")
        term.SetCursor 16, 37
        term.Enter
        
        ' Extract and clean spatial data
        quadrant = Replace(term.GetString(11, 30, 5), " ", "")
        quadricula = Replace(term.GetString(12, 30, 5), " ", "")
        sector = Replace(term.GetString(13, 30, 5), " ", "")
        block = Replace(term.GetString(14, 30, 7), " ", "")
        lot = Replace(term.GetString(15, 30, 11), " ", "")
        
        fullID = quadrant & quadricula & sector & block & lot
        targetSheet.Cells(currentCell.Row, "E").Value = "'" & fullID ' Add apostrophe to treat as text
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Routing IDs Extracted.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: RegisterCommercialAction
' Refactored from: solicCPMouTHD (Screen 24)
' Purpose: Prompts the user to select between service 0131 or 0132 and 
'          registers the action accordingly.
' -------------------------------------------------------------------------
Public Sub RegisterCommercialAction()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim selectedService As String
    
    ' Prompt and validate user input
    selectedService = InputBox("Enter the desired service (0131 or 0132):", "Service Selection")
    If selectedService <> "0131" And selectedService <> "0132" Then
        MsgBox "Invalid service! Please run again and enter 0131 or 0132.", vbExclamation
        Exit Sub
    End If
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        term.PutString 22, 11, "24"
        term.Enter
        term.PutString 9, 28, Format(currentCell.Value, "00000000")
        term.PutString 10, 28, selectedService
        term.PutString 13, 25, "plano de acao comercial"
        
        If selectedService = "0131" Then
            term.PutString 15, 25, "exec 0131 + 0370 fat"
            term.Enter
            term.SendPFKey 2
            term.PutString 24, 31, "n"
            term.Enter
            term.PutString 24, 44, "10"
        ElseIf selectedService = "0132" Then
            term.PutString 14, 48, targetSheet.Cells(currentCell.Row, "B").Value
            term.PutString 15, 25, "agendar com cliente"
            term.Enter
            term.SendPFKey 2
            term.PutString 24, 31, "n"
            term.Enter
            term.PutString 24, 44, "00"
        End If
        
        ' Final confirmation
        term.Enter
        term.PutString 24, 69, "s"
        term.Enter
        targetSheet.Cells(currentCell.Row, "C").Value = term.GetString(14, 39, 17)
        term.Enter
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Commercial Actions Registered.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: FillBlankCellsFromBelow
' Refactored from: SubstituirValoresZeradosENulos
' Purpose: Loops backwards from the bottom row to row 2. If a cell in Cols C-F 
'          is empty or 0, it pulls the value from the row directly beneath it.
' -------------------------------------------------------------------------
Public Sub FillBlankCellsFromBelow()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim col As Integer
    
    On Error Resume Next
    Set ws = Workbooks("ajustaCODOPEpg").Sheets("matInsc")
    On Error GoTo 0
    
    If ws Is Nothing Then Exit Sub
    
    lastRow = Helper_Core.GetLastRow(ws, "C")
    Helper_Core.TogglePerformance False
    
    For i = lastRow - 1 To 2 Step -1
        For col = 3 To 6 ' Columns C through F
            If IsEmpty(ws.Cells(i, col).Value) Or ws.Cells(i, col).Value = "" Or ws.Cells(i, col).Value = 0 Then
                If Not IsEmpty(ws.Cells(i + 1, col).Value) And ws.Cells(i + 1, col).Value <> 0 Then
                    ws.Cells(i, col).Value = ws.Cells(i + 1, col).Value
                End If
            End If
        Next col
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Null values replaced from below.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: TransposeDataBlockToColumnA
' Refactored from: TransporDados
' Purpose: Flattens a massive horizontal grid (F2:BJ112) into a single column.
' -------------------------------------------------------------------------
Public Sub TransposeDataBlockToColumnA()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim sourceRange As Range, cell As Range
    Dim nextRow As Long
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("draft")
    Set wsDest = ThisWorkbook.Sheets("inscricaoImobiliaria")
    On Error GoTo 0
    
    If wsSource Is Nothing Or wsDest Is Nothing Then Exit Sub
    
    Set sourceRange = wsSource.Range("F2:BJ112")
    Helper_Core.TogglePerformance False
    
    For Each cell In sourceRange
        If Trim(cell.Value) <> "" Then
            nextRow = Helper_Core.GetLastRow(wsDest, "A") + 1
            wsDest.Range("A" & nextRow).Value = cell.Value
        End If
    Next cell
    
    Helper_Core.TogglePerformance True
    MsgBox "Data block transposed successfully.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: InjectFormattingFormulas
' Refactored from: completaCruzamento / preencheBairro / preencheQuadras
' Purpose: Injects hardcoded string-manipulation formulas into specific columns.
' -------------------------------------------------------------------------
Public Sub InjectFormattingFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Set ws = ActiveSheet
    lastRow = Helper_Core.GetLastRow(ws, "A")
    If lastRow < 2 Then Exit Sub
    
    Helper_Core.TogglePerformance False
    
    For i = 2 To lastRow
        ' Inject Crossings Formula into Col H
        ws.Cells(i, "H").Formula = "=TEXT(B" & i & ",""00000"") & TEXT(C" & i & ",""00000"") & D" & i & " & TEXT(E" & i & ",""00000"") & TEXT(F" & i & ",""00000"") & G" & i
        
        ' Inject Neighborhood Formula into Col D
        ws.Cells(i, "D").Formula = "=CONCAT(IF(LEN(A" & i & ")=1,""00""&A" & i & ",IF(LEN(A" & i & ")=2,""0""&A" & i & ",A" & i & ")),IF(LEN(B" & i & ")=1,""00""&B" & i & ",IF(LEN(B" & i & ")=2,""0""&B" & i & ",B" & i & ")),C" & i & ")"
        
        ' Inject Block (Quadra) Formula into Col D (Overrides previous if run together)
        ' ws.Cells(i, "D").Formula = "=TEXT(A" & i & ",""000"") & TEXT(B" & i & ",""000"") & TEXT(""02"",""00"") & TEXT(C" & i & ",""0000"")"
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Formatting formulas injected.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ClearFiltersAndRanges
' Refactored from: limp_dev / outras_funcoes
' Purpose: Safely clears AutoFilters and resets specific target zones.
' -------------------------------------------------------------------------
Public Sub ClearFiltersAndRanges()
    Dim wsAuto As Worksheet, wsData As Worksheet
    
    On Error Resume Next
    Set wsAuto = ThisWorkbook.Sheets("Automático")
    Set wsData = ThisWorkbook.Sheets("Dados")
    On Error GoTo 0
    
    Helper_Core.TogglePerformance False
    
    If Not wsAuto Is Nothing Then
        ' Reset Autofilter safely
        If wsAuto.AutoFilterMode Then wsAuto.AutoFilterMode = False
        
        ' Clear specific data regions
        If Helper_Core.GetLastRow(wsAuto, "J") >= 3 Then wsAuto.Range("J3:S" & Helper_Core.GetLastRow(wsAuto, "J")).ClearContents
        If Helper_Core.GetLastRow(wsAuto, "AA") >= 3 Then wsAuto.Range("AA3:AJ" & Helper_Core.GetLastRow(wsAuto, "AA")).ClearContents
        If Helper_Core.GetLastRow(wsAuto, "AQ") >= 3 Then wsAuto.Range("AQ3:AQ" & Helper_Core.GetLastRow(wsAuto, "AQ")).ClearContents
    End If
    
    If Not wsData Is Nothing Then
        wsData.Range("O10:S10, O12:S13, O15:S15, O17:S18, O22:P22").ClearContents
    End If
    
    Helper_Core.TogglePerformance True
    MsgBox "Filters and ranges cleared safely.", vbInformation
End Sub