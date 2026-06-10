Attribute VB_Name = "Data_Formatting_Chain"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExecuteFullDataCleanup
' Refactored from: al_executar
' Purpose: Orchestrates the sequential execution of all data formatting macros.
' -------------------------------------------------------------------------
Public Sub ExecuteFullDataCleanup()
    Dim targetSheet As Worksheet
    Set targetSheet = ActiveSheet
    
    Helper_Core.TogglePerformance False ' Turn off screen updating for speed
    
    ' Execute the chain of formatting events
    Data_Preparation.RemoveAllMergedCells ' Assuming this is in the Data_Preparation module
    FormatColumnA targetSheet
    RemoveInvalidAndEmptyRows targetSheet
    Data_Preparation.DeleteEmptyColumns targetSheet
    InsertHeaders targetSheet
    ReorganizeColumns targetSheet
    
    Helper_Core.TogglePerformance True
    MsgBox "All formatting macros executed successfully!", vbInformation, "Process Finished"
End Sub

' -------------------------------------------------------------------------
' Subroutine: FormatColumnA
' Refactored from: aca_formataColA
' Purpose: Formats Column A to an 8-digit string, removing spaces and commas.
' -------------------------------------------------------------------------
Private Sub FormatColumnA(ByRef ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim rawValue As String, formattedValue As String
    Dim parts() As String
    
    lastRow = Helper_Core.GetLastRow(ws, "A")
    
    For i = 1 To lastRow
        rawValue = ws.Cells(i, 1).Value
        If rawValue <> "" Then
            ' Split by space and take the first part
            parts = Split(rawValue, " ")
            formattedValue = Replace(parts(0), ",", "")
            
            ' Pad with leading zeros to ensure 8 characters
            formattedValue = Right("00000000" & formattedValue, 8)
            
            ws.Cells(i, 1).NumberFormat = "@"
            ws.Cells(i, 1).Value = formattedValue
        End If
    Next i
End Sub

' -------------------------------------------------------------------------
' Subroutine: RemoveInvalidAndEmptyRows
' Refactored from: ac_removerLinhasInuteis
' Purpose: Deletes rows where Column A doesn't match a specific Regex pattern, 
'          or where columns B through T are completely empty.
' -------------------------------------------------------------------------
Private Sub RemoveInvalidAndEmptyRows(ByRef ws As Worksheet)
    Dim lastRow As Long, i As Long, j As Long
    Dim regex As Object
    Dim allColumnsEmpty As Boolean
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "^\d{4}[,.]\d{4}( \d{2}/\d{4}-\d)?$"
    regex.IgnoreCase = False
    regex.Global = False
    
    lastRow = Helper_Core.GetLastRow(ws, "A")
    
    ' Loop backward when deleting rows
    For i = lastRow To 1 Step -1
        If Not regex.Test(ws.Cells(i, 1).Value) Then
            ws.Rows(i).Delete Shift:=xlUp
        Else
            allColumnsEmpty = True
            ' Check columns B (2) through T (20)
            For j = 2 To 20
                If Trim(ws.Cells(i, j).Value) <> "" Then
                    allColumnsEmpty = False
                    Exit For
                End If
            Next j
            
            If allColumnsEmpty Then ws.Rows(i).Delete Shift:=xlUp
        End If
    Next i
    Set regex = Nothing
End Sub

' -------------------------------------------------------------------------
' Subroutine: InsertHeaders
' Refactored from: af_cabecalho
' Purpose: Inserts a new row 1 and applies standard headers.
' -------------------------------------------------------------------------
Private Sub InsertHeaders(ByRef ws As Worksheet)
    ws.Rows(1).Insert Shift:=xlDown
    
    ws.Cells(1, 1).Value = "Matrícula"
    ws.Cells(1, 2).Value = "Referência"
    ws.Cells(1, 3).Value = "Valor Informado"
    ws.Cells(1, 4).Value = "Valor Processado"
    ws.Cells(1, 5).Value = "Diferença"
    ws.Cells(1, 6).Value = "Motivo"
End Sub

' -------------------------------------------------------------------------
' Subroutine: ReorganizeColumns
' Refactored from: ag_moveReferencias through ak_moveMotivo
' Purpose: Consolidates the logic of moving/shifting values to the left 
'          and concatenating the 'Motivo' columns (M through Q).
' -------------------------------------------------------------------------
Private Sub ReorganizeColumns(ByRef ws As Worksheet)
    Dim lastRow As Long, i As Long
    Dim concatMotivo As String
    
    ' Find the absolute last row across the relevant columns
    lastRow = ws.Cells.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If lastRow < 2 Then Exit Sub
    
    For i = 2 To lastRow
        ' Clean whitespace across the working range first
        Dim j As Integer
        For j = 1 To 20
            ws.Cells(i, j).Value = Trim(ws.Cells(i, j).Value)
        Next j
        
        ' 1. Move Referência (C -> B)
        If ws.Cells(i, 2).Value = "" And ws.Cells(i, 3).Value <> "" Then
            ws.Cells(i, 2).Value = ws.Cells(i, 3).Value
            ws.Cells(i, 3).ClearContents
        End If
        
        ' 2. Move Valor Informado (D/E/F -> C)
        If ws.Cells(i, 3).Value = "" Then
            If ws.Cells(i, 4).Value <> "" Then: ws.Cells(i, 3).Value = ws.Cells(i, 4).Value: ws.Cells(i, 4).ClearContents
            ElseIf ws.Cells(i, 5).Value <> "" Then: ws.Cells(i, 3).Value = ws.Cells(i, 5).Value: ws.Cells(i, 5).ClearContents
            ElseIf ws.Cells(i, 6).Value <> "" Then: ws.Cells(i, 3).Value = ws.Cells(i, 6).Value: ws.Cells(i, 6).ClearContents
            End If
        End If
        
        ' 3. Move Valor Processado (G/H/I -> D)
        If ws.Cells(i, 4).Value = "" Then
            If ws.Cells(i, 7).Value <> "" Then: ws.Cells(i, 4).Value = ws.Cells(i, 7).Value: ws.Cells(i, 7).ClearContents
            ElseIf ws.Cells(i, 8).Value <> "" Then: ws.Cells(i, 4).Value = ws.Cells(i, 8).Value: ws.Cells(i, 8).ClearContents
            ElseIf ws.Cells(i, 9).Value <> "" Then: ws.Cells(i, 4).Value = ws.Cells(i, 9).Value: ws.Cells(i, 9).ClearContents
            End If
        End If
        
        ' 4. Move Diferença (I/J/K -> E)
        If ws.Cells(i, 5).Value = "" Then
            If ws.Cells(i, 9).Value <> "" Then: ws.Cells(i, 5).Value = ws.Cells(i, 9).Value: ws.Cells(i, 9).ClearContents
            ElseIf ws.Cells(i, 10).Value <> "" Then: ws.Cells(i, 5).Value = ws.Cells(i, 10).Value: ws.Cells(i, 10).ClearContents
            ElseIf ws.Cells(i, 11).Value <> "" Then: ws.Cells(i, 5).Value = ws.Cells(i, 11).Value: ws.Cells(i, 11).ClearContents
            End If
        End If
        
        ' 5. Concatenate Motivo (M, N, O, P, Q -> F)
        concatMotivo = ""
        If ws.Cells(i, 13).Value <> "" Then concatMotivo = concatMotivo & ws.Cells(i, 13).Value & " "
        If ws.Cells(i, 14).Value <> "" Then concatMotivo = concatMotivo & ws.Cells(i, 14).Value & " "
        If ws.Cells(i, 15).Value <> "" Then concatMotivo = concatMotivo & ws.Cells(i, 15).Value & " "
        If ws.Cells(i, 16).Value <> "" Then concatMotivo = concatMotivo & ws.Cells(i, 16).Value & " "
        If ws.Cells(i, 17).Value <> "" Then concatMotivo = concatMotivo & ws.Cells(i, 17).Value & " "
        
        ws.Cells(i, 6).Value = Trim(concatMotivo)
    Next i
    
    ' Delete all leftover columns to the right of F (Column 6)
    ws.Columns("G:Z").Delete
End Sub