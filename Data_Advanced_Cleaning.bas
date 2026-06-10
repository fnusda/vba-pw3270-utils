Attribute VB_Name = "Data_Advanced_Cleaning"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: SplitLargeNumbers
' Refactored from: formataDados
' Purpose: Loops through specific columns (B, D, F, H, J). If a number is 
'          longer than 9 digits, it keeps the first 9 and moves the rest 
'          to the adjacent right column.
' -------------------------------------------------------------------------
Public Sub SplitLargeNumbers()
    Dim ws As Worksheet
    Dim colLetter As Variant
    Dim targetColumns As Variant
    Dim cell As Range
    Dim lastRow As Long
    Dim cellValue As String
    
    Set ws = ActiveSheet
    targetColumns = Array("B", "D", "F", "H", "J")
    
    Helper_Core.TogglePerformance False
    
    For Each colLetter In targetColumns
        lastRow = Helper_Core.GetLastRow(ws, CStr(colLetter))
        If lastRow >= 1 Then
            For Each cell In ws.Range(colLetter & "1:" & colLetter & lastRow)
                cellValue = Trim(cell.Value)
                
                If Len(cellValue) > 9 Then
                    ' Keep the first 9 digits, move the rest to the offset column
                    cell.Value = Left(cellValue, 9)
                    cell.Offset(0, 1).Value = Mid(cellValue, 10)
                End If
            Next cell
        End If
    Next colLetter
    
    Helper_Core.TogglePerformance True
    MsgBox "Large numbers split successfully.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: GenerateYearlyDates
' Refactored from: criaDatas
' Purpose: Fills Column A with daily dates for the year 2025 across all 
'          worksheets except 'colaboradoresPG'.
' -------------------------------------------------------------------------
Public Sub GenerateYearlyDates()
    Dim ws As Worksheet
    Dim startDate As Date, endDate As Date, currentDate As Date
    Dim rowIdx As Integer
    
    startDate = DateValue("01/01/2025")
    endDate = DateValue("31/12/2025")
    
    Helper_Core.TogglePerformance False
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "colaboradoresPG" Then
            rowIdx = 4 ' Start on row 4 as per original logic
            currentDate = startDate
            
            Do While currentDate <= endDate
                ws.Cells(rowIdx, 1).Value = Format(currentDate, "DDMMYYYY")
                ws.Cells(rowIdx, 1).NumberFormat = "@" ' Format as text
                currentDate = currentDate + 1
                rowIdx = rowIdx + 1
            Loop
        End If
    Next ws
    
    Helper_Core.TogglePerformance True
    MsgBox "Dates generated across all sheets!", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: DeleteConditionallyEmptyRows
' Refactored from: excluirLinhasVazias
' Purpose: Scans Column A to find the last row, then deletes the row only 
'          if columns B, C, and D are completely empty. Uses bulk deletion.
' -------------------------------------------------------------------------
Public Sub DeleteConditionallyEmptyRows()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim rowsToDelete As Range
    
    Set ws = ActiveSheet
    lastRow = Helper_Core.GetLastRow(ws, "A")
    If lastRow < 2 Then Exit Sub
    
    Helper_Core.TogglePerformance False
    
    ' Build a union of rows to delete them all at once (much faster)
    For i = 2 To lastRow
        If IsEmpty(ws.Cells(i, "B")) And IsEmpty(ws.Cells(i, "C")) And IsEmpty(ws.Cells(i, "D")) Then
            If rowsToDelete Is Nothing Then
                Set rowsToDelete = ws.Rows(i)
            Else
                Set rowsToDelete = Union(rowsToDelete, ws.Rows(i))
            End If
        End If
    Next i
    
    If Not rowsToDelete Is Nothing Then rowsToDelete.Delete
    
    Helper_Core.TogglePerformance True
    MsgBox "Conditionally empty rows deleted.", vbInformation
End Sub