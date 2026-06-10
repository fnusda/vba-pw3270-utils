Attribute VB_Name = "Data_Preparation"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ImportAndFormatData
' Refactored from: aa_importarDadosArrecadacao
' Purpose: Prompts the user to select an Excel file, copies the data from 
'          its first sheet, and pastes it as text into "Planilha2".
' -------------------------------------------------------------------------
Sub ImportAndFormatData()
    Dim sourceBook As Workbook
    Dim sourceSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim filePicker As FileDialog
    Dim filePath As String
    
    ' Set up the file picker dialog
    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    filePicker.Title = "Select the Source File"
    filePicker.Filters.Clear
    filePicker.Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
    
    ' Show dialog and get file path
    If filePicker.Show = -1 Then
        filePath = filePicker.SelectedItems(1)
    Else
        MsgBox "No file was selected.", vbExclamation, "Warning"
        Exit Sub
    End If
    
    Helper_Core.TogglePerformance False ' Optimize speed
    
    ' Open source and define target
    Set sourceBook = Workbooks.Open(filePath)
    Set sourceSheet = sourceBook.Sheets(1)
    Set targetSheet = ThisWorkbook.Sheets("Planilha2")
    
    ' Copy data
    sourceSheet.UsedRange.Copy
    
    ' Paste as raw text and apply formatting
    With targetSheet.Cells(1, 1)
        .PasteSpecial Paste:=xlPasteValuesAndNumberFormats
        .NumberFormat = "@" ' Enforce text format
    End With
    
    With targetSheet.Cells
        .Font.Name = "Arial"
        .Font.Size = 10
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    ' Cleanup
    Application.CutCopyMode = False
    sourceBook.Close False
    
    Helper_Core.TogglePerformance True ' Restore Excel state
    MsgBox "Data imported and formatted successfully!", vbInformation, "Done"
End Sub

' -------------------------------------------------------------------------
' Subroutine: RemoveAllMergedCells
' Refactored from: ab_desfazerCelMesc
' Purpose: Loops through the active sheet and unmerges any merged cells.
' -------------------------------------------------------------------------
Sub RemoveAllMergedCells()
    Dim targetSheet As Worksheet
    Dim cellRange As Range
    
    Set targetSheet = ActiveSheet
    Helper_Core.TogglePerformance False
    
    ' Unmerge all cells in the used range
    For Each cellRange In targetSheet.UsedRange
        If cellRange.MergeCells Then
            cellRange.UnMerge
        End If
    Next cellRange
    
    Helper_Core.TogglePerformance True
End Sub

' -------------------------------------------------------------------------
' Subroutine: DeleteEmptyColumns
' Refactored from: ae_removeColVazias
' Purpose: Scans columns from right to left and deletes completely empty ones.
' -------------------------------------------------------------------------
Sub DeleteEmptyColumns()
    Dim targetSheet As Worksheet
    Dim lastCol As Long
    Dim i As Long
    
    Set targetSheet = ActiveSheet
    Helper_Core.TogglePerformance False
    
    ' Find the last used column
    lastCol = targetSheet.Cells(1, targetSheet.Columns.Count).End(xlToLeft).Column
    
    ' Loop backwards to safely delete columns
    For i = lastCol To 1 Step -1
        If Application.WorksheetFunction.CountA(targetSheet.Columns(i)) = 0 Then
            targetSheet.Columns(i).Delete
        End If
    Next i
    
    Helper_Core.TogglePerformance True
End Sub