Attribute VB_Name = "Integration_Office"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractDataFromWord
' Refactored from: Receita
' Purpose: Loops through Word documents in a specified folder, opens them 
'          in the background, extracts specific sentences, and logs them.
' Note: Requires "Microsoft Word Object Library" reference to be enabled.
' -------------------------------------------------------------------------
Public Sub ExtractDataFromWord()
    Dim targetSheet As Worksheet
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim folderPath As String
    Dim currentFileName As String
    Dim fullFilePath As String
    Dim nextRow As Long
    
    On Error GoTo ErrorHandler
    
    Set targetSheet = ThisWorkbook.Worksheets("Receita")
    
    ' Assumes the folder path is stored in Row 1, Column 7 (G1)
    folderPath = targetSheet.Cells(1, 7).Value
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    currentFileName = Dir(folderPath & "*.doc*")
    If currentFileName = "" Then
        MsgBox "No Word documents found in the specified folder.", vbExclamation
        Exit Sub
    End If
    
    Helper_Core.TogglePerformance False
    
    ' Initialize Word Application invisibly for speed
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Do While currentFileName <> ""
        fullFilePath = folderPath & currentFileName
        Set wordDoc = wordApp.Documents.Open(FileName:=fullFilePath, ReadOnly:=True, Visible:=False)
        
        ' Find the next available row in Excel
        nextRow = Helper_Core.GetLastRow(targetSheet, "A") + 1
        
        ' Extract sentences directly to Excel (Assumes specific document structure)
        If wordDoc.Sentences.Count >= 8 Then
            targetSheet.Cells(nextRow, 1).Value = currentFileName
            targetSheet.Cells(nextRow, 2).Value = wordDoc.Sentences(7).Text
            targetSheet.Cells(nextRow, 3).Value = wordDoc.Sentences(8).Text
            targetSheet.Cells(nextRow, 4).Value = wordDoc.Sentences(4).Text & wordDoc.Sentences(5).Text
        Else
            targetSheet.Cells(nextRow, 1).Value = currentFileName
            targetSheet.Cells(nextRow, 2).Value = "Document format invalid or too short."
        End If
        
        wordDoc.Close False
        currentFileName = Dir ' Get the next file
    Loop
    
CleanUp:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Helper_Core.TogglePerformance True
    MsgBox "Word Extraction Completed!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during Word extraction: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' -------------------------------------------------------------------------
' Subroutine: ExtractTextFromPDF
' Refactored from: ExtrairDadosDoPDF
' Purpose: Prompts the user to select a PDF, uses Adobe Acrobat to read 
'          all text from the file, and dumps it into a new worksheet.
' Note: Requires Adobe Acrobat Standard/Pro installed on the machine.
' -------------------------------------------------------------------------
Public Sub ExtractTextFromPDF()
    Dim filePicker As FileDialog
    Dim pdfFilePath As String
    Dim acrobatApp As Object
    Dim acrobatAVDoc As Object
    Dim acrobatPDDoc As Object
    Dim acrobatHiliteList As Object
    Dim pdfTextOutput As String
    Dim newSheet As Worksheet
    Dim i As Integer
    
    ' Setup File Dialog
    Set filePicker = Application.FileDialog(msoFileDialogFilePicker)
    filePicker.Title = "Select the target PDF file"
    filePicker.AllowMultiSelect = False
    filePicker.Filters.Clear
    filePicker.Filters.Add "PDF Files", "*.pdf"
    
    If filePicker.Show = -1 Then
        pdfFilePath = filePicker.SelectedItems(1)
        
        Helper_Core.TogglePerformance False
        
        ' Initialize Acrobat Objects via Late Binding
        On Error GoTo ErrorHandler
        Set acrobatApp = CreateObject("AcroExch.App")
        Set acrobatAVDoc = CreateObject("AcroExch.AVDoc")
        
        If acrobatAVDoc.Open(pdfFilePath, "") Then
            Set acrobatPDDoc = acrobatAVDoc.GetPDDoc
            Set acrobatHiliteList = CreateObject("AcroExch.HiliteList")
            
            ' Extract text across all pages
            For i = 0 To acrobatPDDoc.GetNumPages - 1
                acrobatHiliteList.Add 0, acrobatPDDoc.GetPageNumWords(i), i
            Next i
            
            ' Grab the consolidated text
            pdfTextOutput = acrobatHiliteList.GetText(0)
            
            ' Output to a new Excel Sheet
            Set newSheet = ThisWorkbook.Sheets.Add
            newSheet.Name = "PDF_Data_" & Format(Now, "hhmmss")
            newSheet.Cells(1, 1).Value = pdfTextOutput
            
            acrobatAVDoc.Close True
        Else
            MsgBox "Error opening the PDF file in Acrobat.", vbCritical
        End If
    Else
        MsgBox "No