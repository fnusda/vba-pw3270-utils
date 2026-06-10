Attribute VB_Name = "Web_Automation"
Option Explicit

' --- Configuration Constants ---
Private Const SGC_URL As String = "https://sgc.sanepar.com.br/consultas/sgcConsultaMatriculas.aspx"
Private Const DOWNLOAD_PATH As String = "D:\Automacao_Sanepar\Faturas\"
Private Const HTML_ID_MATRICULA As String = "ctl00_ContentPlaceHolder1_fltMatricula"
Private Const HTML_ID_FILTER_BTN As String = "ctl00_ContentPlaceHolder1_btnFiltrar"

' -------------------------------------------------------------------------
' Subroutine: DownloadSGCInvoices
' Refactored from: RunChrome_SGCWeb / ExecutarEsteiraCompleta
' Purpose: Automates Chrome to fetch and download digital invoices (PDFs) 
'          for a list of customer IDs. Renames them upon successful download.
' -------------------------------------------------------------------------
Public Sub DownloadSGCInvoices()
    Dim chromeDriver As New clsChrome ' Ensure your custom Chrome class is imported
    Dim targetSheet As Worksheet
    Dim lastRow As Long, currentRow As Long
    Dim customerId As String
    Dim targetFolder As String
    Dim isDownloadSuccessful As Boolean

    Set targetSheet = ThisWorkbook.Worksheets("Dados")
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    
    If lastRow < 2 Then Exit Sub

    On Error GoTo ErrorHandler
    Helper_Core.TogglePerformance False

    ' Initialize Browser
    chromeDriver.Start
    chromeDriver.Attach ""

    For currentRow = 2 To lastRow
        Helper_Core.UpdateProgress currentRow - 1, targetSheet.Rows.Count
        customerId = Format(targetSheet.Cells(currentRow, 1).Value, "00000000")
        
        ' 1. Create a dedicated folder for the customer
        targetFolder = DOWNLOAD_PATH & customerId & "\"
        CreateDirectory targetFolder
        
        ' 2. Navigate and inject search parameters via JavaScript (faster than UI clicks)
        chromeDriver.Navigate SGC_URL
        chromeDriver.WaitCompletion
        
        chromeDriver.JSEval "document.getElementById('" & HTML_ID_MATRICULA & "').value = '" & customerId & "';"
        chromeDriver.JSEval "document.getElementById('" & HTML_ID_FILTER_BTN & "').click();"
        chromeDriver.WaitCompletion
        
        ' 3. Navigate to the Debts view
        chromeDriver.JSEval "const img = document.querySelector('img[title=""Visualizar Débitos""]'); if(img) img.click();"
        chromeDriver.WaitCompletion
        
        ' 4. Trigger the hidden PDF download
        isDownloadSuccessful = TriggerPDFDownload(chromeDriver, customerId)
        
        If isDownloadSuccessful Then
            targetSheet.Cells(currentRow, 3).Value = "Download triggered"
            ' Rename the file from the default system name to our standard
            RenameDownloadedPDF targetSheet, currentRow, targetFolder
        Else
            targetSheet.Cells(currentRow, 3).Value = "Error: Invoice link not found"
        End If
    Next currentRow

CleanUp:
    chromeDriver.Quit
    Helper_Core.TogglePerformance True
    MsgBox "Web Automation Finished!", vbInformation, "Completed"
    Exit Sub

ErrorHandler:
    targetSheet.Cells(currentRow, 3).Value = "Critical Error: " & Err.Description
    Resume CleanUp
End Sub

' -------------------------------------------------------------------------
' Function: TriggerPDFDownload (Helper)
' Purpose: Extracts the secure hash from the ASP.NET backend and forces a 
'          file download via JavaScript to avoid popup blockers.
' -------------------------------------------------------------------------
Private Function TriggerPDFDownload(ByRef browser As clsChrome, ByVal customerId As String) As Boolean
    Dim onClickValue As String
    Dim secureHash As String
    Dim pdfUrl As String
    Dim jsDownloadScript As String
    
    ' Extract the onClick attribute which contains the secure session hash
    onClickValue = browser.JSEval("const a = document.querySelector('#dtgListagemConta td:last-of-type a:last-of-type'); a ? a.getAttribute('onclick') : '';")
    
    If onClickValue = "" Or InStr(onClickValue, """") = 0 Then
        TriggerPDFDownload = False
        Exit Function
    End If
    
    ' Parse the hash and build the direct download URL
    secureHash = Split(onClickValue, """")(3)
    pdfUrl = "https://sgc.sanepar.com.br/documento/segundavia.aspx?matMF=" & customerId & "&hash=" & secureHash
    
    ' Inject an invisible anchor tag to force the download silently
    jsDownloadScript = "const a = document.createElement('a'); a.href = '" & pdfUrl & "'; " & _
                       "a.download = 'Fatura_" & customerId & ".pdf'; document.body.appendChild(a); " & _
                       "a.click(); document.body.removeChild(a);"
                       
    browser.JSEval jsDownloadScript
    TriggerPDFDownload = True
End Function

' -------------------------------------------------------------------------
' Subroutine: CreateDirectory (Helper)
' Purpose: Uses the FileSystemObject to ensure a folder exists.
' -------------------------------------------------------------------------
Private Sub CreateDirectory(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

' -------------------------------------------------------------------------
' Subroutine: RenameDownloadedPDF (Helper)
' Purpose: Finds the generic downloaded file and renames it.
' -------------------------------------------------------------------------
Private Sub RenameDownloadedPDF(ByRef targetSheet As Worksheet, ByVal rowIdx As Long, ByVal targetFolder As String)
    Dim oldFilePath As String
    Dim newFilePath As String
    
    ' Assuming Column 1 is Customer ID and Column 4 is the desired new filename
    oldFilePath = DOWNLOAD_PATH & targetSheet.Cells(rowIdx, 1).Value & "-FD.pdf"
    newFilePath = targetFolder & targetSheet.Cells(rowIdx, 4).Value & ".pdf"
    
    ' Wait briefly to ensure Chrome finished saving the file to disk
    Application.Wait Now + TimeValue("00:00:02")
    
    If Dir(oldFilePath) <> "" Then
        On Error Resume Next
        Name oldFilePath As newFilePath
        If Err.Number = 0 Then
            targetSheet.Cells(rowIdx, 5).Value = "Saved & Renamed"
        Else
            targetSheet.Cells(rowIdx, 5).Value = "Failed to rename"
        End If
        On Error GoTo 0
    Else
        targetSheet.Cells(rowIdx, 5).Value = "File not found in directory"
    End If
End Sub