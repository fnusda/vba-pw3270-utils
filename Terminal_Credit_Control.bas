Attribute VB_Name = "Terminal_Credit_Control"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: ExtractCustomerEvaluations
' Refactored from: tela47 / importaDados
' Purpose: Navigates to Screen 47, Option 15, prompts for a unit code, 
'          and paginates to extract credit evaluation data.
' -------------------------------------------------------------------------
Public Sub ExtractCustomerEvaluations()
    Dim term As New clsMacrosSanepar
    Dim targetSheet As Worksheet
    Dim unitCode As String
    Dim currentPage As Integer, lastPage As Integer
    Dim termRow As Integer
    Dim excelRow As Long
    
    unitCode = InputBox("Enter the Unit Code:", "Unit Input")
    If unitCode = "" Then Exit Sub
    
    Set targetSheet = ActiveSheet
    excelRow = 2
    
    Helper_Core.TogglePerformance False
    term.Connect
    term.SendPFKey 3
    
    ' Navigate to 47 > 15
    term.PutString 22, 11, "47"
    term.Enter
    term.PutString 22, 57, "15"
    term.Enter
    term.SetCursor 9, 28
    term.Enter
    
    ' Input Unit Code
    term.PutString 15, 60, unitCode
    term.Enter
    
    Do
        currentPage = CInt(term.GetString(4, 72, 3))
        lastPage = CInt(term.GetString(4, 78, 3))
        
        ' Read grid (Rows 8 to 22)
        For termRow = 8 To 22
            targetSheet.Cells(excelRow, "A").Value = Trim(term.GetString(termRow, 10, 3))
            targetSheet.Cells(excelRow, "B").Value = Trim(term.GetString(termRow, 16, 3))
            
            ' Format Col C as text to preserve leading zeros
            targetSheet.Cells(excelRow, "C").NumberFormat = "@"
            targetSheet.Cells(excelRow, "C").Value = Format(Trim(term.GetString(termRow, 22, 9)), "00000000")
            
            targetSheet.Cells(excelRow, "D").Value = Trim(term.GetString(termRow, 34, 29))
            targetSheet.Cells(excelRow, "E").Value = Trim(term.GetString(termRow, 71, 10))
            excelRow = excelRow + 1
        Next termRow
        
        ' End loop if we reached the final page
        If Trim(term.GetString(2, 4, 11)) = "ULTIMA TELA" Or currentPage >= lastPage Then Exit Do
        
        ' Next Page
        term.SendPFKey 8
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    
    term.SendPFKey 3
    Helper_Core.TogglePerformance True
    MsgBox "Evaluation Data Extraction Complete.", vbInformation
End Sub