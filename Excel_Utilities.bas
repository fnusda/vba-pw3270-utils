Attribute VB_Name = "Excel_Utilities"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: GenerateValidCPFs
' Refactored from: GerarCPFsValidos
' Purpose: Generates a list of valid Brazilian CPF combinations based on a 
'          base string and calculates the two verification digits.
' -------------------------------------------------------------------------
Public Sub GenerateValidCPFs()
    Dim targetSheet As Worksheet
    Dim baseCPF As String, fullCPF As String
    Dim i As Integer, j As Integer, k As Integer
    Dim digit1 As Integer, digit2 As Integer
    Dim currentRow As Long
    
    Set targetSheet = ThisWorkbook.Sheets(1)
    baseCPF = "727409" ' Can be parameterized via InputBox if needed
    currentRow = 2
    
    ' Set up headers
    targetSheet.Cells(1, 2).Resize(1, 11).Value = Array("D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "DV1", "DV2")
    
    Helper_Core.TogglePerformance False

    For i = 0 To 9
        For j = 0 To 9
            For k = 0 To 9
                Dim cpfPrefix As String
                cpfPrefix = CStr(i) & CStr(j) & CStr(k) & baseCPF
                
                digit1 = CalculateCPFDigit(cpfPrefix, 10)
                digit2 = CalculateCPFDigit(cpfPrefix & digit1, 11)
                fullCPF = cpfPrefix & digit1 & digit2
                
                If IsValidCPF(fullCPF) Then
                    Dim col As Integer
                    ' Break down the CPF into individual columns
                    For col = 1 To 11
                        targetSheet.Cells(currentRow, col + 1).Value = Mid(fullCPF, col, 1)
                    Next col
                    currentRow = currentRow + 1
                End If
            Next k
        Next j
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Valid CPFs generated and organized by column!", vbInformation
End Sub

' -------------------------------------------------------------------------
' Function: CalculateCPFDigit (Helper)
' Purpose: Performs the modulo 11 math required to find CPF check digits.
' -------------------------------------------------------------------------
Private Function CalculateCPFDigit(ByVal cpfStr As String, ByVal initialWeight As Integer) As Integer
    Dim sum As Integer, i As Integer
    Dim remainder As Integer
    
    sum = 0
    For i = 1 To Len(cpfStr)
        sum = sum + Val(Mid(cpfStr, i, 1)) * (initialWeight - i)
    Next i
    
    remainder = sum Mod 11
    
    If remainder < 2 Then
        CalculateCPFDigit = 0
    Else
        CalculateCPFDigit = 11 - remainder
    End If
End Function

' -------------------------------------------------------------------------
' Function: IsValidCPF (Helper)
' Purpose: Validates a full CPF string according to algorithmic rules.
' -------------------------------------------------------------------------
Private Function IsValidCPF(ByVal cpfStr As String) As Boolean
    Dim sum As Integer, i As Integer
    Dim formattedSum As String
    
    sum = 0
    For i = 1 To Len(cpfStr)
        sum = sum + Val(Mid(cpfStr, i, 1))
    Next i
    
    formattedSum = Format(sum, "00")
    
    ' Check if the first and last digits of the sum are identical
    IsValidCPF = (Left(formattedSum, 1) = Right(formattedSum, 1))
End Function

' -------------------------------------------------------------------------
' Subroutine: CreateSheetsFromList
' Refactored from: criaPlanilhas
' Purpose: Creates new worksheet tabs based on a list of names in "colaboradoresPG".
' -------------------------------------------------------------------------
Public Sub CreateSheetsFromList()
    Dim sourceSheet As Worksheet, newSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim sheetName As String
    
    Set sourceSheet = ThisWorkbook.Sheets("colaboradoresPG")
    lastRow = Helper_Core.GetLastRow(sourceSheet, "A")
    
    Helper_Core.TogglePerformance False
    
    For i = 2 To lastRow
        ' Extract the first name only to use as the sheet tab name
        sheetName = Split(sourceSheet.Cells(i, 2).Value, " ")(0)
        
        On Error Resume Next
        Set newSheet = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        
        ' If the sheet does not exist, create it
        If newSheet Is Nothing Then
            Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            newSheet.Name = sheetName
        End If
        Set newSheet = Nothing
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Worksheets created successfully!", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: SortSheetsAlphabetically
' Refactored from: organizarPlanilhas
' Purpose: Sorts all worksheet tabs in the workbook in alphabetical order.
' -------------------------------------------------------------------------
Public Sub SortSheetsAlphabetically()
    Dim i As Integer, j As Integer
    Dim tempName As String
    Dim sheetNames() As String
    
    ReDim sheetNames(1 To ThisWorkbook.Sheets.Count)
    
    ' Gather all sheet names
    For i = 1 To ThisWorkbook.Sheets.Count
        sheetNames(i) = ThisWorkbook.Sheets(i).Name
    Next i
    
    ' Standard Bubble Sort algorithm
    For i = LBound(sheetNames) To UBound(sheetNames) - 1
        For j = i + 1 To UBound(sheetNames)
            If UCase(sheetNames(i)) > UCase(sheetNames(j)) Then
                tempName = sheetNames(i)
                sheetNames(i) = sheetNames(j)
                sheetNames(j) = tempName
            End If
        Next j
    Next i
    
    Helper_Core.TogglePerformance False
    
    ' Reorder the actual sheets to match the sorted array
    For i = 1 To UBound(sheetNames)
        ThisWorkbook.Sheets(sheetNames(i)).Move After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Next i
    
    Helper_Core.TogglePerformance True
    MsgBox "Worksheets organized alphabetically!", vbInformation
End Sub