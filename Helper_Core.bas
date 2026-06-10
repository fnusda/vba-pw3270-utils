Attribute VB_Name = "Helper_Core"
Option Explicit

' -------------------------------------------------------------------------
' Function: TogglePerformance
' Purpose: Speeds up macro execution by disabling screen updates and calculations
' -------------------------------------------------------------------------
Public Sub TogglePerformance(ByVal turnOn As Boolean)
    Application.ScreenUpdating = turnOn
    Application.EnableEvents = turnOn
    If turnOn Then
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
    Else
        Application.Calculation = xlCalculationManual
    End If
End Sub

' -------------------------------------------------------------------------
' Function: GetLastRow
' Purpose: Finds the last used row in a specific column dynamically
' -------------------------------------------------------------------------
Public Function GetLastRow(ByVal targetSheet As Worksheet, ByVal colLetter As String) As Long
    GetLastRow = targetSheet.Cells(targetSheet.Rows.Count, colLetter).End(xlUp).Row
End Function

' -------------------------------------------------------------------------
' Function: UpdateProgress
' Purpose: Displays the current execution progress in the Excel Status Bar
' -------------------------------------------------------------------------
Public Sub UpdateProgress(ByVal currentRow As Long, ByVal totalRows As Long)
    Dim percentComplete As Double
    If totalRows > 0 Then
        percentComplete = (currentRow / totalRows) * 100
        Application.StatusBar = "Progress: " & Format(percentComplete, "0.00") & "%"
    End If
End Sub