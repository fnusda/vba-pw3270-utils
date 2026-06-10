Attribute VB_Name = "Terminal_Specialized_Updates"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: UpdatePropertyNumbers
' Refactored from: insereNumero (Screen 36)
' Purpose: Updates specific property identifiers (Imóvel, Lado, Frente).
' -------------------------------------------------------------------------
Public Sub UpdatePropertyNumbers()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim propertyNo As String, sideNo As String, frontNo As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        propertyNo = targetSheet.Cells(currentCell.Row, "B").Value
        sideNo = targetSheet.Cells(currentCell.Row, "C").Value
        frontNo = targetSheet.Cells(currentCell.Row, "D").Value
        
        term.PutString 22, 11, "36"
        term.Enter
        
        term.PutString 13, 40, Format(currentCell.Value, "00000000")
        term.SetCursor 16, 37
        term.Enter
        
        If propertyNo <> "" Then term.PutString 21, 30, Format(propertyNo, "00000")
        If sideNo <> "" Then term.PutString 22, 30, Format(sideNo, "00000")
        If frontNo <> "" Then term.PutString 22, 51, Format(frontNo, "00000")
        
        term.Enter
        term.PutString 23, 78, "s"
        term.Enter
        term.SendPFKey 3
        
        ' Bypass routing/origin alerts
        If term.GetString(1, 2, 3) = "MUD" Or term.GetString(1, 2, 3) = "ORI" Then
            term.SendPFKey 1
            term.SendPFKey 3
        End If
            
        targetSheet.Cells(currentCell.Row, "F").Value = "Updated"
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Property numbers updated.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: BillSealReplacement
' Refactored from: faturamentoLacres (Screen 26)
' Purpose: Registers service 8955 for seal replacements (lacres).
' -------------------------------------------------------------------------
Public Sub BillSealReplacement()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        term.PutString 22, 11, "26"
        term.Enter
        
        term.PutString 9, 28, Format(currentCell.Value, "00000000")
        term.PutString 10, 28, "8955"
        term.PutString 13, 25, "faturamento / leitura"
        term.PutString 15, 25, "lacrar ligacao"
        term.Enter
        
        term.SendPFKey 2
        term.PutString 24, 31, "n"
        term.Enter
        
        ' Check if billing applies (Col B)
        If LCase(targetSheet.Cells(currentCell.Row, "B").Value) = "sim" Then
            term.PutString 24, 44, "01"
            term.Enter
            term.PutString 24, 69, "s"
            term.Enter
            targetSheet.Cells(currentCell.Row, "C").Value = "Billed"
        Else
            term.PutString 24, 44, "00"
            term.Enter
            term.PutString 24, 69, "s"
            term.Enter
            targetSheet.Cells(currentCell.Row, "C").Value = "Not Billed"
        End If
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Seal replacements registered.", vbInformation
End Sub

' -------------------------------------------------------------------------
' Subroutine: ExtractPropertySpecifics
' Refactored from: consultaEconomias / confereLogELote (Screen 03)
' Purpose: Extracts economy categories (RCPIU) and Log/Lote data.
' -------------------------------------------------------------------------
Public Sub ExtractPropertySpecifics()
    Dim term As New clsTerminalMacros
    Dim targetSheet As Worksheet
    Dim dataRange As Range, currentCell As Range
    Dim lastRow As Long
    Dim economies As String
    
    Set targetSheet = ActiveSheet
    lastRow = Helper_Core.GetLastRow(targetSheet, "A")
    If lastRow < 2 Then Exit Sub

    Helper_Core.TogglePerformance False
    Set dataRange = targetSheet.Range("A2:A" & lastRow)
    
    term.Connect
    term.SendPFKey 3

    For Each currentCell In dataRange
        term.PutString 22, 11, "03" & Format(currentCell.Value, "00000000")
        term.Enter
            
        If term.GetString(15, 26, 4) = "ATEN" Then
            term.PutString 16, 72, "n"
            term.Enter
        End If
        
        ' Extract RCPIU Economies
        economies = Replace(term.GetString(5, 58, 3), " ", "R") & _
                    Replace(term.GetString(5, 63, 3), " ", "C") & _
                    Replace(term.GetString(5, 68, 3), " ", "I") & _
                    Replace(term.GetString(5, 73, 3), " ", "U") & _
                    Replace(term.GetString(5, 78, 3), " ", "P")
                    
        targetSheet.Cells(currentCell.Row, "B").Value = economies
        targetSheet.Cells(currentCell.Row, "D").Value = Trim(term.GetString(7, 10, 5)) ' Logradouro
        targetSheet.Cells(currentCell.Row, "I").Value = Trim(term.GetString(9, 74, 6)) ' Lote
        
        term.SendPFKey 3
    Next currentCell

    Helper_Core.TogglePerformance True
    MsgBox "Property specifics extracted.", vbInformation
End Sub