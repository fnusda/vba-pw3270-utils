Attribute VB_Name = "API_Integration"
Option Explicit

' -------------------------------------------------------------------------
' Subroutine: FetchAndProcessPendingData
' Refactored from: requisicaoDados / jsonConverter2
' Purpose: Orchestrates fetching JSON data from an external API and 
'          injecting the parsed data into the terminal.
' Note: Requires the standard VBA-JSON library to be present in your project.
' -------------------------------------------------------------------------
Public Sub FetchAndProcessPendingData()
    Dim term As New clsTerminalMacros
    Dim parsedJson As Object
    Dim jsonItem As Variant
    Dim apiUrl As String
    
    ' 1. Configuration
    apiUrl = "https://api.yourdomain.com.br/v1/dados-pendentes"
    
    Helper_Core.TogglePerformance False ' Optimize Excel performance
    On Error GoTo ErrorHandler
    
    ' 2. Fetch Data via API
    Set parsedJson = FetchJsonFromAPI(apiUrl)
    
    ' Ensure data was retrieved successfully
    If parsedJson Is Nothing Then
        MsgBox "Failed to retrieve or parse data from the API.", vbCritical, "API Error"
        GoTo CleanUp
    End If
    
    ' 3. Connect to the Terminal
    term.Connect
    term.SendPFKey 3 ' Home screen
    
    ' 4. Process each JSON item in the terminal
    For Each jsonItem In parsedJson
        ProcessItemInTerminal term, jsonItem
    Next jsonItem
    
    MsgBox "API Integration and Terminal processing finished successfully!", vbInformation, "Completed"

CleanUp:
    On Error Resume Next
    term.Disconnect
    Set term = Nothing
    Helper_Core.TogglePerformance True
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred during API Integration: " & Err.Description, vbCritical, "Error " & Err.Number
    Resume CleanUp
End Sub

' -------------------------------------------------------------------------
' Function: FetchJsonFromAPI (Helper)
' Purpose: Makes an HTTP GET request to the specified URL and parses the response.
' -------------------------------------------------------------------------
Private Function FetchJsonFromAPI(ByVal targetUrl As String) As Object
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    With httpRequest
        .Open "GET", targetUrl, False
        .send
        
        ' Check for successful HTTP status
        If .Status = 200 Then
            ' Assuming JsonConverter is the standard VBA-JSON module in your project
            Set FetchJsonFromAPI = JsonConverter.ParseJson(.responseText)
        Else
            Set FetchJsonFromAPI = Nothing
        End If
    End With
End Function

' -------------------------------------------------------------------------
' Subroutine: ProcessItemInTerminal (Helper)
' Purpose: Maps the parsed JSON dictionary keys to terminal fields.
' -------------------------------------------------------------------------
Private Sub ProcessItemInTerminal(ByRef term As clsTerminalMacros, ByVal item As Object)
    ' Navigate to the required screen (e.g., Screen 03)
    term.PutString 22, 11, "03"
    term.Enter
    
    ' Inject data using standard dictionary keys from your JSON response
    ' Example: Assuming the API returns a "cpf_cnpj" and "customer_id" key
    term.PutString 16, 38, item("customer_id")
    term.Enter
    
    ' ... Add specific screen navigation and validation logic here ...
    
    term.SendPFKey 3 ' Return to home screen for next item
End Sub