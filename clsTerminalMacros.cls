VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTerminalMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -------------------------------------------------------------------------
' Class: clsTerminalMacros
' Purpose: PW3270 Terminal Automation Wrapper using libhllapi.dll
' Refactored for Agile RPA, 64-bit compatibility, and robust error handling.
' -------------------------------------------------------------------------

' --- Configuration Constants ---
Private Const SCREEN_ROWS As Integer = 24
Private Const SCREEN_COLS As Integer = 80
Private Const DEFAULT_TIMEOUT_SEC As Integer = 30
Private Const DLL_PATH_A As String = "C:\Program Files\PW3270\hllapi.dll"
Private Const DLL_PATH_B As String = "C:\Windows\SysWOW64\libhllapi.dll"

' --- HLLAPI DLL Declarations ---
Private Declare PtrSafe Function hllapi_init Lib "libhllapi.dll" (ByVal tp As String) As Long
Private Declare PtrSafe Function hllapi_deinit Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_get_revision Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_connect Lib "libhllapi.dll" (ByVal uri As String, ByVal wait As Integer) As Long
Private Declare PtrSafe Function hllapi_disconnect Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_wait_for_ready Lib "libhllapi.dll" (ByVal timeOut As Integer) As Long
Private Declare PtrSafe Function hllapi_get_screen_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_enter Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_get_message_id Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_set_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_wait Lib "libhllapi.dll" (ByVal timeOut As Integer) As Long
Private Declare PtrSafe Function hllapi_pfkey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare PtrSafe Function hllapi_pakey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare PtrSafe Function hllapi_cmp_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_is_connected Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_set_unlock_delay Lib "libhllapi.dll" (ByVal delay As Integer) As Long
Private Declare PtrSafe Function hllapi_get_cursor_address Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_set_cursor_address Lib "libhllapi.dll" (ByVal addr As Integer) As Long
Private Declare PtrSafe Function hllapi_action Lib "libhllapi.dll" (ByVal keyname As String) As Long
Private Declare PtrSafe Function hllapi_find_text Lib "libhllapi.dll" (ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_set_session_parameter Lib "libhllapi.dll" (ByVal text As String, ByVal size As Integer, ByVal value As Integer) As Long

' --- Windows API Declarations (32-bit & 64-bit Compatible) ---
Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal wCmd As Long) As LongPtr
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal nCmdShow As Long) As Long
#Else
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
    Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
#End If

Private Const GW_HWNDNEXT = 2
Private pWinTitle As String
Private pFullTitle As String

' ==============================================================================
' Core Terminal Functions
' ==============================================================================

Public Function Connect(Optional ByVal host As String = "") As Boolean
    Dim hostArray As Variant
    Dim i As Integer
    
    ' Pre-load DLLs safely (ignores error if path doesn't exist)
    On Error Resume Next
    LoadLibrary DLL_PATH_A
    LoadLibrary DLL_PATH_B
    On Error GoTo 0

    If host <> "" Then
        If hllapi_init(host) <> 0 Then
            Err.Raise vbObjectError + 1, "clsTerminalMacros.Connect", "Failed to connect! Ensure terminal session '" & host & "' is open."
        Else
            hllapi_set_unlock_delay 1
            Connect = True
        End If
    Else
        ' Try default sessions if none provided
        hostArray = Array("pw3270:A", "pw3270:B", "pw3270:C", "pw3270:D")
        For i = 0 To 3
            If hllapi_init(CStr(hostArray(i))) = 0 Then
                hllapi_set_unlock_delay 1
                Connect = True
                Exit Function
            End If
        Next i
        
        Err.Raise vbObjectError + 2, "clsTerminalMacros.Connect", "Connection attempts exhausted! Ensure a terminal session is active."
    End If
End Function

Public Sub Disconnect()
    hllapi_deinit
End Sub

Public Sub WaitHost(ByVal seconds As Integer)
    hllapi_wait seconds
End Sub

Public Function WaitHostOK(ByVal timeoutSec As Integer) As Boolean
    hllapi_wait_for_ready timeoutSec
    hllapi_wait 0.1
    WaitForReady
    WaitHostOK = True
End Function

Public Function GetString(ByVal row As Integer, ByVal col As Integer, ByVal length As Long) As String
    Dim txtBuffer As String
    txtBuffer = String$(length, " ") ' Faster string allocation than Space()
    
    hllapi_get_screen_at row, col, txtBuffer
    GetString = txtBuffer
End Function

Public Sub PutString(ByVal row As Integer, ByVal col As Integer, ByVal text As String)
    hllapi_set_text_at CInt(row), CInt(col), text
    WaitForReady
End Sub

Public Sub Enter()
    hllapi_enter
    WaitForReady
End Sub

Public Sub SendPFKey(ByVal keycode As Integer)
    hllapi_pfkey keycode
    WaitForReady
End Sub

Public Sub SendPAKey(ByVal keycode As Integer)
    hllapi_pakey keycode
    WaitForReady
End Sub

Public Sub SendEspKey(ByVal keyname As String)
    Dim mappedKey As String
    
    Select Case UCase(keyname)
        Case "HOME": mappedKey = "firstfield"
        Case "END":  mappedKey = "fieldend"
        Case "TAB":  mappedKey = "nextfield"
        Case Else:   mappedKey = keyname
    End Select

    hllapi_action mappedKey
    WaitForReady
End Sub

Public Function FindText(ByVal text As String) As Long
    FindText = hllapi_find_text(text)
End Function

' ==============================================================================
' Cursor & Screen Management
' ==============================================================================

Public Function SetCursor(ByVal row As Integer, ByVal col As Integer)
    Dim addr As Integer
    addr = ((row - 1) * SCREEN_COLS) + col
    hllapi_set_cursor_address addr
    WaitForReady
End Function

Public Function SetCursorADDR(ByVal addr As Integer)
    hllapi_set_cursor_address addr
    WaitForReady
End Function

Public Function GetCursor() As Integer
    GetCursor = hllapi_get_cursor_address()
End Function

Public Function WaitForCursor(ByVal row As Integer, ByVal col As Integer) As Boolean
    Dim targetAddr As Integer
    targetAddr = ((row - 1) * SCREEN_COLS) + col
    WaitForCursor = (hllapi_get_cursor_address() = targetAddr)
End Function

Public Function GetScreen(Optional ByVal start_row As Integer = 1, Optional ByVal end_row As Integer = SCREEN_ROWS) As String
    Dim txtBuffer As String
    Dim resultTxt As String
    Dim i As Integer

    txtBuffer = String$(SCREEN_COLS, " ")
    
    For i = start_row To end_row
        hllapi_get_screen_at i, 1, txtBuffer
        resultTxt = resultTxt & txtBuffer & vbCrLf
    Next i

    GetScreen = resultTxt
    WaitForReady
End Function

Public Sub ChangeParam(ByVal name As String, ByVal size As Integer, ByVal value As Integer)
    hllapi_set_session_parameter name, size, value
End Sub

Public Function Status() As Long
    Status = hllapi_get_message_id()
End Function

Public Function IsConnected() As Boolean
    IsConnected = (hllapi_is_connected() <> 0)
End Function

Public Sub MoveTo(ByVal row As Variant, ByVal col As Variant, Optional ByVal page As Integer = 0)
    SetCursor CInt(row), CInt(col)
End Sub

' ==============================================================================
' Advanced Synchronization & Parsing
' ==============================================================================

' Safely waits for the terminal to process the last action.
' Replaces the dangerous 'While Status <> 0: DoEvents: Wend' infinite loops.
Public Sub WaitExec()
    WaitForReady
End Sub

Private Sub WaitForReady(Optional ByVal timeoutSec As Integer = DEFAULT_TIMEOUT_SEC)
    Dim endTime As Double
    endTime = Timer + timeoutSec
    
    Do While hllapi_get_message_id() <> 0
        If Timer > endTime Then
            Err.Raise vbObjectError + 2, "clsTerminalMacros", "Terminal operation timed out."
        End If
        DoEvents
        Sleep 50 ' Prevent 100% CPU lock while polling
    Loop
End Sub

Public Sub WaitForString(ByVal text As String, ByVal row As Integer, ByVal col As Integer, Optional ByVal timeoutSec As Integer = DEFAULT_TIMEOUT_SEC)
    Dim endTime As Double
    Dim currentText As String
    
    endTime = Timer + timeoutSec
    
    Do While Timer < endTime
        currentText = Me.GetString(row, col, Len(text))
        If currentText = text Then Exit Sub
        
        Sleep 500
        DoEvents
    Loop
    
    Err.Raise vbObjectError + 3, "clsTerminalMacros.WaitForString", "Wait timeout exceeded looking for: '" & text & "'"
End Sub

' Highly optimized SendKeys method using string tokenization
' Replaces the fragile GoTo jumps and array shuffling.
Public Sub SendKeys(ByVal commandString As String)
    Dim i As Integer, endIdx As Integer
    Dim char As String, tag As String
    Dim currRow As Integer, currCol As Integer
    Dim cursorAddr As Integer
    
    ' Fetch current physical position
    cursorAddr = Me.GetCursor()
    currRow = 1 + Int(cursorAddr / SCREEN_COLS)
    currCol = cursorAddr Mod SCREEN_COLS
    If currCol = 0 Then
        currRow = currRow - 1
        currCol = SCREEN_COLS
    End If
    
    i = 1
    Do While i <= Len(commandString)
        char = Mid(commandString, i, 1)
        
        If char = "<" Then
            endIdx = InStr(i + 1, commandString, ">")
            If endIdx > 0 Then
                tag = UCase(Mid(commandString, i + 1, endIdx - i - 1))
                i = endIdx ' Fast-forward loop pointer
                ProcessSendKeyTag tag, currRow, currCol
            Else
                Err.Raise vbObjectError + 4, "clsTerminalMacros.SendKeys", "Unclosed tag '<' in SendKeys string."
            End If
        Else
            Me.PutString currRow, currCol, char
            currCol = currCol + 1
            If currCol > SCREEN_COLS Then
                currCol = 1
                currRow = currRow + 1
            End If
        End If
        i = i + 1
    Loop
End Sub

Private Sub ProcessSendKeyTag(ByVal tag As String, ByRef currRow As Integer, ByRef currCol As Integer)
    Select Case tag
        Case "ENTER": Me.Enter
        Case "TAB":   Me.SendEspKey "TAB"
        Case "HOME":  Me.SendEspKey "HOME"
        Case "END":   Me.SendEspKey "END"
        Case Else
            If Left(tag, 2) = "PF" Then
                Me.SendPFKey CInt(Mid(tag, 3))
            End If
    End Select
    
    ' Resync local coordinate trackers after special key execution
    Dim cursorAddr As Integer
    cursorAddr = Me.GetCursor()
    currRow = 1 + Int(cursorAddr / SCREEN_COLS)
    currCol = cursorAddr Mod SCREEN_COLS
    If currCol = 0 Then
        currRow = currRow - 1
        currCol = SCREEN_COLS
    End If
End Sub

' ==============================================================================
' Process Management (Window Hooks)
' ==============================================================================

Public Function NewNamedSess(ByVal sessName As String, Optional ByVal hidden As Boolean = True) As Boolean
    Dim wh As LongPtr
    Dim oShell As Object
    Dim waitTicks As Integer
    
    Set oShell = CreateObject("WScript.Shell")
    oShell.Run "C:\Program Files\pw3270\pw3270.EXE --session " & sessName
    
    ' Poll for window authorization
    For waitTicks = 1 To DEFAULT_TIMEOUT_SEC
        If GetHndPartCaption(wh, sessName) Then
            If hidden Then ShowWindow wh, 0
            hllapi_connect pWinTitle, 0
            hllapi_wait 1
            NewNamedSess = True
            Exit Function
        Else
            Sleep 1000
            DoEvents
        End If
    Next waitTicks
    
    Err.Raise vbObjectError + 5, "clsTerminalMacros.NewNamedSess", "Timeout waiting for terminal authorization."
End Function

Public Sub DisconnectNamedSess(ByVal sessName As String)
    Dim wh As LongPtr
    Dim oShell As Object
    Dim waitTicks As Integer
    
    Set oShell = CreateObject("WScript.Shell")
    hllapi_deinit
    
    For waitTicks = 1 To DEFAULT_TIMEOUT_SEC
        If GetHndPartCaption(wh, sessName) Then
            ShowWindow wh, 1
            hllapi_wait 1
            Exit For
        Else
            Sleep 1000
            DoEvents
        End If
    Next waitTicks
    
    If waitTicks > DEFAULT_TIMEOUT_SEC Then
        Err.Raise vbObjectError + 6, "clsTerminalMacros.DisconnectNamedSess", "Timeout trying to close terminal."
    End If
    
    ' Forcefully kill window
    oShell.Run "taskkill /F /FI ""WINDOWTITLE eq " & pFullTitle & """ /T"
End Sub

Private Function GetHndPartCaption(ByRef lWnd As LongPtr, ByVal sCaption As String) As Boolean
    Dim lhWndP As LongPtr
    Dim sStr As String
    Dim textLength As Long

    GetHndPartCaption = False
    lhWndP = FindWindow(vbNullString, vbNullString)
    
    Do While lhWndP <> 0
        textLength = GetWindowTextLength(lhWndP)
        If textLength > 0 Then
            sStr = String$(textLength, Chr$(0))
            GetWindowText lhWndP, sStr, textLength + 1
            
            If InStr(1, sStr, sCaption, vbTextCompare) > 0 Then
                GetHndPartCaption = True
                lWnd = lhWndP
                pWinTitle = Mid(sStr, InStr(1, sStr, sCaption), Len(sCaption) + 2)
                pFullTitle = sStr
                Exit Do
            End If
        End If
        lhWndP = GetWindow(lhWndP, GW_HWNDNEXT)
    Loop
End Function