Attribute VB_Name = "System_Execution"
Option Explicit

' --- API Configuration Constants ---
Private Const STARTF_USESTDHANDLES As Long = &H100
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_HIDE As Long = 0
Private Const CREATE_NO_WINDOW As Long = &H8000000

' --- Encapsulated API Types ---
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type

' --- Kernel32 Declarations ---
Private Declare PtrSafe Function CreatePipe Lib "kernel32" (phReadPipe As LongPtr, phWritePipe As LongPtr, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

' -------------------------------------------------------------------------
' Function: ExecuteConsoleCommand
' Refactored from: ExecuteCommand / modExec2
' Purpose: Executes a command-line string silently in the background, pipes 
'          the output, and returns the console response as a String.
' -------------------------------------------------------------------------
Public Function ExecuteConsoleCommand(ByVal commandLine As String) As String
    Dim secAttributes As SECURITY_ATTRIBUTES
    Dim startInfo As STARTUPINFO
    Dim procInfo As PROCESS_INFORMATION
    Dim hReadPipe As LongPtr, hWritePipe As LongPtr
    Dim readBuffer As String * 1024
    Dim bytesRead As Long
    Dim consolidatedResult As String
    Dim apiSuccess As Long

    ' 1. Set Security Attributes to allow handle inheritance
    secAttributes.nLength = Len(secAttributes)
    secAttributes.bInheritHandle = 1
    secAttributes.lpSecurityDescriptor = 0

    ' 2. Create the Communication Pipe
    If CreatePipe(hReadPipe, hWritePipe, secAttributes, 0) = 0 Then
        ExecuteConsoleCommand = "Error: Failed to create Pipe."
        Exit Function
    End If

    ' 3. Configure Startup Info for a Hidden Process
    startInfo.cb = Len(startInfo)
    startInfo.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    startInfo.wShowWindow = SW_HIDE
    startInfo.hStdOutput = hWritePipe
    startInfo.hStdError = hWritePipe ' Route errors to the same output channel

    ' 4. Launch the Process
    apiSuccess = CreateProcessA(0, commandLine, ByVal 0&, ByVal 0&, 1, CREATE_NO_WINDOW, 0, 0, startInfo, procInfo)

    If apiSuccess <> 0 Then
        ' Close the write handle in the parent process to prevent ReadFile from hanging
        CloseHandle hWritePipe
        
        ' 5. Loop to read output from the Pipe until empty
        Do
            apiSuccess = ReadFile(hReadPipe, readBuffer, Len(readBuffer), bytesRead, 0)
            If bytesRead > 0 Then
                consolidatedResult = consolidatedResult & Left$(readBuffer, bytesRead)
            End If
        Loop While apiSuccess <> 0
        
        ' Clean up process handles
        CloseHandle procInfo.hProcess
        CloseHandle procInfo.hThread
    Else
        consolidatedResult = "Error: Could not launch process for command: " & commandLine
        CloseHandle hWritePipe
    End If

    CloseHandle hReadPipe
    ExecuteConsoleCommand = consolidatedResult
End Function