Attribute VB_Name = "Process"
Option Explicit

Private lngProcess As Long
Private lngThread As Long
Private lngProcessID As Long
Private lngThreadID As Long
Private lngReply As Long

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
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
        lpReserved2 As Byte
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcess Lib "kernel32" Alias _
    "CreateProcessA" (ByVal lpApplicationName As String, ByVal _
    lpCommandLine As String, lpProcessAttributes As Any, _
    lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal _
    dwCreationFlags As Any, lpEnvironment As Any, ByVal _
    lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&

'Public Sub Main()
'    CreateProcessX "c:\windows\notepad.exe"
'End Sub

Public Function CreateProcessX(ByVal sPath As String) As Long
    Dim sInfo As STARTUPINFO
    Dim pInfo As PROCESS_INFORMATION
    
    lngReply = CreateProcess(vbNullString, sPath, ByVal 0&, ByVal 0&, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, sInfo, pInfo)
    CreateProcessX = pInfo.hProcess
End Function

Public Sub TerminateProcessX(ByVal hProcess As Long)
    Dim x As Long
    
    TerminateProcess hProcess, x
End Sub
