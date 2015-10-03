VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
lpStartupInfo As STARTUPINFO, lpProcessInformation As _
PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&
Private Const WM_USER = &H400&
   
'Const WM_KEYDOWN = &H100
'Const WM_KEYUP = &H101
'Const WM_CHAR = &H102
'
'' WM_KEYUP/DOWN/CHAR HIWORD(lParam) flags
'Const KF_EXTENDED = &H1000000
'Const KF_DLGMODE = &H8000000
'Const KF_MENUMODE = &H10000000
'Const KF_ALTDOWN = &H20000000
'Const KF_REPEAT = &H40000000
'Const KF_UP = &H80000000
'
'' Virtual Keys, Standard Set
'Const VK_LBUTTON = &H1
'Const VK_RBUTTON = &H2
'Const VK_CANCEL = &H3
'Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON
'
'Const VK_BACK = &H8
'Const VK_TAB = &H9
'
'Const VK_CLEAR = &HC
'Const VK_RETURN = &HD
'
'Const VK_SHIFT = &H10
'Const VK_CONTROL = &H11
'Const VK_MENU = &H12
'Const VK_PAUSE = &H13
'Const VK_CAPITAL = &H14
'
'Const VK_ESCAPE = &H1B
'
'Const VK_SPACE = &H20
'Const VK_PRIOR = &H21
'Const VK_NEXT = &H22
'Const VK_END = &H23
'Const VK_HOME = &H24
'Const VK_LEFT = &H25
'Const VK_UP = &H26
'Const VK_RIGHT = &H27
'Const VK_DOWN = &H28
'Const VK_SELECT = &H29
'Const VK_PRINT = &H2A
'Const VK_EXECUTE = &H2B
'Const VK_SNAPSHOT = &H2C
'Const VK_INSERT = &H2D
'Const VK_DELETE = &H2E
'Const VK_HELP = &H2F
'
'' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
'' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'
'
'Const VK_NUMPAD0 = &H60
'Const VK_NUMPAD1 = &H61
'Const VK_NUMPAD2 = &H62
'Const VK_NUMPAD3 = &H63
'Const VK_NUMPAD4 = &H64
'Const VK_NUMPAD5 = &H65
'Const VK_NUMPAD6 = &H66
'Const VK_NUMPAD7 = &H67
'Const VK_NUMPAD8 = &H68
'Const VK_NUMPAD9 = &H69
'Const VK_MULTIPLY = &H6A
'Const VK_ADD = &H6B
'Const VK_SEPARATOR = &H6C
'Const VK_SUBTRACT = &H6D
'Const VK_DECIMAL = &H6E
'Const VK_DIVIDE = &H6F
'Const VK_F1 = &H70
'Const VK_F2 = &H71
'Const VK_F3 = &H72
'Const VK_F4 = &H73
'Const VK_F5 = &H74
'Const VK_F6 = &H75
'Const VK_F7 = &H76
'Const VK_F8 = &H77
'Const VK_F9 = &H78
'Const VK_F10 = &H79
'Const VK_F11 = &H7A
'Const VK_F12 = &H7B
'Const VK_F13 = &H7C
'Const VK_F14 = &H7D
'Const VK_F15 = &H7E
'Const VK_F16 = &H7F
'Const VK_F17 = &H80
'Const VK_F18 = &H81
'Const VK_F19 = &H82
'Const VK_F20 = &H83
'Const VK_F21 = &H84
'Const VK_F22 = &H85
'Const VK_F23 = &H86
'Const VK_F24 = &H87
'
'Const VK_NUMLOCK = &H90
'Const VK_SCROLL = &H91
'
''
''   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
''   Used only as parameters to GetAsyncKeyState() and GetKeyState().
''   No other API or message will distinguish left and right keys in this way.
''  /
'Const VK_LSHIFT = &HA0
'Const VK_RSHIFT = &HA1
'Const VK_LCONTROL = &HA2
'Const VK_RCONTROL = &HA3
'Const VK_LMENU = &HA4
'Const VK_RMENU = &HA5
'
'Const VK_ATTN = &HF6
'Const VK_CRSEL = &HF7
'Const VK_EXSEL = &HF8
'Const VK_EREOF = &HF9
'Const VK_PLAY = &HFA
'Const VK_ZOOM = &HFB
'Const VK_NONAME = &HFC
'Const VK_PA1 = &HFD
'Const VK_OEM_CLEAR = &HFE

Private Const KEYEVENTF_KEYUP = &H2
Private Const VK_SHIFT = &H10
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Declare Function OemKeyScan Lib "user32" (ByVal wOemChar As Long) As Long
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Public Sub WriteText(ByVal sKeys As String)
    Dim VK As Integer
    Dim nShiftScan As Integer
    Dim nScan As Integer
    Dim sOemChar As String
    Dim nShiftKey As Integer
    Dim i As Integer
    For i = 1 To Len(sKeys)
        DoEvents
        'Loop through entire string being passed
        'and send each character individually.
        
        'Get the virtual key code for this character
        VK = VkKeyScan(Asc(Mid(sKeys, i, 1))) And &HFF
        
        'See if shift key needs to be pressed
        nShiftKey = VkKeyScan(Asc(Mid(sKeys, i, 1))) And 256
        sOemChar = " " '2 character buffer
        'Get the OEM character - preinitialize the buffer
        CharToOem Left$(Mid(sKeys, i, 1), 1), sOemChar
        'Get the nScan code for this key
        nScan = OemKeyScan(Asc(sOemChar)) And &HFF
        
        'Send the key down
        If nShiftKey = 256 Then
            'if shift key needs to be pressed
            nShiftScan = MapVirtualKey(VK_SHIFT, 0)
            'press down the shift key
            keybd_event VK_SHIFT, nShiftScan, 0, 0
        End If
        
        'press key to be sent
        keybd_event VK, nScan, 0, 0
        
        'Send the key up
        If nShiftKey = 256 Then
            'keyup for shift key
            keybd_event VK_SHIFT, nShiftScan, KEYEVENTF_KEYUP, 0
        End If
        
        'keyup for key sent
        keybd_event VK, nScan, KEYEVENTF_KEYUP, 0
    Next
End Sub

Private Sub Form_Load()
    Dim hNotepad As Long
    Dim sClassName As String
    Dim sWindowName As String
    Dim lParam As Long
    
    'ShellExecute "D:\Program Files\WinMX\WinMX.exe", "D:\Program Files\WinMX\"
    'hNotepad = ShellExecute("D:\Applications\VirtualDub 1.5.10\VirtualDub.exe", "D:\Applications\VirtualDub 1.5.10\")
    hNotepad = ShellExecute("C:\WINDOWS\SYSTEM32\NOTEPAD.EXE", "C:\WINDOWS\SYSTEM32\")
        
    WriteText "ABCDE"
    
'    sClassName = "VirtualDub"
'    sWindowName = "VirtualDub 1.5.10 (build 18160/release) by Avery Lee"
'    hNotepad = FindWindow(sClassName, sWindowName)
'    PostMessage hNotepad, WM_KEYDOWN, VK_F1, ByVal &H3B0001
'    'SendMessage hNotepad, WM_CHAR, 65, 0
'    PostMessage hNotepad, WM_KEYUP, VK_F1, ByVal KF_REPEAT + KF_UP + &H3B0001
    Unload Me
End Sub

Private Function ShellExecute(ByVal sPath As String, ByVal sFolderPath As String) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sCommand As String
    Dim wait As Long
    Dim x As Long
    
    start.cb = Len(start)
        
    If CreateProcessA(sPath, vbNullString, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, ByVal sFolderPath, start, proc) <> 0 Then
        WaitForSingleObject proc.hProcess, 5000
'        If TerminateProcess(proc.hProcess, x) <> 0 Then
'            Call CloseHandle(proc.hThread)
'            Call CloseHandle(proc.hProcess)
'        End If
    End If
End Function
