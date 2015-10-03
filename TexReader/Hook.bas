Attribute VB_Name = "Hook"
Option Explicit

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Public Const WH_MIN = (-1)
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_MOUSE = 7
Public Const WH_HARDWARE = 8
Public Const WH_DEBUG = 9
Public Const WH_SHELL = 10
Public Const WH_FOREGROUNDIDLE = 11
Public Const WH_MAX = 11

Public mhHook As Long

Public Function MouseProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'If wParam = 522 Then
        Debug.Print ncode & ":" & wParam & ":" & lParam
    'End If
    
    MouseProc = CallNextHookEx(mhHook, ncode, wParam, lParam)
End Function
