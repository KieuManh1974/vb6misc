Attribute VB_Name = "KeyboardHandler"
Option Explicit

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadID As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal cb As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type KBDLLHOOKSTRUCT
  vkCode As Long
  scanCode As Long
  flags As Long
  time As Long
  dwExtraInfo As Long
End Type

' Low-Level Keyboard Constants
Private Const HC_ACTION = 0
Private Const LLKHF_EXTENDED = &H1
Private Const LLKHF_INJECTED = &H10
Private Const LLKHF_ALTDOWN = &H20
Private Const LLKHF_UP = &H80

' Virtual Keys
Public Const VK_TAB = &H9
Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_ESCAPE = &H1B
Public Const VK_DELETE = &H2E

Private Const WH_KEYBOARD_LL = 13&

Public KeyboardHandle As Long

Public KeyboardHook As IKeyboardHook

Public Function KeyboardCallback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Static Hookstruct As KBDLLHOOKSTRUCT
    Dim lShift As Long
    
    If (Code = HC_ACTION) Then
        ' Copy the keyboard data out of the lParam (which is a pointer)
        Call CopyMemory(Hookstruct, ByVal lParam, Len(Hookstruct))
         
        If (Hookstruct.flags And LLKHF_UP) = 0 Then
            If CBool(GetAsyncKeyState(VK_CONTROL) And &H8000) Then
                lShift = lShift Or 2
            End If
            
            If CBool(GetAsyncKeyState(VK_SHIFT) And &H8000) Then
                lShift = lShift Or 1
            End If
            
            If KeyboardHook.ProcessKey(Hookstruct.vkCode, lShift) Then
                KeyboardCallback = 1
                Exit Function
            End If
        End If
        
        KeyboardCallback = CallNextHookEx(KeyboardHandle, Code, wParam, lParam)
    End If
End Function

Public Sub HookKeyboard()
  KeyboardHandle = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf KeyboardCallback, App.hInstance, 0&)

  Call CheckHooked
End Sub

Public Sub CheckHooked()
  If (Hooked) Then
    Debug.Print "Keyboard hooked"
  Else
    Debug.Print "Keyboard hook failed: " & Err.LastDllError
  End If
End Sub

Private Function Hooked()
  Hooked = KeyboardHandle <> 0
End Function

Public Sub UnhookKeyboard()
  If (Hooked) Then
    Call UnhookWindowsHookEx(KeyboardHandle)
  End If
End Sub

