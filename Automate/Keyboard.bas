Attribute VB_Name = "Keyboard"
Option Explicit

Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2

Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function VkKeyScanW Lib "user32" (ByVal cChar As Integer) As Integer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)


' Virtual Keys, Standard Set
Public Const VK_LBUTTON = &H1
Public Const VK_RBUTTON = &H2
Public Const VK_CANCEL = &H3
Public Const VK_MBUTTON = &H4             '  NOT contiguous with L RBUTTON

Public Const VK_BACK = &H8
Public Const VK_TAB = &H9

Public Const VK_CLEAR = &HC
Public Const VK_RETURN = &HD

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11
Public Const VK_MENU = &H12
Public Const VK_PAUSE = &H13
Public Const VK_CAPITAL = &H14

Public Const VK_ESCAPE = &H1B

Public Const VK_SPACE = &H20
Public Const VK_PRIOR = &H21
Public Const VK_NEXT = &H22
Public Const VK_END = &H23
Public Const VK_HOME = &H24
Public Const VK_LEFT = &H25
Public Const VK_UP = &H26
Public Const VK_RIGHT = &H27
Public Const VK_DOWN = &H28
Public Const VK_SELECT = &H29
Public Const VK_PRINT = &H2A
Public Const VK_EXECUTE = &H2B
Public Const VK_SNAPSHOT = &H2C
Public Const VK_INSERT = &H2D
Public Const VK_DELETE = &H2E
Public Const VK_HELP = &H2F

' VK_A thru VK_Z are the same as their ASCII equivalents: 'A' thru 'Z'
' VK_0 thru VK_9 are the same as their ASCII equivalents: '0' thru '9'

Public Const VK_NUMPAD0 = &H60
Public Const VK_NUMPAD1 = &H61
Public Const VK_NUMPAD2 = &H62
Public Const VK_NUMPAD3 = &H63
Public Const VK_NUMPAD4 = &H64
Public Const VK_NUMPAD5 = &H65
Public Const VK_NUMPAD6 = &H66
Public Const VK_NUMPAD7 = &H67
Public Const VK_NUMPAD8 = &H68
Public Const VK_NUMPAD9 = &H69
Public Const VK_MULTIPLY = &H6A
Public Const VK_ADD = &H6B
Public Const VK_SEPARATOR = &H6C
Public Const VK_SUBTRACT = &H6D
Public Const VK_DECIMAL = &H6E
Public Const VK_DIVIDE = &H6F
Public Const VK_F1 = &H70
Public Const VK_F2 = &H71
Public Const VK_F3 = &H72
Public Const VK_F4 = &H73
Public Const VK_F5 = &H74
Public Const VK_F6 = &H75
Public Const VK_F7 = &H76
Public Const VK_F8 = &H77
Public Const VK_F9 = &H78
Public Const VK_F10 = &H79
Public Const VK_F11 = &H7A
Public Const VK_F12 = &H7B
Public Const VK_F13 = &H7C
Public Const VK_F14 = &H7D
Public Const VK_F15 = &H7E
Public Const VK_F16 = &H7F
Public Const VK_F17 = &H80
Public Const VK_F18 = &H81
Public Const VK_F19 = &H82
Public Const VK_F20 = &H83
Public Const VK_F21 = &H84
Public Const VK_F22 = &H85
Public Const VK_F23 = &H86
Public Const VK_F24 = &H87

Public Const VK_NUMLOCK = &H90
Public Const VK_SCROLL = &H91

'   VK_L VK_R - left and right Alt, Ctrl and Shift virtual keys.
'   Used only as parameters to GetAsyncKeyState() and GetKeyState().
'   No other API or message will distinguish left and right keys in this way.

Public Const VK_LSHIFT = &HA0
Public Const VK_RSHIFT = &HA1
Public Const VK_LCONTROL = &HA2
Public Const VK_RCONTROL = &HA3
Public Const VK_LMENU = &HA4
Public Const VK_RMENU = &HA5

Public Const VK_ATTN = &HF6
Public Const VK_CRSEL = &HF7
Public Const VK_EXSEL = &HF8
Public Const VK_EREOF = &HF9
Public Const VK_PLAY = &HFA
Public Const VK_ZOOM = &HFB
Public Const VK_NONAME = &HFC
Public Const VK_PA1 = &HFD
Public Const VK_OEM_CLEAR = &HFE


Public Sub KeyDown(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY, 0
End Sub

Public Sub KeyUp(ByVal vKey As KeyCodeConstants)
   keybd_event vKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Sub

Public Sub KeyPress(ByVal Key As KeyCodeConstants)
    Call keybd_event(Key, 0, 0 Or 0, 0)
    Call keybd_event(Key, 0, 0 Or KEYEVENTF_KEYUP, 0)
End Sub


Public Function KeyCode(ByVal sChar As String) As KeyCodeConstants
    Dim bNt As Boolean
    Dim iKeyCode As Integer
    Dim b() As Byte
    Dim iKey As Integer
    Dim vKey As KeyCodeConstants
    Dim iShift As ShiftConstants

    ' Determine if we have Unicode support or not:
    bNt = ((GetVersion() And &H80000000) = 0)
   
    ' Get the keyboard scan code for the character:
    If (bNt) Then
        b = sChar
        CopyMemory iKey, b(0), 2
        iKeyCode = VkKeyScanW(iKey)
    Else
        b = StrConv(sChar, vbFromUnicode)
        iKeyCode = VkKeyScan(b(0))
    End If
   
    KeyCode = (iKeyCode And &HFF&)
End Function

Private Sub SetKeyState(ByVal Key As Long, ByVal State As Boolean)
    Dim Keys(0 To 255) As Byte
    
    Call GetKeyboardState(Keys(0))
    Keys(Key) = Abs(CInt(State))
    Call SetKeyboardState(Keys(0))
End Sub

Public Sub AllKeys()
    Dim x As Long
    For x = 0 To 255
        SetKeyState x, False
    Next
End Sub

Public Property Get CapsLock() As Boolean
    CapsLock = GetKeyState(KeyCodeConstants.vbKeyCapital) = 1
End Property

Public Property Let CapsLock(ByVal Value As Boolean)
    Call SetKeyState(KeyCodeConstants.vbKeyCapital, Value)
End Property

Private Property Get NumLock() As Boolean
    NumLock = GetKeyState(KeyCodeConstants.vbKeyNumlock) = 1
End Property

Private Property Let NumLock(ByVal Value As Boolean)
    Call SetKeyState(KeyCodeConstants.vbKeyNumlock, Value)
End Property

Private Property Get ScrollLock() As Boolean
    ScrollLock = GetKeyState(KeyCodeConstants.vbKeyScrollLock) = 1
End Property

Private Property Let ScrollLock(ByVal Value As Boolean)
    Call SetKeyState(KeyCodeConstants.vbKeyScrollLock, Value)
End Property








