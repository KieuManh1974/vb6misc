Attribute VB_Name = "Module1"
Option Explicit

Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Const VK_SNAPSHOT As Byte = &H2C
Public Const VK_MENU = &H12
Public Const KEYEVENTF_KEYUP = &H2
