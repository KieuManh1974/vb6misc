Attribute VB_Name = "Module2"
Option Explicit

Private Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000

Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2

Private Const CREATE_ALWAYS As Long = 2
Private Const CREATE_NEW As Long = 1
Private Const OPEN_ALWAYS As Long = 4
Private Const OPEN_EXISTING As Long = 3
Private Const TRUNCATE_EXISTING As Long = 5

Private Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
Private Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
Private Const FILE_FLAG_DELETE_ON_CLOSE As Long = &H4000000
Private Const FILE_FLAG_NO_BUFFERING As Long = &H20000000
Private Const FILE_FLAG_OVERLAPPED As Long = &H40000000
Private Const FILE_FLAG_POSIX_SEMANTICS As Long = &H1000000
Private Const FILE_FLAG_RANDOM_ACCESS As Long = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN As Long = &H8000000
Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000

Public Sub CreateFileOnDisk(ByVal sFileNamePath As String)
    CreateFile sFileNamePath, GENERIC_READ Or GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0
End Sub
