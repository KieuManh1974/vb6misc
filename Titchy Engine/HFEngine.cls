VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HFEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Memory As String
Public PC As Long
Public MC As Long

Public Sub Execute()
    Dim iOffset As Long
    Dim sInstruction As String
    Dim iTemp As Long
    
'    If Int(Rnd * 200) = 123 Then
'        Mid$(Memory, Int(Rnd * Len(Memory)), 1) = CStr(Int(Rnd * 2))
'    End If
    sInstruction = LoadBits(PC)
    PC = PC + 2
    Select Case sInstruction
        Case "00"
            Mid$(Memory, MC, 1) = CStr(1 - Val(Mid$(Memory, MC, 1)))
        Case "01"
            Mid$(Memory, MC, 1) = CStr(1 - Val(Mid$(Memory, MC, 1)))
            MC = MC + LoadAddress
        Case "10"
            iTemp = PC
            iOffset = LoadAddress
            If Mid$(Memory, MC, 1) = "1" Then
                If iOffset > -1 Then
                    PC = PC + LoadAddress
                Else
                    PC = iTemp + LoadAddress + 1
                End If
            End If
        Case "11"
    End Select
    Adjust MC
    Adjust PC
End Sub

Private Function LoadAddress() As Long
    Dim iStart As Long
    Dim iEnd As Long
    Dim iOffset As Long
    Dim sValue As String
    
    Do
        sValue = LoadBits(PC)
        PC = PC + 2
        Select Case sValue
            Case "00"
                LoadAddress = iOffset
                Exit Function
            Case "01"
                iOffset = 2 * iOffset + 0
            Case "10"
                iOffset = 2 * iOffset + 1
            Case "11"
                LoadAddress = -iOffset
                Exit Function
        End Select
        If iOffset > 2048 Then
            Exit Do
        End If
    Loop
End Function

Private Sub Class_Initialize()
    PC = 1
    MC = 1
End Sub

Private Function LoadBits(iLocation) As String
    iLocation = ((iLocation - 1) Mod Len(Memory)) + 1
    LoadBits = Mid$(Memory & Memory, iLocation, 2)
End Function

Private Sub Adjust(iLocation)
    iLocation = iLocation - 1
    While iLocation < 0
        iLocation = iLocation + Len(Memory)
    Wend
    iLocation = iLocation Mod Len(Memory)
    iLocation = iLocation + 1
End Sub
