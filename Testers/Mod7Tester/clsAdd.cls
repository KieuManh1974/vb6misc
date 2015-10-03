VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQuestion

Private moPaper As Form
Private msWords As String
Private mlCount As Long

Private Sub Class_Initialize()
    msWords = "senunikanakiku"
    mlCount = 2
End Sub

Public Function IQuestion_GenerateItem() As Variant
    Dim vQuestion As Variant
    Dim lIndex As Long
    
    vQuestion = Array()
    ReDim vQuestion(mlCount - 1)
    
    For lIndex = 0 To mlCount - 1
        vQuestion(lIndex) = Int(Rnd * 7)
    Next
    IQuestion_GenerateItem = vQuestion
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    Dim lIndex As Long
    Dim sCaption As String
    
    For lIndex = 0 To mlCount - 1
        sCaption = sCaption & Mid$(msWords, vItem(lIndex) * 2 + 1, 2)
    Next
    moPaper.lblNumber.Caption = sCaption
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound As Long
    Dim lIndex As Long
    Dim lSum As Long
    Dim lAnswer As Long
    
    If Len(vAnswer) <> 2 Then
        Exit Function
    End If

    lFound = InStr(msWords, vAnswer)
    If lFound < 1 Or (lFound Mod 2) = 0 Then
        Exit Function
    End If
    lAnswer = (lFound - 1) \ 2
    
    For lIndex = 0 To mlCount - 1
        lSum = lSum + vQuestion(lIndex)
    Next
    IQuestion_Compare = (lSum Mod 7) = lAnswer
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 7 ^ mlCount
End Property


