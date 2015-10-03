VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecadeYearPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQuestion

Private moPaper As Form
Private msWords As String

Private Sub Class_Initialize()
    msWords = "senunikanakiku"
End Sub

Public Function IQuestion_GenerateItem() As Variant
    IQuestion_GenerateItem = Array(Int(Rnd * 10), Int(Rnd * 10))
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = Format$(vItem(0) * 10 + vItem(1), "2000")
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound1 As Long
    Dim lFound2 As Long
    Dim lAnswer1 As Long
    Dim lAnswer2 As Long
    Dim lDecadeOffset As Long
    Dim lYearOffset As Long
    
    lFound1 = InStr(msWords, Left$(vAnswer, 2))
    lFound2 = InStr(msWords, Right$(vAnswer, 2))
    If Len(vAnswer) > 4 Or Len(vAnswer) < 4 Or lFound1 < 1 Or lFound2 < 1 Or (lFound1 Mod 2) = 0 Or (lFound2 Mod 2) = 0 Then
        IQuestion_Compare = False
    Else
        lAnswer1 = ((lFound1 - 1) \ 2)
        lAnswer2 = ((lFound2 - 1) \ 2)
        lDecadeOffset = (Weekday(DateSerial(2000 + vQuestion(0) * 10, 10, 7), vbFriday) - 1)
        lYearOffset = (((Weekday(DateSerial(2000 + vQuestion(0) * 10 + vQuestion(1), 10, 7), vbFriday) - 1) - lDecadeOffset) + 7) Mod 7
        IQuestion_Compare = lDecadeOffset = lAnswer1
        IQuestion_Compare = IQuestion_Compare And (lYearOffset = lAnswer2)
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 100
End Property




