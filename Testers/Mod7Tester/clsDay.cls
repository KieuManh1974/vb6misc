VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDay"
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
    IQuestion_GenerateItem = Int(Rnd * 31) + 1
    While (IQuestion_GenerateItem Mod 7 = 0) Or (IQuestion_GenerateItem Mod 7 = 1) Or (IQuestion_GenerateItem Mod 7 = 6)
        IQuestion_GenerateItem = Int(Rnd * 31) + 1
    Wend
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = vItem
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound As Long
    
    lFound = InStr(msWords, vAnswer)
    If lFound < 1 Or (lFound Mod 2) = 0 Then
        IQuestion_Compare = False
    Else
        IQuestion_Compare = (lFound - 1) \ 2 = ((vQuestion - 0) Mod 7)
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 31 - 4 - 5 - 4
End Property
