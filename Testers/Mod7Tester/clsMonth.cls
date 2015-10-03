VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMonth"
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
    IQuestion_GenerateItem = Int(Rnd * 12) + 1
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = Mid$("janfebmaraprmayjunjulaugsepoctnovdec", vItem * 3 - 2, 3)
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound As Long
    Dim vRemap As Variant
    
    lFound = InStr(msWords, vAnswer)
    If lFound < 1 Or (lFound Mod 2) = 0 Then
        IQuestion_Compare = False
    Else
        vRemap = Array(0, 3, 3, 6, 1, 4, 6, 2, 5, 0, 3, 5)
        IQuestion_Compare = ((lFound - 1) \ 2) = vRemap(vQuestion - 1)
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 12
End Property

