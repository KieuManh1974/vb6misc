VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecade"
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
    IQuestion_GenerateItem = Int(Rnd * 10)
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = Format$(vItem * 10 + 2000, "00")
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound As Long
    Dim vRemap As Variant
    
    lFound = InStr(msWords, vAnswer)
    If lFound < 1 Or (lFound Mod 2) = 0 Then
        IQuestion_Compare = False
    Else
        IQuestion_Compare = ((lFound - 1) \ 2) = (Weekday(CDate("1 oct " & vQuestion * 10 + 2000), vbSaturday) - 1)
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 10
End Property



