VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQuestion

Private moPaper As Form
Private msWords As String
Private msDays As String

Private Sub Class_Initialize()
    msWords = "senunikanakiku"
    msDays = "frisatsunmontuewedthu"
End Sub

Public Function IQuestion_GenerateItem() As Variant
    IQuestion_GenerateItem = Int(Rnd * 365)
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = Format$(CDate(vItem + CLng(CDate("1 jan 2005"))), "DD MMM")
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim lFound As Long
    Dim vRemap As Variant

    lFound = InStr(msDays, vAnswer)
    If lFound < 1 Or (lFound Mod 3) <> 1 Then
        IQuestion_Compare = False
    Else
        IQuestion_Compare = ((lFound - 1) \ 3) = ((vQuestion + 1) Mod 7)
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 365
End Property



