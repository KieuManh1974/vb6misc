VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumberDay"
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

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    moPaper.lblNumber.Caption = vItem
End Sub

Public Function IQuestion_GenerateItem() As Variant
    Dim sNumberWords As String
    
    sNumberWords = "senunikanakiku"
    IQuestion_GenerateItem = Mid$(sNumberWords, Int(Rnd * 7) * 2 + 1, 2)
End Function


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim sDays As String
    Dim lFound As Long
    Dim sNumberWords As String
    Dim lFound2 As Long
    
    sNumberWords = "senunikanakiku"
    sDays = "frisatsunmontuewedthu"
    lFound = InStr(sDays, vAnswer)
    If Len(vAnswer) <> 3 Or lFound < 1 Or ((lFound - 1) Mod 3) <> 0 Then
        IQuestion_Compare = False
    Else
        lFound2 = InStr(sNumberWords, vQuestion)
        IQuestion_Compare = Mid$(sDays, (lFound2 \ 2) * 3 + 1, 3) = vAnswer
    End If
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 7
End Property

