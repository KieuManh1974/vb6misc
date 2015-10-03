VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumberQuestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IQuestion

Private moPaper As Form

Public Function IQuestion_GenerateItem() As Variant
    Dim x&

    For x& = 1 To 2
        IQuestion_GenerateItem = IQuestion_GenerateItem & CStr(Int(Rnd * 10))
    Next
End Function

Public Property Set IQuestion_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IQuestion_Render(vItem As Variant)
    Dim lBarIndex As Long
    Dim x&
    Dim show$
    Dim v
    
    v = Array("a", "e", "i", "o", "u", "aa", "ee", "ii", "oo", "uu", "k", "g", "t", "d", "s", "z", "f", "v", "p", "b")
    
    For x& = 1 To Len(vItem)
        show$ = show$ & CStr(v(Val(Mid$(vItem, x&, 1)) + (x Mod 2) * 10))
    Next
    moPaper.lblNumber.Caption = show$
End Sub


Public Function IQuestion_Compare(vQuestion As Variant, vAnswer As Variant) As Boolean
    Dim index&
    Dim answer$
    
'    For index& = 1 To Len(vQuestion)
'        answer$ = answer$ & Trim$(Mid$("a e i o u aaeeiioouuk g t d s z f v p b", 1 + Val(Mid$(vQuestion, index&, 1)) * 2 + ((index& + 0) Mod 2) * 20, 2))
'    Next
    
    IQuestion_Compare = LCase$(vAnswer) = vQuestion
End Function


Public Property Get IQuestion_SetSize() As Long
    IQuestion_SetSize = 100
End Property
