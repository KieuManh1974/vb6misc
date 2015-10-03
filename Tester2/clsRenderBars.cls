VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderBars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRender

Private moPaper As Form

Public Property Set IRender_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IRender_Render(vItem As Variant)
    Dim lBarIndex As Long
    
    moPaper.Cls
    
    For lBarIndex = 0 To vItem - 1
        moPaper.Line (40, 120 - (lBarIndex + lBarIndex \ 3) * 8)-Step(30, 2), vbBlack, BF
    Next
End Sub
