VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRenderBlocks"
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
    Dim lIndex As Long
    
    For lIndex = 0 To 6
        moPaper.pctBlock(lIndex).BackColor = vbButtonFace
    Next
    
    For lIndex = 0 To UBound(vItem)
        moPaper.pctBlock(vItem(lIndex)).BackColor = vbBlack
    Next
End Sub
