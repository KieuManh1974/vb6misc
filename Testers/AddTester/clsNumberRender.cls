VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumberRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRender

Private moPaper As frmPaper

Public Property Set IRender_Paper(oForm As Form)
    Set moPaper = oForm
End Property

Public Sub IRender_Render(loGenerator As IGenerator)
    Dim vItem As Variant
    
    vItem = loGenerator.GetItem
    
    moPaper.Cls
    moPaper.lblOperand(0) = vItem(0)
    moPaper.lblOperand(1) = vItem(1)
End Sub
