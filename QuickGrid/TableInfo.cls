VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TableInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mlCellHeight As Long
Public mlCellColour As Long
Public mlCellOffsetLeft As Long
Public mlCellOffsetTop As Long
Public mlCellSeparator As Long
Public mvTabs As Variant
Public moTable As Table

Private Sub Class_Initialize()
    mvTabs = Array()
End Sub
