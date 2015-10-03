VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParseTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SubTree As New SubTrees
Attribute SubTree.VB_VarMemberFlags = "200"
Attribute SubTree.VB_VarUserMemId = 0
Public Text As String
Public Index As Integer
Public ParentIndex As Integer
Public Control As Long
Public TextStart As Long
Public TextEnd As Long

Public Name As String
Public Description As String

Public Location As String

Public Error As Boolean

