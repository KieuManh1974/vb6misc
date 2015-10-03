VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLPhrase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public oConnection As Connection

Public Function FieldAlias(ByVal sText As String, ByVal lPosition As Long, lIndex As Long, sResultText, ByVal vParameters) As Boolean
    Dim sSQL As String
    Dim oFieldAlias As Recordset
    
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     *"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     FieldAlias"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     Alias = '" & vParameters(0) & "'"
    sSQL = sSQL & "     AND Table = '" & vParameters(1) & "'"
    
    Set oFieldAlias = New Recordset
    oFieldAlias.Open sSQL, oConnection, adOpenForwardOnly, adLockOptimistic
End Function

Private Sub Class_Initialize()
    Set oConnection = New Connection
    oConnection.Open "Provider=Microsoft.Access.OLEDB.10.0;Persist Security Info=False;Data Source=" & App.Path & "\englishsql.mdb;User ID=Admin;Data Provider=Microsoft.Jet.OLEDB.4.0"
End Sub

Private Sub Class_Terminate()
    oConnection.Close
    Set oConnection = Nothing
End Sub
