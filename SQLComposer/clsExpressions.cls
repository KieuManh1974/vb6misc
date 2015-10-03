VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExpressions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Expressions As New Dictionary
Public SelectedExpressions As New Dictionary

Public Sub PopulateExpressions()
    Dim sSQL As String
    Dim oExpressions As New Recordset
    Dim oExpression As clsExpression
    Dim lIndex As Long
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     *"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     tblExpressions"

    oExpressions.Open sSQL, goCon, adOpenForwardOnly, , adCmdText
    
    While Not oExpressions.EOF
        Set oExpression = New clsExpression
        oExpression.Name = oExpressions!Name
        oExpression.Expression = oExpressions!Expression
        Select Case oExpressions!Type
            Case "WHERE"
                oExpression.ExpressionType = etWhere
            Case "SELECT"
                oExpression.ExpressionType = etSelect
        End Select
        Expressions.Add CStr(lIndex), oExpression
        lIndex = lIndex + 1
        oExpressions.MoveNext
    Wend
End Sub
