VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Identifier As Long
Public Fields As Dictionary
Public Expressions As clsExpressions
Public SelectedFields As New Dictionary
Public SelectedFilters As New Dictionary
Public SelectedExpressions As New Dictionary
Public ListBoxControl As ListBox
Attribute ListBoxControl.VB_VarHelpID = -1

Private moSelectedTable As clsTableInfo
Private mlCurrentListIdentifier As Long
Private mlCurrentFieldIdentifier As Long
Public ListLines As New Collection

Public Property Set SelectedTable(oTable As clsTableInfo)
    Set moSelectedTable = oTable
    PopulateFields
End Property

Public Property Get SelectedTable() As clsTableInfo
    Set SelectedTable = moSelectedTable
End Property

Public Function CreateField(sName As String, oTable As clsTableInfo, lFieldType As Long) As clsFieldInfo
    Dim oField As New clsFieldInfo
    
    With oField
        .Name = sName
        Set .TableInfo = oTable
        .Identifier = mlCurrentFieldIdentifier
        .FieldType = lFieldType
    End With
    mlCurrentFieldIdentifier = mlCurrentFieldIdentifier + 1
    
    Set CreateField = oField
End Function

Public Function Findfield(ByVal sFieldName As String, ByVal oTable As clsTableInfo) As clsFieldInfo
    Dim vField As Variant
    
    For Each vField In Fields.Items
        If vField.Name = sFieldName And vField.TableInfo Is oTable Then
            Set Findfield = vField
            Exit Function
        End If
    Next
    Set Findfield = CreateField(sFieldName, oTable, 0)
End Function

Public Sub PopulateFields()
    Dim sSQL As String
    Dim oSources As New Recordset
    Dim oRecordFields As New Recordset
    Dim oRecordField As Field
    Dim oField As clsFieldInfo
    Dim vExpression As Variant
    Dim oIdentifier As clsIdentifier
    Dim bIncludeExpression As Boolean
    Dim oFieldListLine As clsFieldListLine
    
    Set Fields = New Dictionary
    
    ListBoxControl.Clear
    If Not moSelectedTable Is Nothing Then
        Set ListLines = New Collection
        mlCurrentListIdentifier = 1
        
        For Each vExpression In goExpressions.Expressions.Items
            bIncludeExpression = False
            For Each oIdentifier In vExpression.Identifiers
                If oIdentifier.Table = moSelectedTable.Name Or oIdentifier.Table = "ANY" Then
                    Set oFieldListLine = New clsFieldListLine
                    oFieldListLine.Identifier = mlCurrentListIdentifier
                    oFieldListLine.LineType = fltExpression
                    Set oFieldListLine.Expression = vExpression
                    ListBoxControl.AddItem "{" & vExpression.Name & "}"
                    ListBoxControl.ItemData(ListBoxControl.NewIndex) = mlCurrentListIdentifier
                    ListLines.Add oFieldListLine
                    mlCurrentListIdentifier = mlCurrentListIdentifier + 1
                End If
            Next
        Next
        
        sSQL = ""
        sSQL = sSQL & " SELECT"
        sSQL = sSQL & "     Source"
        sSQL = sSQL & " FROM"
        sSQL = sSQL & "     tblSources"
        sSQL = sSQL & " WHERE"
        sSQL = sSQL & "     SourceID = " & moSelectedTable.SourceID
        
        oSources.Open sSQL, goCon, adOpenForwardOnly, , adCmdText
        
        oRecordFields.Open "SELECT * FROM " & moSelectedTable.Name & " WHERE 1=0", CStr(oSources!Source), adOpenForwardOnly, , adCmdText
        
        mlCurrentFieldIdentifier = NewIdentifier
        For Each oRecordField In oRecordFields.Fields
            Set oField = CreateField(oRecordField.Name, moSelectedTable, oRecordField.Type)
            Fields.Add oField.UniqueName, oField
            
            Set oFieldListLine = New clsFieldListLine
            Set oFieldListLine.Field = oField
            oFieldListLine.LineType = fltField
            oFieldListLine.Identifier = mlCurrentListIdentifier
            
            ListBoxControl.AddItem oField.DisplayName
            ListBoxControl.ItemData(ListBoxControl.NewIndex) = mlCurrentListIdentifier
            ListLines.Add oFieldListLine
            mlCurrentListIdentifier = mlCurrentListIdentifier + 1
        Next
    End If
End Sub

Private Function NewIdentifier() As Long
    Dim vField As Variant
    
    NewIdentifier = -1
    For Each vField In SelectedFields.Items
        If vField.Identifier > NewIdentifier Then
            NewIdentifier = vField.Identifier
        End If
    Next
    NewIdentifier = NewIdentifier + 1
End Function

Public Function GetLineByIdentifier(lIdentifier As Long) As clsFieldListLine
    Dim vField As Variant
    
    For Each vField In ListLines
        If vField.Identifier = lIdentifier Then
            Set GetLineByIdentifier = vField
            Exit Function
        End If
    Next
End Function


Public Sub ToggleSelection(oField As clsFieldInfo)
    If SelectedFields.Exists(oField.UniqueName) Then
        SelectedFields.Remove oField.UniqueName
    Else
        SelectedFields.Add oField.UniqueName, oField
    End If
End Sub

Public Sub ToggleFilter(oField As clsFieldInfo)
    If SelectedFilters.Exists(oField.UniqueName) Then
        SelectedFilters.Remove oField.UniqueName
    Else
        SelectedFilters.Add oField.UniqueName, oField
    End If
End Sub

Public Sub ToggleExpression(oExpression As clsExpression)
    If SelectedExpressions.Exists(oExpression.UniqueIdentifier) Then
        SelectedExpressions.Remove oExpression.UniqueIdentifier
    Else
        SelectedExpressions.Add oExpression.UniqueIdentifier, oExpression
    End If
End Sub
