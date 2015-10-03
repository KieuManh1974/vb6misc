VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Tables As Dictionary
Public SelectedTables As Dictionary
Public ListBoxControl As ListBox

Private mlCurrentIdentifier As Long
Private moSelectedDatabase As clsDatabaseInfo

Public Property Set SelectedDatabase(oDatabase As clsDatabaseInfo)
    Set moSelectedDatabase = oDatabase
    PopulateTables
End Property

Public Property Get SelectedDatabase() As clsDatabaseInfo
    Set SelectedDatabase = moSelectedDatabase
End Property

Private Function CreateTable(sName As String, lSourceID As Long, sAliases As Variant) As clsTableInfo
    Dim oTable As New clsTableInfo
    
    With oTable
        .Name = sName
        .SourceID = lSourceID
        .Aliases = Split(sAliases, " ")
        .Identifier = mlCurrentIdentifier
        mlCurrentIdentifier = mlCurrentIdentifier + 1
    End With
    Set CreateTable = oTable
End Function

Private Sub AddTable(oTable As clsTableInfo)
    If Not Tables.Exists(oTable.UniqueName) Then
        Tables.Add oTable.UniqueName, oTable
    End If
End Sub

Private Sub RemoveTable(oTable As clsTableInfo)
    If Tables.Exists(oTable.UniqueName) Then
        Tables.Remove oTable.UniqueName
    End If
End Sub

Public Sub PopulateTables()
    Dim oTables As New Recordset
    Dim oExpressions As New Recordset
    Dim sSQL As String
    Dim oTable As clsTableInfo
    
    Set Tables = New Dictionary
    Set SelectedTables = New Dictionary

    ListBoxControl.Clear
    
    If Not moSelectedDatabase Is Nothing Then
        sSQL = ""
        sSQL = sSQL & " SELECT"
        sSQL = sSQL & "     A.TableName,"
        sSQL = sSQL & "     A.SourceID,"
        sSQL = sSQL & "     A.Aliases"
        sSQL = sSQL & " FROM"
        sSQL = sSQL & "     tblAliases A"
        sSQL = sSQL & " WHERE"
        sSQL = sSQL & "     A.SourceID = " & moSelectedDatabase.SourceID
        
        oTables.Open sSQL, goCon, adOpenForwardOnly, , adCmdText
        
        While Not oTables.EOF
            Set oTable = CreateTable(oTables!TableName, oTables!SourceID, oTables!Aliases)
            AddTable oTable
            ListBoxControl.AddItem oTable.DisplayName
            ListBoxControl.ItemData(ListBoxControl.NewIndex) = oTable.Identifier
            oTables.MoveNext
        Wend
    End If
End Sub

Public Function GetTableByIdentifier(lIdentifier As Long) As clsTableInfo
    Dim vTable As Variant
    
    For Each vTable In Tables.Items
        If vTable.Identifier = lIdentifier Then
            Set GetTableByIdentifier = vTable
            Exit Function
        End If
    Next
End Function

Public Sub ToggleSelection(oTable As clsTableInfo)
    If SelectedTables.Exists(oTable.UniqueName) Then
        SelectedTables.Remove oTable.UniqueName
    Else
        SelectedTables.Add oTable.UniqueName, oTable
    End If
End Sub

Public Function FindTable(sTable As String, sDatabase As String) As clsTableInfo
    Dim vTable As Variant
    
    For Each vTable In Tables.Items
        If vTable.SourceShortName = sDatabase Then
            If vTable.Name = sTable Then
                Set FindTable = vTable
            End If
        End If
    Next
End Function
