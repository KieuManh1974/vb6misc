VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDatabases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Databases As Dictionary
Public SelectedDatabases As Dictionary
Public ListBoxControl As ListBox

Private mlCurrentIdentifier As Long

Private Function CreateDatabase(sSourceShortName As String, lSourceID As Long) As clsDatabaseInfo
    Dim oDatabase As New clsDatabaseInfo
    
    With oDatabase
        .SourceID = lSourceID
        .SourceShortName = sSourceShortName
        .Identifier = mlCurrentIdentifier
        mlCurrentIdentifier = mlCurrentIdentifier + 1
    End With
    Set CreateDatabase = oDatabase
End Function

Private Sub AddDatabase(oDatabase As clsDatabaseInfo)
    If Not Databases.Exists(oDatabase.UniqueName) Then
        Databases.Add oDatabase.UniqueName, oDatabase
    End If
End Sub

Private Sub RemoveDatabase(oDatabase As clsDatabaseInfo)
    If Databases.Exists(oDatabase.UniqueName) Then
        Databases.Remove oDatabase.UniqueName
    End If
End Sub

Public Sub PopulateDatabases()
    Dim oDatabases As New Recordset
    Dim oExpressions As New Recordset
    Dim sSQL As String
    Dim oDatabase As clsDatabaseInfo
    
    Set Databases = New Dictionary
    Set SelectedDatabases = New Dictionary

    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     S.ShortName,"
    sSQL = sSQL & "     S.SourceID"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     tblSources S"

    oDatabases.Open sSQL, goCon, adOpenForwardOnly, , adCmdText
    
    While Not oDatabases.EOF
        Set oDatabase = CreateDatabase(oDatabases!ShortName, oDatabases!SourceID)
        AddDatabase oDatabase
        ListBoxControl.AddItem oDatabase.DisplayName
        ListBoxControl.ItemData(ListBoxControl.NewIndex) = oDatabase.Identifier
        oDatabases.MoveNext
    Wend
End Sub

Public Function GetDatabaseByIdentifier(lIdentifier As Long) As clsDatabaseInfo
    Dim vDatabase As Variant
    
    For Each vDatabase In Databases.Items
        If vDatabase.Identifier = lIdentifier Then
            Set GetDatabaseByIdentifier = vDatabase
            Exit Function
        End If
    Next
End Function

Public Sub ToggleSelection(oDatabase As clsDatabaseInfo)
    If SelectedDatabases.Exists(oDatabase.UniqueName) Then
        SelectedDatabases.Remove oDatabase.UniqueName
    Else
        SelectedDatabases.Add oDatabase.UniqueName, oDatabase
    End If
End Sub

'Public Function FindDatabase(sDatabase As String, sDatabase As String) As clsDatabaseInfo
'    Dim vDatabase As Variant
'
'    For Each vDatabase In Databases.Items
'        If vDatabase.SourceShortName = sDatabase Then
'            If vDatabase.Name = sDatabase Then
'                Set FindDatabase = vDatabase
'            End If
'        End If
'    Next
'End Function

