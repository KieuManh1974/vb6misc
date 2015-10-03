VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ListBoxControl As ListBox
Attribute ListBoxControl.VB_VarHelpID = -1
Public TextBoxControl As TextBox

Private goSQLLines As New Dictionary
Private mlCurrentLine As Long

Public Sub Display()
    Dim vLine As Variant
    Dim lIndex As Long
    Dim oField As clsFieldInfo
    Dim oTable As clsTableInfo
    Dim oTable2 As clsTableInfo
    Dim oJoin As clsJoin
    Const sIndent = "    "
    Dim sType As String
    
    BuildSQL
    ListBoxControl.Clear
    
    For lIndex = 0 To goSQLLines.Count - 1
        Select Case VarType(goSQLLines.Items(lIndex).LineItem)
            Case vbString
                ListBoxControl.AddItem goSQLLines.Items(lIndex).LineItem
                ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
            Case vbObject
                Set vLine = goSQLLines.Items(lIndex)
                If TypeOf vLine.LineItem Is clsFieldInfo Then
                    Select Case vLine.LineType
                        Case sltSelect
                            Set oField = vLine.LineItem
                            ListBoxControl.AddItem sIndent & oField.TableInfo.Alias & "." & oField.Name & IIf(Not vLine.Last, ",", "")
                            ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                        Case sltWhere
                            Set oField = vLine.LineItem
                            Select Case vLine.LineItem.FieldType
                                Case adNumeric
                                    sType = "99999"
                                Case adVarChar
                                    sType = "'XXXX'"
                                Case adDBTimeStamp
                                    sType = "TO_DATE('01/01/1970',  'dd/mm/yyyy')"
                            End Select
                            ListBoxControl.AddItem sIndent & IIf(Not vLine.First, "AND ", "") & oField.TableInfo.Alias & "." & oField.Name & " = " & sType
                            ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                    End Select
                ElseIf TypeOf vLine.LineItem Is clsTableInfo Then
                    Set oTable = vLine.LineItem
                    ListBoxControl.AddItem sIndent & oTable.Name & " " & oTable.Alias & IIf(Not vLine.Last, ",", "")
                    ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                ElseIf TypeOf vLine.LineItem Is clsJoin Then
                    Set oJoin = vLine.LineItem
                    ListBoxControl.AddItem sIndent & IIf(Not vLine.First, "AND ", "") & oJoin.FromFieldFunction & IIf(oJoin.FromFieldFunction <> "", "(", "") & oJoin.FromField.TableInfo.Alias & "." & oJoin.FromField.Name & IIf(oJoin.JoinType = "RIGHT", " (+)", "") & IIf(oJoin.FromFieldFunction <> "", ")", "") & " = " & oJoin.ToFieldFunction & IIf(oJoin.ToFieldFunction <> "", "(", "") & oJoin.ToField.TableInfo.Alias & "." & oJoin.ToField.Name & IIf(oJoin.JoinType = "LEFT", " (+)", "") & IIf(oJoin.ToFieldFunction <> "", ")", "")
                    ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                ElseIf TypeOf vLine.LineItem Is clsExpression Then
                    Select Case vLine.LineType
                        Case sltSelect
                            ListBoxControl.AddItem sIndent & vLine.LineItem.Aliased
                            ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                        Case sltWhere
                            ListBoxControl.AddItem sIndent & IIf(Not vLine.First, "AND ", "") & vLine.LineItem.Aliased
                            ListBoxControl.ItemData(ListBoxControl.NewIndex) = lIndex
                    End Select
                End If
        End Select
    Next
    
    Dim sClipboard As String
    
    TextBoxControl.Text = ""
    For lIndex = 0 To ListBoxControl.ListCount - 1
        TextBoxControl.Text = TextBoxControl.Text & ListBoxControl.List(lIndex) & vbCrLf
        sClipboard = sClipboard & ListBoxControl.List(lIndex) & vbCrLf
    Next
    Clipboard.SetText sClipboard
End Sub

Private Sub BuildSQL()
    Dim oTable As clsTableInfo
    Dim oAliases As New Dictionary
    Dim vAlias As Variant
    Dim lCurrentLine As Long
    Dim oField As clsFieldInfo
    Dim vField As Variant
    Dim lIndex As Long
    Dim oJoinRS As Recordset
    Dim sSQL As String
    Dim oJoin As clsJoin
    Dim vTable As Variant
    Dim oSQLLine As clsSQLLine
    Dim bFirst As Boolean
    Dim bLast As Boolean
    Dim vFromField As Variant
    Dim vToField As Variant
    Dim sJoinType As String
    Dim oThisLine As clsSQLLine
    Dim oNextLine As clsSQLLine
    Dim sSelectedDatabaseShortName As String
    
    mlCurrentLine = 0
    
    Set goSQLLines = New Dictionary
    
    For Each vTable In goTables.SelectedTables.Items
        For Each vAlias In vTable.Aliases
            If Not oAliases.Exists(vAlias) Then
                vTable.Alias = vAlias
                oAliases.Add CStr(vAlias), 0
                Exit For
            End If
        Next
    Next
    
    AddSQLLine "SELECT", sltSelect
    For lIndex = 0 To UBound(goFields.SelectedFields.Items)
        AddSQLLine goFields.SelectedFields.Items(lIndex), sltSelect
    Next
    
    For lIndex = 0 To UBound(goFields.SelectedExpressions.Items)
        If goFields.SelectedExpressions.Items(lIndex).ExpressionType = etSelect Then
            AddSQLLine goFields.SelectedExpressions.Items(lIndex), sltSelect
        End If
    Next
    
    AddSQLLine "FROM", sltFrom
    For lIndex = 0 To UBound(goTables.SelectedTables.Items)
        AddSQLLine goTables.SelectedTables.Items(lIndex), sltFrom
    Next
    AddSQLLine "WHERE", sltWhere
    
    sSelectedDatabaseShortName = goTables.SelectedDatabase.SourceShortName & "."
    ' Joins
    bFirst = True
    For lIndex = 0 To UBound(goTables.SelectedTables.Items) - 1
        sSQL = ""
        sSQL = sSQL & " SELECT"
        sSQL = sSQL & "     *"
        sSQL = sSQL & " FROM"
        sSQL = sSQL & "     tblJoins"
        sSQL = sSQL & " WHERE"
        sSQL = sSQL & "     (FromTable = '" & sSelectedDatabaseShortName & goTables.SelectedTables.Items(lIndex).UniqueName & "'"
        sSQL = sSQL & "     AND ToTable = '" & sSelectedDatabaseShortName & goTables.SelectedTables.Items(lIndex + 1).UniqueName & "')"
        sSQL = sSQL & "    OR"
        sSQL = sSQL & "     (FromTable = '" & sSelectedDatabaseShortName & goTables.SelectedTables.Items(lIndex + 1).UniqueName & "'"
        sSQL = sSQL & "     AND ToTable = '" & sSelectedDatabaseShortName & goTables.SelectedTables.Items(lIndex).UniqueName & "')"

        Set oJoinRS = New Recordset
        oJoinRS.Open sSQL, goCon, adOpenForwardOnly, , adCmdText
        
        While Not oJoinRS.EOF
            Set oJoin = New clsJoin
            
            vFromField = ExtractFunctionAndField(oJoinRS!FromField)
            vToField = ExtractFunctionAndField(oJoinRS!ToField)
            sJoinType = oJoinRS!JoinType
            If oJoinRS!FromTable = sSelectedDatabaseShortName & goTables.SelectedTables.Items(lIndex).UniqueName Then
                Set oJoin.FromField = goFields.Findfield(vToField(0), goTables.SelectedTables.Items(lIndex + 1))
                Set oJoin.ToField = goFields.Findfield(vFromField(0), goTables.SelectedTables.Items(lIndex))
                oJoin.FromFieldFunction = vToField(1)
                oJoin.ToFieldFunction = vFromField(1)
                If sJoinType = "LEFT" Then
                    oJoin.JoinType = "RIGHT"
                ElseIf sJoinType = "RIGHT" Then
                    oJoin.JoinType = "LEFT"
                Else
                    oJoin.JoinType = "INNER"
                End If
            Else
                Set oJoin.FromField = goFields.Findfield(vFromField(0), goTables.SelectedTables.Items(lIndex + 1))
                Set oJoin.ToField = goFields.Findfield(vToField(0), goTables.SelectedTables.Items(lIndex))
                oJoin.FromFieldFunction = vFromField(1)
                oJoin.ToFieldFunction = vToField(1)
                oJoin.JoinType = sJoinType
            End If

            AddSQLLine oJoin, sltWhere
            bFirst = False
            oJoinRS.MoveNext
        Wend
    Next
    
    ' Expressions
    For lIndex = 0 To UBound(goFields.SelectedExpressions.Items)
        If goFields.SelectedExpressions.Items(lIndex).ExpressionType = etWhere Then
            AddSQLLine goFields.SelectedExpressions.Items(lIndex), sltWhere
        End If
    Next
            
    ' Filters
    For lIndex = 0 To UBound(goFields.SelectedFilters.Items)
        AddSQLLine goFields.SelectedFilters.Items(lIndex), sltWhere
        bFirst = False
    Next
    
    For lIndex = 0 To UBound(goSQLLines.Items)
        Set oThisLine = goSQLLines.Items(lIndex)
        If lIndex < UBound(goSQLLines.Items) Then
            Set oNextLine = goSQLLines.Items(lIndex + 1)
        Else
            Set oNextLine = Nothing
        End If
        
        If Not oNextLine Is Nothing Then
            If VarType(oNextLine.LineItem) = vbString Then
                If oNextLine.LineItem = "FROM" Or oNextLine.LineItem = "WHERE" Then
                    oThisLine.Last = True
                End If
            End If
            If VarType(oThisLine.LineItem) = vbString Then
                If oThisLine.LineItem = "WHERE" Then
                    oNextLine.First = True
                End If
            End If
        End If
    Next
End Sub

Private Sub AddSQLLine(vLine As Variant, lLineType As SQLLineTypes)
    Dim oSQLLine As clsSQLLine
    
    Set oSQLLine = New clsSQLLine
    oSQLLine.LineType = lLineType
    If VarType(vLine) = vbObject Then
        Set oSQLLine.LineItem = vLine
    Else
        oSQLLine.LineItem = vLine
    End If
    goSQLLines.Add CStr(mlCurrentLine), oSQLLine
    mlCurrentLine = mlCurrentLine + 1
End Sub

Private Function ExtractFunctionAndField(sDescriptor As String) As Variant
    Dim oTree As New ParseTree
    
    Stream.Text = sDescriptor
    If goParseFunction.Parse(oTree) Then
        ExtractFunctionAndField = Array(oTree(2).Text, oTree(1).Text)
    Else
        ExtractFunctionAndField = Array(sDescriptor, "")
    End If
End Function
