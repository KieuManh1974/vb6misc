VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvTable As Variant
Private msText As String
Private mlLastColumn As Long
Public msRowDelimiter
Public msColumnDelimiter As String
Public mlTotalColumns As Long
Public mlTotalRows As Long
Public moParsing As Parsing

Private mlTextStamp As Long
Private mlTableStamp As Long
Private mlLastColumnStamp As Long

Public Property Let Table(ByVal vTable As Variant)
    mvTable = vTable
    StampTableChange
End Property

Public Property Let TableCell(ByVal lRow As Long, ByVal lColumn As Long, ByVal sValue As String)
    mvTable(lRow)(lColumn) = sValue
    StampTableChange
End Property

Public Property Let TableRow(ByVal lRow As Long, ByVal vValue As Variant)
    mvTable(lRow) = vValue
    StampTableChange
End Property

Public Property Let Text(ByVal sText As String)
    msText = sText
    StampTextChange
End Property

Public Property Get Table() As Variant
    CheckTextChange
    Table = mvTable
End Property

Public Property Get TableCell(ByVal lRow As Long, ByVal lColumn As Long) As String
    CheckTextChange
    TableCell = mvTable(lRow)(lColumn)
End Property

Public Property Get TableRow(ByVal lRow As Long) As Variant
    CheckTextChange
    TableRow = mvTable(lRow)
End Property

Public Property Get Text() As String
    CheckTableChange
    Text = msText
End Property

Public Property Get LastRow() As Long
    CheckTextChange
    LastRow = UBound(mvTable)
End Property

Public Property Get LastColumn() As Long
    Dim lRow As Long
    Dim lWidth As Long
    
    For lRow = 0 To UBound(Table)
        If UBound(mvTable(lRow)) > lWidth Then
            lWidth = UBound(mvTable(lRow))
        End If
    Next
    LastColumn = lWidth
End Property

Private Sub StampTextChange()
    mlTextStamp = mlTableStamp + 1
End Sub

Private Sub StampTableChange()
    mlTableStamp = mlTextStamp + 1
End Sub

Private Sub CheckTextChange()
    If mlTextStamp > mlTableStamp Then
        ConvertTextToTable
    End If
End Sub

Private Sub CheckTableChange()
    If mlTableStamp > mlTextStamp Then
        ConvertTableToText
    End If
End Sub


Public Sub ConvertTextToTable()
    Dim vSelectedCellsRow As Variant
    Dim lRowIndex As Long
    Dim vTable As Variant
    
    vTable = Split(msText, msRowDelimiter)
    mvTable = Array()
    If UBound(vTable) > -1 Then
        ReDim mvTable(UBound(vTable))
    Else
        ReDim mvTable(0)
    End If

    If UBound(vTable) > -1 Then
        For lRowIndex = 0 To UBound(vTable)
            If vTable(lRowIndex) <> "" Then
                mvTable(lRowIndex) = Split(vTable(lRowIndex), msColumnDelimiter)
            Else
                mvTable(lRowIndex) = Array("")
            End If
        Next
    Else
        mvTable(0) = Array("")
    End If
    
    mlTableStamp = mlTextStamp
End Sub

Public Sub ConvertTableToText()
    Dim vRow As Variant
    Dim vRows As Variant
    
    vRows = Array()
    For Each vRow In mvTable
        ReDim Preserve vRows(UBound(vRows) + 1)
        vRows(UBound(vRows)) = Join(vRow, msColumnDelimiter)
    Next
    msText = Join(vRows, msRowDelimiter)
    
    mlTextStamp = mlTableStamp
End Sub

Public Sub InsertText(sText, oCellPosition As CellPosition)
    Text = Left$(msText, oCellPosition.TextPosition) & sText & Mid$(msText, oCellPosition.TextPosition + 1)
    oCellPosition.TextPosition = oCellPosition.TextPosition + Len(sText)
End Sub

Public Sub DeleteText(lTextPosition As Long, ByVal lLength As Long)
    Text = Left$(Text, lTextPosition) & Mid$(Text, lTextPosition + lLength + 1)
End Sub

Public Function ArrayFromTableColumn(ByVal lColumn As Long) As Variant
    Dim lIndex As Long
    Dim vResult As Variant
    Dim vRow As Variant
    
    vResult = Array()
    For lIndex = 0 To UBound(Table)
        If UBound(TableRow(lIndex)) >= lColumn Then
            ReDim Preserve vResult(lIndex)
            vResult(lIndex) = TableCell(lIndex, lColumn)
        Else
            ArrayFromTableColumn = vResult
            Exit Function
        End If
    Next
    ArrayFromTableColumn = vResult
End Function

Public Function ArrayFromTableRow(ByVal lRow As Long) As Variant
    Dim lIndex As Long
    Dim vResult As Variant
    
    vResult = Array()
    For lIndex = 0 To UBound(TableRow(lRow))
        ReDim Preserve vResult(lIndex)
        vResult(lIndex) = TableCell(lRow, lIndex)
    Next
    ArrayFromTableRow = vResult
End Function

Public Sub ArrayIntoTableColumn(ByVal lColumn As Long, ByVal vColumn As Variant)
    Dim lIndex As Long
    Dim vRow As Variant
    Dim lFillIndex As Long
    Dim lOriginalRowSize As Long
    Dim lOriginalSize As Long
    
    lOriginalSize = UBound(mvTable)
    
    If lOriginalSize < UBound(vColumn) Then
        ReDim Preserve mvTable(UBound(vColumn))
    End If
    For lIndex = lOriginalSize + 1 To UBound(vColumn)
        mvTable(lIndex) = Array()
    Next
    For lIndex = 0 To UBound(vColumn)
        lOriginalRowSize = UBound(mvTable(lIndex))
        If lOriginalRowSize < lColumn Then
            vRow = mvTable(lIndex)
            ReDim Preserve vRow(lColumn)
            For lFillIndex = lOriginalRowSize + 1 To lColumn - 1
                vRow(lFillIndex) = ""
            Next
            vRow(lColumn) = vColumn(lIndex)
            mvTable(lIndex) = vRow
        Else
            mvTable(lIndex)(lColumn) = vColumn(lIndex)
        End If
    Next
    mlTableStamp = mlTextStamp + 1
End Sub

Public Sub ArrayIntoTableRow(ByVal lRow As Long, ByVal vRow As Variant)
    TableRow(lRow) = vRow
    mlTableStamp = mlTextStamp + 1
End Sub

Public Sub InsertBlankColumn(ByVal lColumn As Long)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(mvTable)
        If UBound(mvTable(lIndex)) >= lColumn Then
            TableRow(lIndex) = Concat(Splice(TableRow(lIndex), 0, lColumn - 1), Array(""), Splice(TableRow(lIndex), lColumn))
        End If
    Next
End Sub

Public Sub DeleteColumn(ByVal lColumn As Long)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(mvTable)
        If UBound(mvTable(lIndex)) >= lColumn Then
            TableRow(lIndex) = Concat(Splice(TableRow(lIndex), 0, lColumn - 1), Splice(TableRow(lIndex), lColumn + 1))
        End If
    Next
End Sub

Public Sub ClearCell(lRow As Long, lColumn As Long)
    TableCell(lRow, lColumn) = ""
End Sub

Public Sub TextReplace(ByVal sFind As String, ByVal sReplace As String)
    sFind = moParsing.Decode(sFind)
    sReplace = moParsing.Decode(sReplace)
    
    Text = Replace$(Text, sFind, sReplace)
End Sub

Public Function DecodeString(ByVal sString As String) As String
    DecodeString = moParsing.Decode(sString)
End Function

Public Sub SortByColumn(ByVal lColumn As Long, Optional ByVal bDescending As Boolean, Optional ByVal bIntelligentCompare As Boolean)
    Dim vTemp As Variant
    Dim bSorted As Boolean
    Dim lRow As Long
    
    While Not bSorted
        bSorted = True
        For lRow = 0 To UBound(mvTable) - 1
            If Not bDescending Then
                If Compare(TableCell(lRow, lColumn), TableCell(lRow + 1, lColumn), bIntelligentCompare) Then
                    vTemp = TableRow(lRow)
                    TableRow(lRow) = TableRow(lRow + 1)
                    TableRow(lRow + 1) = vTemp
                    bSorted = False
                End If
            Else
                If Compare(TableCell(lRow + 1, lColumn), TableCell(lRow, lColumn), bIntelligentCompare) Then
                    vTemp = TableRow(lRow)
                    TableRow(lRow) = TableRow(lRow + 1)
                    TableRow(lRow + 1) = vTemp
                    bSorted = False
                End If
            End If
        Next
    Wend
End Sub


Private Function Compare(ByVal sString1 As String, ByVal sString2 As String, ByVal bIntelligentCompare As Boolean) As Boolean
    Dim oString1 As New SaffronTree
    Dim oString2 As New SaffronTree
    Dim bSame As Boolean
    Dim lIndex As Long
    Dim sText1 As String
    Dim sText2 As String
    
    If Not bIntelligentCompare Then
        If sString1 > sString2 Then
            Compare = True
        End If
        Exit Function
    End If
    
    SaffronStream.Text = sString1
    moParsing.moSortLex.Parse oString1
    
    SaffronStream.Text = sString2
    moParsing.moSortLex.Parse oString2
     
    bSame = True
    lIndex = 1
    
    Do
        If lIndex > oString1.SubTree.Count And lIndex <= oString2.SubTree.Count Then
            Compare = False
            Exit Function
        ElseIf lIndex > oString2.SubTree.Count And lIndex <= oString1.SubTree.Count Then
            Compare = True
            Exit Function
        ElseIf lIndex > oString1.SubTree.Count And lIndex > oString2.SubTree.Count Then
            Compare = False
            Exit Function
        End If
        
        SaffronStream.Text = sString1
        sText1 = oString1.SubTree(lIndex).Text
        SaffronStream.Text = sString2
        sText2 = oString2.SubTree(lIndex).Text
        
        If oString1.SubTree(lIndex).Index = 1 And oString2.SubTree(lIndex).Index = 1 Then
            If Val(sText1) > Val(sText2) Then
                Compare = True
                Exit Function
            ElseIf Val(sText1) < Val(sText2) Then
                Compare = False
                Exit Function
            End If
        Else
            If sText1 > sText2 Then
                Compare = True
                Exit Function
            ElseIf sText1 < sText2 Then
                Compare = False
                Exit Function
            End If
        End If
        lIndex = lIndex + 1
    Loop
End Function

Private Sub Class_Initialize()
    mvTable = Array(Array())
End Sub

