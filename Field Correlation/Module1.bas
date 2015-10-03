Attribute VB_Name = "Module1"
Option Explicit

Private sFieldList() As String
Private vFieldValues As Variant

Private oConEBS As ADODB.Connection

Sub main()
    OpenConnections
    PerformCorrelation GetRecords
    CloseConnections
End Sub

Private Sub OpenConnections()
    Set oConEBS = New ADODB.Connection
    oConEBS.Open "Provider=msdaora.1;Data Source=students;User id=report;Password=reporting;"
End Sub

Private Sub CloseConnections()
    oConEBS.Close
    Set oConEBS = Nothing
End Sub

Private Function GetRecords() As ADODB.Recordset
    Dim sSQL As String
    
    sSQL = ""
    sSQL = sSQL & " SELECT"
    sSQL = sSQL & "     A.PROGRESS_CODE APC,"
    sSQL = sSQL & "     U.PROGRESS_CODE AUPC"
    sSQL = sSQL & " FROM"
    sSQL = sSQL & "     APPLICATIONS A,"
    sSQL = sSQL & "     APPLICATION_UNITS U"
    sSQL = sSQL & " WHERE"
    sSQL = sSQL & "     U.APP_APPLICATION_NUMBER = A.APPLICATION_NUMBER"
    sSQL = sSQL & "     AND U.CALENDAR_OCCURRENCE_CODE IN ('0506', '0607')"
    
    Set GetRecords = New ADODB.Recordset
    GetRecords.Open sSQL, oConEBS, adOpenForwardOnly, , adCmdText
End Function

Private Sub PerformCorrelation(oRecords As ADODB.Recordset)
    Dim iFieldIndex As Long
    Dim iFieldIndex2 As Long
    Dim iValueIndex1 As Long
    Dim iValueIndex2 As Long
    Dim bFound As Boolean
    Dim bFound2 As Boolean
    Dim vNew As Variant
    Dim vCell As Variant
    Dim vSubCell As Variant
    
    ReDim sFieldList(oRecords.Fields.Count) As String
    ReDim vFieldValues(oRecords.Fields.Count - 1, oRecords.Fields.Count - 1) As Variant
    
    For iFieldIndex = 0 To oRecords.Fields.Count - 1
        sFieldList(iFieldIndex) = oRecords.Fields(iFieldIndex).Name
        For iFieldIndex2 = 0 To oRecords.Fields.Count - 1
            vFieldValues(iFieldIndex, iFieldIndex2) = Array()
        Next
    Next
    
    With oRecords
        While Not .EOF
            
            For iFieldIndex = 0 To oRecords.Fields.Count - 1
                For iFieldIndex2 = 0 To oRecords.Fields.Count - 1
                    If iFieldIndex <> iFieldIndex2 Then
                        vCell = vFieldValues(iFieldIndex, iFieldIndex2)
                        bFound = False
                        For iValueIndex1 = 0 To UBound(vCell)
                            If vCell(iValueIndex1)(0) = .Fields(iFieldIndex).Value Or (IsNull(vCell(iValueIndex1)(0)) And IsNull(.Fields(iFieldIndex).Value)) Then
                                bFound = True
                                bFound2 = False
                                For iValueIndex2 = 0 To UBound(vCell(iValueIndex1)(1))
                                    If vCell(iValueIndex1)(1)(iValueIndex2)(0) = .Fields(iFieldIndex2).Value Or (IsNull(vCell(iValueIndex1)(1)(iValueIndex2)(0)) And IsNull(.Fields(iFieldIndex2).Value)) Then
                                        bFound2 = True
                                        vFieldValues(iFieldIndex, iFieldIndex2)(iValueIndex1)(1)(iValueIndex2)(1) = vCell(iValueIndex1)(1)(iValueIndex2)(1) + 1
                                        Exit For
                                    End If
                                Next
                                If Not bFound2 Then
                                    vSubCell = vCell(iValueIndex1)(1)
                                    ReDim Preserve vSubCell(UBound(vSubCell) + 1) As Variant
                                    vSubCell(UBound(vSubCell)) = Array(.Fields(iFieldIndex2).Value, 1)
                                    vCell(iValueIndex1)(1) = vSubCell
                                    vFieldValues(iFieldIndex, iFieldIndex2) = vCell
                                End If
                            End If
                        Next
                        If Not bFound Then
                            ReDim Preserve vCell(UBound(vCell) + 1) As Variant
                            vCell(UBound(vCell)) = Array(.Fields(iFieldIndex).Value, Array(Array(.Fields(iFieldIndex2).Value, 1)))
                            vFieldValues(iFieldIndex, iFieldIndex2) = vCell
                        End If
                    End If
                Next
            Next
    
            .MoveNext
        Wend
    End With
End Sub

