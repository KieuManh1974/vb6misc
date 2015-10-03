Attribute VB_Name = "Convert"
Option Explicit

Public Sub ConvertExcel2iPod()
    Dim oExcel As Excel.Application
    Dim lColumn As Long
    Dim lRow As Long
    Dim oWS As Worksheet
    
    Dim sText As String
    Dim sSpacer As String
    
    Dim nSpacer As Single
    Dim nTextWidth As Single
    Dim nNextColumn As Single
    Dim nCurrentPosition As Single
    Dim lPadCount As Long
    Dim nNextPosition As Single
    Dim nColumnWidth As Single
    
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim oFile As File
    Dim lDot As Long
    
    sSpacer = FormMeasure.txtSC
    nSpacer = FormMeasure.picText.TextWidth(sSpacer)
    nColumnWidth = nSpacer * Val(FormMeasure.txtCCW)
    
    For Each oFile In oFSO.GetFolder(App.Path).Files
        
        lDot = InStrRev(oFile.Name, ".")
        If UCase$(Mid$(oFile.Name, lDot + 1)) = "XLS" Then
            Set oTS = oFSO.OpenTextFile(App.Path & "\" & Left$(oFile.Name, lDot - 1) & ".txt", ForWriting, True)

            Set oExcel = New Excel.Application
            
            oExcel.Workbooks.Open App.Path & "\" & Left$(oFile.Name, lDot - 1) & ".xls"
            Set oWS = oExcel.Workbooks(1).Worksheets(1)

            For lRow = 1 To 1000
                nCurrentPosition = 0
                For lColumn = 1 To Val(FormMeasure.txtColumns)
                    sText = CStr(oWS.Cells(lRow, lColumn).Value)
                    If sText = "STOP" Then
                        GoTo EndNow
                    End If
                    
                    nTextWidth = FormMeasure.picText.TextWidth(sText)
                    nNextColumn = lColumn * nColumnWidth
                    lPadCount = Int((nNextColumn - nTextWidth - nCurrentPosition) / nSpacer)
                    
                    nNextPosition = nCurrentPosition + nTextWidth + lPadCount * nSpacer
                    
                    If nNextPosition <> nNextColumn Then
                        If nNextPosition < nNextColumn Then
                            If (nNextColumn - nNextPosition) > (nNextPosition + nSpacer - nNextColumn) Then
                                lPadCount = lPadCount + 1
                                nNextPosition = nNextPosition + nSpacer
                            End If
                        ElseIf nNextPosition > nNextColumn Then
                            If (nNextPosition - nNextColumn) > (nNextPosition - nNextColumn - nSpacer) Then
                                lPadCount = lPadCount + 1
                                nNextPosition = nNextPosition + nSpacer
                            End If
                        End If
                    End If
                
                    If lPadCount > 0 Then
                        oTS.Write sText & String(lPadCount, sSpacer)
                    Else
                        oTS.Write sText & sSpacer
                    End If
                    nCurrentPosition = nNextPosition
                Next
                oTS.WriteLine
            Next
            
EndNow:
            
            oTS.Close
            oExcel.Quit
            Set oExcel = Nothing
        End If
    Next
End Sub
