Attribute VB_Name = "Initialise"
Option Explicit

Public oParser As IParseObject
Public mdCloseTime As Date

Public Sub InitialiseParser()
    Dim definition As String
    Dim oFSO As New FileSystemObject
    
    definition = oFSO.OpenTextFile(App.Path & "\bidder.pdl").ReadAll

    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = ParserObjects("data")
End Sub

Public Function Analyse(sEbayText As String, oForm As frmAnalyse) As Collection
    Dim sCheck As String
    Dim lChar As Long
    Dim oTree As ParseTree
    Dim cAmount As Currency
    Dim dDate As Date
    Dim oDatum As clsDatum
    Dim dNow As Date
    Dim oDataItem As ParseTree
    
    dNow = Now()
    
    oForm.pbProgress.Max = Len(sEbayText)
    
    Stream.Text = sEbayText
    Set oTree = New ParseTree
    If Not oParser.Parse(oTree) Then
        Exit Function
    End If
    
    Set Analyse = New Collection
    
    For Each oDataItem In oTree.SubTree
        If oDataItem(1).Index = 1 Then
            Select Case oDataItem(1)(1).Index
                Case 1 ' amount
                    cAmount = Mid$(oDataItem.Text, 2)
                Case 2 ' date
                    dDate = CDate(oDataItem.Text)
                    If oTree.Index = 2 Then ' date
                        lChar = lChar + 3
                    End If
                    Set oDatum = New clsDatum
                    oDatum.BidDate = dDate
                    oDatum.BidAmount = cAmount
                    Analyse.Add oDatum
                Case 3 ' Time left
                    Select Case oTree(1).Index
                        Case 1 ' sec
                            mdCloseTime = dNow + oTree(1)(1)(1).Text / 86400
                        Case 2 ' min sec
                            mdCloseTime = dNow + oTree(1)(1)(1).Text / 1440 + oTree(1)(1)(2).Text / 86400
                        Case 3 ' hour min sec
                            mdCloseTime = dNow + oTree(1)(1)(1).Text / 24 + oTree(1)(1)(2).Text / 1440 + oTree(1)(1)(3).Text / 86400
                        Case 4 ' day hour min sec
                            mdCloseTime = dNow + oTree(1)(1)(1).Text + oTree(1)(1)(2).Text / 24 + oTree(1)(1)(3).Text / 1440 + oTree(1)(1)(4).Text / 86400
                    End Select
            End Select
        End If
    Next
    oForm.pbProgress.Value = 0
End Function
