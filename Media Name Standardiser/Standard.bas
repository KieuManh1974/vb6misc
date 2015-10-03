Attribute VB_Name = "Standard"
Option Explicit

Public oParser As IParseObject

Sub Main()
    InitialiseParser
    'RenameFiles
    Debug.Print TidyName("Doctor Who 055 - Terror of the Autons 1.4.avi")
End Sub

Public Sub RenameFiles()
    Dim oFile As File
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    For Each oFile In oFSO.GetFolder(App.Path).Files
        'oFile.Name = TidyName(oFile.Name)
        oFSO.CreateTextFile App.Path & "\" & TidyName(oFile.Name) & ".txt", True
    Next
End Sub

Public Sub InitialiseParser()
    Dim definition As String
    Dim oFSO As New FileSystemObject

    definition = oFSO.OpenTextFile("D:\Programming Projects\Projects\Miscellaneous\Media Name Standardiser\Standard.pdl").ReadAll

    If Not SetNewDefinition(definition) Then
        Debug.Print ErrorString
        End
    End If

    Set oParser = ParserObjects("description")
End Sub

Private Function TidyName(sName As String) As String
    Dim oTree As New ParseTree
    Dim oPart As ParseTree
    Dim iNameCount As Integer
    Dim sExtension As String
    Dim sSeason As String
    Dim sEpisode As String
    Dim sSerialNumber As String
    Dim sTotalEpisodes As String
    Dim sLevel() As String
    Dim iLevelCount As Integer
    Dim oWord As ParseTree
    
    ParserTextString.ParserText = sName
    If Not oParser.Parse(oTree) Then
        TidyName = sName
        Exit Function
    End If
    
    For Each oPart In oTree.SubTree
        Select Case oPart.Index
            Case 1 'index
                Select Case oPart(1)(1).Index
                    Case 1 ' xxx of xxx
                        sEpisode = oPart(1)(1)(1)(1).Text
                        sTotalEpisodes = oPart(1)(1)(1)(2).Text
                    Case 2 ' part xxx
                        sEpisode = oPart(1)(1).Text
                    Case 3 ' episode xxx
                        sEpisode = oPart(1)(1).Text
                    Case 4 ' xxx x xxx
                        If iLevelCount < 2 Then
                            sSeason = oPart(1)(1)(1)(1).Text
                            sEpisode = oPart(1)(1)(1)(2).Text
                        Else
                            sEpisode = oPart(1)(1)(1)(1).Text
                            sTotalEpisodes = oPart(1)(1)(1)(2).Text
                        End If
                    Case 5 ' sxxx exxx
                        sSeason = oPart(1)(1).Text
                        sEpisode = oPart(1)(1).Text
                    Case 6 ' sxxx
                        sSeason = oPart(1)(1).Text
                    Case 7 ' xxx
                        sSerialNumber = oPart(1)(1).Text
                End Select
            Case 2 'extension
                sExtension = LCase$(oPart.Text)
            Case 3 'name
                ReDim Preserve sLevel(iLevelCount) As String
                For Each oWord In oPart(1).SubTree
                    sLevel(iLevelCount) = sLevel(iLevelCount) & " " & Capitalise(oWord.Text)
                Next
                sLevel(iLevelCount) = Mid$(sLevel(iLevelCount), 2)
                iLevelCount = iLevelCount + 1
            Case 4 'space
        End Select
    Next
    
    Dim iNameIndex As Long
    
    For iNameIndex = 0 To iLevelCount - 1
        If sSerialNumber <> "" And iNameIndex = iLevelCount - 1 Then
            TidyName = TidyName & " " & sSerialNumber
        End If
        TidyName = TidyName & " - " & sLevel(iNameIndex)
    Next
    If sSeason <> "" Then
        sSeason = sSeason & "-"
    End If
    TidyName = Mid$(TidyName, 4) & " - "
    TidyName = TidyName & sSeason & sEpisode & "." & sTotalEpisodes & "." & sExtension
End Function

Private Function Capitalise(sWord As String)
    If sWord <> "BBC" Then
        Capitalise = UCase$(Left$(sWord, 1)) & LCase$(Mid$(sWord, 2))
    Else
        Capitalise = UCase$(sWord)
    End If
End Function
