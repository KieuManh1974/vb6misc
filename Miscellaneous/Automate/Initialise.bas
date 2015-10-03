Attribute VB_Name = "Initialise"
Option Explicit

Public Sub Main()
    Dim oTS As TextStream
    Dim oTree As New ParseTree
    Dim sPath As String
    
    Definition.Initialise
    
    If Command$ <> "" Then
'        MsgBox Command$
        sPath = Command$
        If Left$(sPath, 1) = """" Then
            sPath = Mid$(sPath, 2, Len(sPath) - 2)
        End If
        With New FileSystemObject
            Set oTS = .OpenTextFile(sPath, ForReading)
            Stream.Text = oTS.ReadAll
            If oParser.Parse(oTree) Then
                ExecuteScript oTree
            End If
        End With
    End If
End Sub
