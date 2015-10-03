VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3Tag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msPath As String * 30
Private msName As String * 30
Private msArtist As String * 30
Private msAlbum As String * 30
Private msYear As String * 4
Private msComment As String * 30
Private msGenre As String * 1

Private msFilename As String

Private mlID3TagStart As Long

Private mbValid As Boolean

Public Property Let Path(ByVal sPath As String)
    Dim sTag As String * 3
    Dim oFSO As New FileSystemObject
    
    msFilename = oFSO.GetFile(sPath).Name
    
    mbValid = False
    Open sPath For Binary As #1
    
    mlID3TagStart = FileLen(sPath) - 127
    If mlID3TagStart > 0 Then
        Seek #1, mlID3TagStart
        
        'Get #1, , msTest
        Get #1, , sTag
        If sTag = "TAG" Then
            mbValid = True
            
            Get #1, , msPath
            Get #1, , msName
            Get #1, , msArtist
            Get #1, , msYear
            Get #1, , msComment
            Get #1, , msGenre
        End If
    End If
    Close #1
End Property

Public Property Get Path() As String
    Path = msPath
End Property


Public Property Let Name(ByVal sPath As String)

End Property

Public Property Get Name() As String
    Name = msName
End Property


Public Property Let Artist(ByVal sPath As String)

End Property

Public Property Get Artist() As String
    Artist = msArtist
End Property


Public Property Let Album(ByVal sPath As String)

End Property

Public Property Get Album() As String
    Album = msAlbum
End Property


Public Property Let Year(ByVal sPath As String)

End Property

Public Property Get Year() As String
    Year = msYear
End Property


Public Property Let Comment(ByVal sPath As String)

End Property

Public Property Get Comment() As String
    Comment = msComment
End Property


Public Property Let Genre(ByVal sPath As String)

End Property

Public Property Get Genre() As String
    Genre = msGenre
End Property

Public Property Get IdealFilename() As String
    Dim sFilename As String
    
    If mbValid Then
        sFilename = UCase$(msFilename)
        
        If InStr(sFilename, UCase$(msName)) > 0 And InStr(sFilename, UCase$(msArtist)) > 0 Then
            IdealFilename = StrConv(msArtist, vbProperCase) & " - " & StrConv(msName, vbProperCase)
        Else
            IdealFilename = msFilename
        End If
    Else
        IdealFilename = msFilename
    End If
End Property
