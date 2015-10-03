VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form1 
   Caption         =   "Get NS Issue"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtIssueNumber 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "current"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdGetIssue 
      Caption         =   "Get Issue"
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1800
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Issue Number"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TextTypes
    ttTitle
    ttHeading
    ttHeading2
    ttParagraph
End Enum

Private Sub cmdGetIssue_Click()
    Inet1.Execute "https://www.newscientist.com/user/login", "POST", "loginId=guille%2Ephillips%40googlemail%2Ecom&password=andalucia72&remember=true&source=form&redirectURL=", "Content-Type: application/x-www-form-urlencoded"
End Sub

Private Sub DeleteFiles()
    Dim oFSO As New FileSystemObject
    Dim oFile As File
    Dim lDot As Long
    Dim sExt As String
    
    For Each oFile In oFSO.GetFolder(App.path & "/article/oebps").Files
        lDot = InStrRev(oFile.Name, ".")
        sExt = LCase$(Mid$(oFile.Name, lDot + 1))
        Select Case sExt
            Case "html", "ncx", "opf"
                oFile.Delete True
        End Select
    Next
End Sub

Private Function ExtractIssue(ByVal sURL As String)
    Dim vFind As Variant
    Dim lPosition As Long
    Dim vFindA As Variant
    Dim vFindAClose As Variant
    Dim vFindAText As Variant
    Dim oScan As New clsScan
    
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim lIndex As Long
    Dim vArticleCaptions As Variant
    Dim vFindSection As Variant
    Dim vFindSectionEnd As Variant
    Dim sArticleCaption As String
    
    DeleteFiles
    
    vArticleCaptions = Array()
    
    oScan.msText = Inet1.OpenURL(sURL)
    vFind = oScan.FindNext(1, "<ul id=""issueArticles"" class=""markerlist"">")
    If vFind(0) Then
        lPosition = vFind(1)
        vFindA = oScan.FindNext(lPosition, "<h2><a href=""")
        Do While vFindA(0)
            vFindAClose = oScan.FindNext(vFindA(1), """>")
            vFindAText = oScan.FindNext(vFindAClose(1), "</a>")
            
            vFindSection = oScan.FindNext(vFindAText(1), "<p class=""lowlight"">")
            vFindSectionEnd = oScan.FindNext(vFindSection(1), "<span style=""padding")
            

            vFindSectionEnd(2) = Replace$(Replace$(vFindSectionEnd(2), "&amp;", "&"), "&amp;", "&")
            vFindSectionEnd(2) = Replace$(vFindSectionEnd(2), "&", "&amp;")
            Debug.Print vFindSectionEnd(2) & ": " & vFindAClose(2)
                        
            ArrayAdd vArticleCaptions, vFindSectionEnd(2) & ": " & vFindAText(2)
            
            lIndex = lIndex + 1
            Set oTS = oFSO.CreateTextFile(App.path & "/article/oebps/article" & lIndex & ".html", True, False)
            oTS.Write ArticleXHTML(lIndex, "Article " & lIndex, vFindSectionEnd(2), ExtractArticle(sURL & "/../.." & vFindAClose(2) & "?full=true"))
            oTS.Close

            lPosition = vFindA(1)
            vFindA = oScan.FindNext(lPosition, "<h2><a href=""")
        Loop
    End If
    
    CreateContentFile vArticleCaptions
    CreateTOCHTMLFile vArticleCaptions
    CreateTOCNCXFile vArticleCaptions
    CreateZipFile
End Function

Private Function ExtractArticle(ByVal sURL As String) As Variant
    Dim lStart As Long
    Dim vFind As Variant
    Dim lPosition As Long
    Dim vFindP As Variant
    Dim vFindCloseP As Variant
    Dim vLines As Variant
    Dim oScan As New clsScan
    
    vLines = Array()
    
    oScan.msText = Inet1.OpenURL(sURL)
    
    vFind = oScan.FindNext(1, "<h1>")
    vFind = oScan.FindNext(vFind(1), "</h1>")
    
    ArrayAddV vLines, Array(ttTitle, vFind(2))
    
    vFind = oScan.FindNext(1, "<div id=""maincol"" class=""floatleft"">")

    If vFind(0) Then
        lPosition = vFind(1)
        vFindP = oScan.FindNext(lPosition, "<p>", "<p class=""infuse"">", "<p class=""infotext"">", "<h3 class=""crosshead"">", "<h3 id=""")
        
        While vFindP(0) And vFindP(4) <> "<p class=""infotext"">"
            Select Case vFindP(4)
                Case "<h3 class=""crosshead"">"
                    vFindCloseP = oScan.FindNext(vFindP(1), "</h3>")
                    ArrayAddV vLines, Array(ttHeading, RemoveLinks(vFindCloseP(2)))
                    'Debug.Print vFindCloseP(2)
                    lPosition = vFindCloseP(1)
                Case "<h3 id="""
                    vFindCloseP = oScan.FindNext(vFindP(1), """>")
                    vFindCloseP = oScan.FindNext(vFindCloseP(1), "</h3>")
                    ArrayAddV vLines, Array(ttHeading2, RemoveLinks(vFindCloseP(2)))
                    'Debug.Print vFindCloseP(2)
                    lPosition = vFindCloseP(1)
                Case Else
                    vFindCloseP = oScan.FindNext(vFindP(1), "</p>")
                    ArrayAddV vLines, Array(ttParagraph, RemoveLinks(vFindCloseP(2)))
                    'Debug.Print vFindCloseP(2)
                    lPosition = vFindCloseP(1)
            End Select
            vFindP = oScan.FindNext(lPosition, "<p>", "<p class=""infuse"">", "<p class=""infotext"">", "<h3 class=""crosshead"">", "<h3 id=""")
        Wend
        
    End If
    ExtractArticle = vLines
End Function

Private Function RemoveLinks(ByVal sText As String) As String
    Dim oScanSub As clsScan
    Dim lPosition As Long
    Dim vStart As Variant
    Dim vEnd As Variant
    Dim vClose As Variant
    
    Set oScanSub = New clsScan
    oScanSub.msText = sText
    lPosition = 1
    
    vStart = oScanSub.FindNext(lPosition, "<a href=")
    lPosition = lPosition
    While vStart(0)
        RemoveLinks = RemoveLinks & vStart(2)
        vEnd = oScanSub.FindNext(vStart(1), ">")
        If vEnd(0) Then
            vClose = oScanSub.FindNext(vEnd(1), "</a>")
            RemoveLinks = RemoveLinks & vClose(2)
            lPosition = vClose(1)
        End If
        vStart = oScanSub.FindNext(lPosition, "<a href=")
        lPosition = lPosition
    Wend
    RemoveLinks = RemoveLinks & Mid$(sText, lPosition)
End Function

Private Sub ArrayAdd(vArray As Variant, ByVal sItem As String)
    ReDim Preserve vArray(UBound(vArray) + 1)
    vArray(UBound(vArray)) = sItem
End Sub

Private Sub ArrayAddV(vArray As Variant, ByVal vItem As Variant)
    ReDim Preserve vArray(UBound(vArray) + 1)
    vArray(UBound(vArray)) = vItem
End Sub

Private Function ArticleXHTML(ByVal lIndex As Long, sTitle As String, ByVal sSection As String, vLines) As String
    Dim sArticle As String
    Dim vLine As Variant
    
    sArticle = sArticle & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    sArticle = sArticle & "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd"">" & vbCrLf
    sArticle = sArticle & "<html xmlns=""http://www.w3.org/1999/xhtml"">" & vbCrLf
    sArticle = sArticle & "<head>" & vbCrLf
    sArticle = sArticle & "<meta http-equiv=""Content-Type"" content=""application/xhtml+xml; charset=UTF-8"" />"
    sArticle = sArticle & "<title>" & sTitle & "</title>" & vbCrLf
    sArticle = sArticle & "<link href=""stylesheet.css"" type=""text/css"" rel=""stylesheet"" />" & vbCrLf
    sArticle = sArticle & "<link rel=""stylesheet"" type=""application/vnd.adobe-page-template+xml"" href=""page-template.xpgt""/>" & vbCrLf
    sArticle = sArticle & "</head>" & vbCrLf
    sArticle = sArticle & "<body>" & vbCrLf
    
    sArticle = sArticle & "<h1>" & sSection & "</h1>" & vbCrLf
    
    For Each vLine In vLines
        vLine(1) = Replace$(vLine(1), "&", "&amp;")
        Select Case vLine(0)
            Case ttTitle
                sArticle = sArticle & "<h2>" & vLine(1) & "</h2>" & vbCrLf
            Case ttHeading
                sArticle = sArticle & "<h2>" & vLine(1) & "</h2>" & vbCrLf
            Case ttHeading2
                sArticle = sArticle & "<h2 class=""greybox"">" & vLine(1) & "</h2>" & vbCrLf
            Case ttParagraph
                sArticle = sArticle & "<p>" & vLine(1) & "</p>" & vbCrLf
        End Select
    Next
    
    sArticle = sArticle & "</body>" & vbCrLf
    sArticle = sArticle & "</html>" & vbCrLf
    
    ArticleXHTML = sArticle
End Function

Private Sub CreateTOCHTMLFile(vLines As Variant)
    Dim sFile As String
    Dim vLine As Variant
    Dim lIndex As Long
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    
    sFile = sFile & "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf
    sFile = sFile & "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.1//EN"" ""http://www.w3.org/TR/xhtml11/DTD/xhtml11.dtd"">" & vbCrLf
    sFile = sFile & "<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"">" & vbCrLf
    sFile = sFile & "<head>" & vbCrLf
    sFile = sFile & "<meta http-equiv=""Content-Type"" content=""application/xhtml+xml; charset=UTF-8"" />"
    sFile = sFile & "<title>Table of Contents</title>" & vbCrLf
    sFile = sFile & "<link href=""stylesheet.css"" type=""text/css"" rel=""stylesheet"" />" & vbCrLf
    sFile = sFile & "<link rel=""stylesheet"" type=""application/vnd.adobe-page-template+xml"" href=""page-template.xpgt""/>" & vbCrLf
    sFile = sFile & "</head>" & vbCrLf
    sFile = sFile & "<body>" & vbCrLf
    sFile = sFile & "<h3 align=""center"">CONTENTS</h3>" & vbCrLf
    
    lIndex = 1
    For Each vLine In vLines
        sFile = sFile & "<p><a href=""article" & lIndex & ".html"">" & vLine & "</a></p>" & vbCrLf
        lIndex = lIndex + 1
    Next
    
    sFile = sFile & "</body>" & vbCrLf
    sFile = sFile & "</html>" & vbCrLf

    Set oTS = oFSO.CreateTextFile(App.path & "/article/oebps/toc.html", True, False)
    oTS.Write sFile
    oTS.Close
End Sub

Private Sub CreateTOCNCXFile(vLines As Variant)
    Dim sFile As String
    Dim vLine As Variant
    Dim lIndex As Long
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    
    sFile = sFile & "<ncx xmlns=""http://www.daisy.org/z3986/2005/ncx/"" version=""2005-1"">" & vbCrLf
    sFile = sFile & "<head>" & vbCrLf
    sFile = sFile & "    <meta name=""dtb:uid"" content=""jedisaber.com082120071415""/>" & vbCrLf
    sFile = sFile & "    <meta name=""dtb:depth"" content=""1""/>" & vbCrLf
    sFile = sFile & "    <meta name=""dtb:totalPageCount"" content=""0""/>" & vbCrLf
    sFile = sFile & "    <meta name=""dtb:maxPageNumber"" content=""0""/>" & vbCrLf
    sFile = sFile & "</head>" & vbCrLf
    sFile = sFile & "<docTitle>" & vbCrLf
    sFile = sFile & "    <text>New Scientist</text>" & vbCrLf
    sFile = sFile & "</docTitle>" & vbCrLf
    sFile = sFile & "<navMap>" & vbCrLf
  
    sFile = sFile & "<navPoint id=""navpoint-1"" playOrder=""1"">" & vbCrLf
    sFile = sFile & "<navLabel>" & vbCrLf
    sFile = sFile & "<text>Table of Contents</text>" & vbCrLf
    sFile = sFile & "</navLabel>" & vbCrLf
    sFile = sFile & "<content src=""toc.html""/>" & vbCrLf
    sFile = sFile & "</navPoint>" & vbCrLf
    
    lIndex = 1
    For Each vLine In vLines
        sFile = sFile & "<navPoint id=""navpoint-" & lIndex + 1 & """ playOrder=""" & lIndex + 1 & """>" & vbCrLf
        sFile = sFile & "<navLabel>" & vbCrLf
        sFile = sFile & "<text>" & ReplaceHTMLEntities(vLine) & "</text>" & vbCrLf
        sFile = sFile & "</navLabel>" & vbCrLf
        sFile = sFile & "<content src=""article" & lIndex & ".html""/>" & vbCrLf
        sFile = sFile & "</navPoint>" & vbCrLf
        lIndex = lIndex + 1
    Next
    sFile = sFile & "  </navMap>" & vbCrLf
    sFile = sFile & "</ncx>" & vbCrLf

    Set oTS = oFSO.CreateTextFile(App.path & "/article/oebps/toc.ncx", True, False)
    oTS.Write sFile
    oTS.Close
End Sub

Private Function ReplaceHTMLEntities(ByVal sString As String) As String
    ReplaceHTMLEntities = Replace(sString, "&ndash;", "&#x2013;")
End Function

Private Sub CreateContentFile(vLines As Variant)
    Dim sFile As String
    Dim vLine As Variant
    Dim sSpine As String
    Dim sManifest As String
    Dim lIndex As Long
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    
    sFile = sFile & "<?xml version=""1.0""?>" & vbCrLf
    sFile = sFile & "<!DOCTYPE package PUBLIC ""+//ISBN 0-9673008-1-9//DTD OEB 1.2 Package//EN""  ""http://openebook.org/dtds/oeb-1.2/oebpkg12.dtd"">" & vbCrLf
    sFile = sFile & "<package xmlns=""http://www.idpf.org/2007/opf"" unique-identifier=""bookid"" version=""2.0"">" & vbCrLf
    sFile = sFile & "<metadata xmlns:dc=""http://purl.org/dc/elements/1.1/"">" & vbCrLf
    sFile = sFile & "    <dc:title>New Scientist " & Format$(Now, "YYYYMMDD") & "</dc:title>" & vbCrLf
    sFile = sFile & "    <dc:language>en</dc:language>" & vbCrLf
    sFile = sFile & "    <dc:identifier id=""bookid""/>" & vbCrLf
    sFile = sFile & "    <dc:creator>Guillermo Phillips</dc:creator>" & vbCrLf
    sFile = sFile & "</metadata>" & vbCrLf
    
    sManifest = sManifest & "<manifest>" & vbCrLf
    sManifest = sManifest & "<item id=""ncx"" href=""toc.ncx"" media-type=""text/xml""/>" & vbCrLf
    sManifest = sManifest & "<item id=""style"" href=""stylesheet.css"" media-type=""text/css""/>" & vbCrLf
    sManifest = sManifest & "<item id=""pagetemplate"" href=""page-template.xpgt"" media-type=""application/vnd.adobe-page-template+xml""/>" & vbCrLf
    sManifest = sManifest & "<item id=""tableofc"" href=""toc.html"" media-type=""text/html""/>" & vbCrLf
    
    sSpine = sSpine & "<spine toc=""ncx"">" & vbCrLf
    sSpine = sSpine & "<itemref idref=""tableofc""/>" & vbCrLf

    lIndex = 1
    For Each vLine In vLines
        sManifest = sManifest & "<item id=""article" & lIndex & """ href=""article" & lIndex & ".html"" media-type=""text/html""/>" & vbCrLf
        sSpine = sSpine & "<itemref idref=""article" & lIndex & """/>" & vbCrLf
        lIndex = lIndex + 1
    Next
    
    sManifest = sManifest & "</manifest>" & vbCrLf
        
    sSpine = sSpine & "</spine>" & vbCrLf
    
    sFile = sFile & sManifest
    sFile = sFile & sSpine
    sFile = sFile & "</package>" & vbCrLf
    
    Set oTS = oFSO.CreateTextFile(App.path & "/article/oebps/content.opf", True, False)
    oTS.Write sFile
    oTS.Close
End Sub



Public Sub CreateZipFile()
    Dim oZip As New clsCZip

    With oZip
      .ZipFile = App.path & "\article\NewScientist.epub"
      .Encrypt = False
      .AddComment = False
      .BasePath = App.path & "\article"
      .ClearFileSpecs
      .AddFileSpec "mimetype"
      .StoreFolderNames = True
      .RecurseSubDirs = False
      .Zip
      
      .ClearFileSpecs
      .AddFileSpec "*.html"
      .AddFileSpec "*.opf"
      .AddFileSpec "*.xml"
      .AddFileSpec "*.xpgt"
      .AddFileSpec "*.css"
      .AddFileSpec "*.ncx"
      .RecurseSubDirs = True
      .Update = True
      .Zip
    End With
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
    Dim sResponse As String
    Dim vChunk As Variant
    
    Select Case State
        Case icError
            MsgBox "could not log in"
        Case icResponseCompleted
            'Debug.Print Inet1.GetHeader()
            'Debug.Print Inet1.GetChunk(10000)
            ExtractIssue "http://www.newscientist.com/issue/" & txtIssueNumber.Text
            MsgBox "Done", vbOKOnly, "Get NS Issue"
    End Select
End Sub
