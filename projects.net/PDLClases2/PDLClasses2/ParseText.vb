<ComClass(ParserText.ClassId, ParserText.InterfaceId, ParserText.EventsId)> _
Public Class ParserText

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "1F5D0322-8221-4CCD-8284-5D4C4179942F"
    Public Const InterfaceId As String = "BCC0D07C-A332-414B-85FC-A56217ADB42F"
    Public Const EventsId As String = "C96ABC9E-9F66-49CC-94EE-E6D8CD1DDCE9"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub


    Public Property ParserText() As String
        Get
            Return TextString
        End Get
        Set(ByVal sTextString As String)
            TextString = sTextString
            TextPosition = 1
            LenTextString = Len(sTextString)
        End Set
    End Property

    Public Sub ResetText()
        TextPosition = 1
    End Sub

    Public Property ParserTextPosition() As Long
        Get
            Return TextPosition
        End Get
        Set(ByVal lPosition As Long)
            TextPosition = lPosition
        End Set
    End Property
End Class


