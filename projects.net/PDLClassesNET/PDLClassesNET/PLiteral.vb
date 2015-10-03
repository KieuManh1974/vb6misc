<ComClass(ComClass2.ClassId, ComClass2.InterfaceId, ComClass2.EventsId)> _
Public Class ComClass2
    'Implements IParseObject

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "02D5455C-30EC-4A97-8E3D-0AC84FEA7F65"
    Public Const InterfaceId As String = "2DA3FDD0-36B2-405D-9CF2-CFD7D83192DB"
    Public Const EventsId As String = "5887F258-2E79-4F5F-A869-700D3D3AC973"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    ' Generic variables
    Private myErrorString As String
    Private myResultControl As ParseControl

    ' Object specific variables
    Private myLiteralString As String
    Private myOriginalLiteralString As String
    Private myLiteralStringLength As Long
    Private myCaseInsensitive As Boolean

    ' Initialises parameters used for parsing
    Public Function Initialise(ByVal pcParseControl As ParseControl, ByVal sErrorMessage As String, ByVal ParamArray pInitParam() As Object) As IParseObject 'Implements IParseObject.Initialise
        'myResultControl = pcParseControl
        'myErrorString = sErrorMessage
        'myLiteralString = pInitParam(0)
        'myLiteralStringLength = Len(myLiteralString)
        'myOriginalLiteralString = myLiteralString
        'If UBound(pInitParam) = 1 Then
        '    myCaseInsensitive = pInitParam(UBound(pInitParam))
        '    myLiteralString = UCase(myLiteralString)
        'End If
        'Initialise = Me
    End Function

    ' Will perform the parsing function on the object - if parsing fails will return FALSE.
    Public Function Parse(ByVal omyResult As ParseTree, Optional ByVal myIndex As Long = 0) As Boolean 'Implements IParseObject.Parse
        'Dim myPosition As Long
        'Dim myGetChar As String
        'Dim myOrigGetChar As String

        'myPosition = TextPosition
        'If (TextPosition + myLiteralStringLength - 1) > LenTextString Then
        '    omyResult.TextStart = myPosition
        '    omyResult.ErrorMessage = myErrorString
        '    Exit Function
        'End If

        'myOrigGetChar = Mid$(TextString, TextPosition, myLiteralStringLength)

        'TextPosition = TextPosition + myLiteralStringLength

        'myGetChar = myOrigGetChar
        'If myCaseInsensitive Then
        '    myGetChar = UCase$(myGetChar)
        'End If
        'If myGetChar = myLiteralString Then
        '    omyResult.Control = myResultControl
        '    If myResultControl <> ParseControl.pcOmit Then
        '        omyResult.Text = myOrigGetChar
        '        omyResult.TextStart = myPosition
        '        omyResult.TextEnd = TextPosition - 1
        '    End If

        '    Parse = True
        'Else
        '    TextPosition = myPosition
        '    omyResult.TextStart = myPosition
        '    omyResult.ErrorMessage = myErrorString
        'End If
    End Function

End Class


