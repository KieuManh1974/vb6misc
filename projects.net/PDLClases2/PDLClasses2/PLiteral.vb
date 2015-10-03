<ComClass(PLiteral.ClassId, PLiteral.InterfaceId, PLiteral.EventsId)> _
Public Class PLiteral
    Implements IParseObject

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "241ECFCC-58D4-4F9E-9242-A529A81ED69F"
    Public Const InterfaceId As String = "4688D029-E3DA-41F7-9731-7375AACEF6D6"
    Public Const EventsId As String = "FA6FE1B6-6610-443B-8A12-CDD645F37450"
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
    Public Function Initialise(ByVal pcParseControl As ParseControl, ByVal sErrorMessage As String, ByVal ParamArray pInitParam() As Object) As IParseObject Implements IParseObject.Initialise
        myResultControl = pcParseControl
        myErrorString = sErrorMessage
        myLiteralString = pInitParam(0)
        myLiteralStringLength = Len(myLiteralString)
        myOriginalLiteralString = myLiteralString
        If UBound(pInitParam) = 1 Then
            myCaseInsensitive = pInitParam(UBound(pInitParam))
            myLiteralString = UCase(myLiteralString)
        End If
        Initialise = Me
    End Function

    ' Will perform the parsing function on the object - if parsing fails will return FALSE.
    Public Function Parse(ByVal omyResult As ParseTree, Optional ByVal myIndex As Integer = 0) As Boolean Implements IParseObject.Parse
        Dim myPosition As Long
        Dim myGetChar As String
        Dim myOrigGetChar As String

        myPosition = TextPosition
        If (TextPosition + myLiteralStringLength - 1) > LenTextString Then
            omyResult.TextStart = myPosition
            omyResult.ErrorMessage = myErrorString
            Exit Function
        End If

        myOrigGetChar = Mid$(TextString, TextPosition, myLiteralStringLength)

        TextPosition = TextPosition + myLiteralStringLength

        myGetChar = myOrigGetChar
        If myCaseInsensitive Then
            myGetChar = UCase$(myGetChar)
        End If
        If myGetChar = myLiteralString Then
            omyResult.Control = myResultControl
            If myResultControl <> ParseControl.pcOmit Then
                omyResult.Text = myOrigGetChar
                omyResult.TextStart = myPosition
                omyResult.TextEnd = TextPosition - 1
            End If

            Parse = True
        Else
            TextPosition = myPosition
            omyResult.TextStart = myPosition
            omyResult.ErrorMessage = myErrorString
        End If
    End Function

    Public Function Start(ByVal x As Int32, ByVal ParamArray pInitParam() As Object) As Boolean Implements IParseObject.Start

    End Function
End Class


