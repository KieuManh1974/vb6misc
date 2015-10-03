Public Interface IParseObject
    ' Initialises parameters used for parsing
    Function Initialise(ByVal pcParseControl As ParseControl, ByVal sErrorMessage As String, ByVal ParamArray pInitParam() As Object) As IParseObject

    ' Will perform the parsing function on the object - if parsing fails will return FALSE.
    Function Parse(ByVal oResultString As ParseTree, Optional ByVal iIndex As Long = 0) As Boolean
End Interface
