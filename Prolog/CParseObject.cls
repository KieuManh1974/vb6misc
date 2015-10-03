VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CParseObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'
'' /// PRIVATE VARIABLES
'
'
'' // Result ignore
'Public Ignore As Boolean
'
'' // For character choice
'Public CharChoice As String
'
'' // For character range
'Public StartRange As String ' * 1
'Public EndRange As String ' * 1
'
'' // For literal
'Public LiteralString As String
'
'' // For join, choice, repeat until or repeat collection
'Public ParseCollection As New Collection
'
'' // For join, choice, repeat or repeat until collection
'Public Complex As Boolean
'
'' // For choice collection
'Public CanIgnore As Boolean
'
'' // For repeat collection
'Public MinimumRepeat As Long
'Public MaximumRepeat As Long
'
'' // For repeat until collection
'Public TerminatorObject As Variant
'
'' // Error if fails
'Public ErrorString As String
'
'
'Private Enum ObjectEnum
'    CharacterChoiceObject
'    CharacterRangeObject
'    LiteralObject
'    JoinObject
'    ChoiceObject
'    RepeatObject
'    RepeatUntilObject
'    NotObject
'End Enum
'
'Public ObjectType As Long
'
'
'' /// OBJECT INITIALISE FUNCTIONS
'
'' // Initialise object as Not
'Public Sub Neg(vElement As Variant)
'    ParseCollection.Add vElement
'End Sub
'
'' // Initialise object as Join
'Public Sub Join(bIgnore As Boolean, bComplex As Boolean, ParamArray vElements() As Variant)
'    Dim iIndex As Long
'    For iIndex = LBound(vElements) To UBound(vElements)
'        ParseCollection.Add vElements(iIndex)
'    Next
'
'    Ignore = bIgnore
'    Complex = bComplex
'    ObjectType = JoinObject
'End Sub
'
'' // Initialise object as Choice
'Public Sub Choice(bIgnore As Boolean, bComplex As Boolean, bCanIgnore As Boolean, ParamArray vElements() As Variant)
'    Dim iIndex As Long
'    For iIndex = LBound(vElements) To UBound(vElements)
'        ParseCollection.Add vElements(iIndex)
'    Next
'    Ignore = bIgnore
'    Complex = bComplex
'    CanIgnore = bCanIgnore
'    ObjectType = ChoiceObject
'End Sub
'
'' // Initialise object as Repeat
'Public Sub Repeat(bIgnore As Boolean, bComplex As Boolean, MinR As Long, MaxR As Long, ParamArray vElements() As Variant)
'    MinimumRepeat = MinR
'    MaximumRepeat = MaxR
'
'    Dim iIndex As Long
'    For iIndex = LBound(vElements) To UBound(vElements)
'        ParseCollection.Add vElements(iIndex)
'    Next
'
'    Complex = bComplex
'    Ignore = bIgnore
'    ObjectType = RepeatObject
'End Sub
'
''// Intialise object as RepeatUntil
'Public Sub RepeatUntil(bIgnore As Boolean, bComplex As Boolean, vTermCondition As Variant, ParamArray vElements() As Variant)
'
'    Dim iIndex As Long
'    For iIndex = LBound(vElements) To UBound(vElements)
'        ParseCollection.Add vElements(iIndex)
'    Next
'
'    Complex = bComplex
'    Ignore = bIgnore
'    Set TerminatorObject = vTermCondition
'    ObjectType = RepeatUntilObject
'End Sub
'
'' // Intialise object as CharacterRange
'Public Sub CharacterRange(bIgnore As Boolean, StartOfRange As String, EndOfRange As String)
'    StartRange = StartOfRange
'    EndRange = EndOfRange
'    Ignore = bIgnore
'    ObjectType = CharacterRangeObject
'End Sub
'
'
'' // Initialise object as Literal
'Public Sub Literal(bIgnore As Boolean, sLiteral As String)
'    Ignore = bIgnore
'    LiteralString = sLiteral
'    ObjectType = LiteralObject
'End Sub
'
'' // Initialise object as CharacterChoice
'Public Sub CharacterChoice(bIgnore As Boolean, sChars As String)
'    CharChoice = sChars
'    Ignore = bIgnore
'    ObjectType = CharacterChoiceObject
'End Sub
'
