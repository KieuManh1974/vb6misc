VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PDLObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ErrorString As String

Private oParserObjects As Collection

Public Property Get ParserObjects() As Collection
    Set ParserObjects = oParserObjects
End Property

Public Function SetNewDefinition(sLanguageDefinitionString As String) As Boolean
    Dim oLanguageLexicalAnalyser As New CLanguageLexAn
    Dim oLangauageCompiler As New CCompileLanguage
    Dim oParsedLanguage As New ParseTree

    ErrorString = ""
    
    Set oParserObjects = Nothing
    Set oParserObjects = New Collection
    
    ' Load in the lexical analysis description
    Stream.Text = sLanguageDefinitionString
    oLanguageLexicalAnalyser.CreateStructure
    If oLanguageLexicalAnalyser.Parse(oParsedLanguage) Then
        Set oParserObjects = oLangauageCompiler.CompileLanguage(oParsedLanguage)
        SetNewDefinition = True
    Else
        ErrorString = "Compile error at position " & oParsedLanguage.TextStart & vbCrLf
        ErrorString = ErrorString & Left(sLanguageDefinitionString, oParsedLanguage.TextStart)
    End If
End Function

