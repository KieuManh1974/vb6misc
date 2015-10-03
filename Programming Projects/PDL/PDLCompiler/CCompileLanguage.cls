VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompileLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Objs As New Collection
Private ErrorDescriptions As New Collection

Private UniqueIndex As Integer

Private Enum StatementTypes
    stLiteral = 1
    stAnd
    stPerm
    stOr
    stIn
    stNot
    stRepeatUntil
    stRepeat
    stOptional
    stList
    stSkip
    stEOS
    stBOS
    stPass
    stFail
    stExternal
    stEach
    stVariable
End Enum

'Private Const stExternal As Long = -1
   
Private Enum ExpressionType
    etIgnore = 1
    etOmit = 2
    etOverride = 3
    etBracket = 4
    etLeave = 5
End Enum

Friend Function CompileLanguage(oReturnProgram As ParseTree) As Collection
    Dim ObjectName As String
    Dim oObject As IParseObject
    Dim oStatement As ParseTree
    Dim lStatementIndex As Long
    
    ' Create variables
    For Each oStatement In oReturnProgram.SubTree
        ObjectName = oStatement(1)(1).Text
        Objs.Add CreateObjects(oStatement(1)(2)), ObjectName
    Next
    
    For Each oStatement In oReturnProgram.SubTree
        ObjectName = oStatement(1)(1).Text
        CompileFunction oStatement(1)(2), Objs(ObjectName), ObjectName
    Next
    
    Set CompileLanguage = Objs
End Function

Private Function CreateObjects(ByVal oFunctionIn As ParseTree) As IParseObject
    Dim StatementType As Integer
    Dim oFunction As ParseTree
    
    Set oFunction = oFunctionIn(1)(1)
    
    StatementType = CInt(oFunction.Index)
    Select Case StatementType
        Case stLiteral
            Set CreateObjects = Create("Literal")
            
        Case stAnd
            Set CreateObjects = Create("And")
            
        Case stPerm
            Set CreateObjects = Create("Perm")
                        
        Case stOr
            Set CreateObjects = Create("Or")
            
        Case stIn
            Set CreateObjects = Create("In")
            
        Case stNot
            Set CreateObjects = Create("Not")
            
        Case stRepeatUntil
            Set CreateObjects = Create("RepeatUntil")
        
        Case stRepeat
            Set CreateObjects = Create("Repeat")
        
        Case stOptional
            Set CreateObjects = Create("Optional")

        Case stList
            Set CreateObjects = Create("List")
            
        Case stSkip
            Set CreateObjects = Create("Skip")
                        
        Case stEOS
            Set CreateObjects = Create("EOS")
            
        Case stBOS
            Set CreateObjects = Create("BOS")
                        
        Case stPass
            Set CreateObjects = Create("Null")
                       
        Case stFail
            Set CreateObjects = Create("Null")
                              
        Case stExternal
            Set CreateObjects = Create("External")
        
        Case stEach
            Set CreateObjects = Create("Each")
        
        Case stVariable
            Set CreateObjects = Objs(oFunction.Text)
    End Select
End Function

Private Function ControlType(ByVal oFunctionIn As ParseTree) As ParseControl
    Select Case oFunctionIn.Index
        Case etIgnore
            ControlType = pcHide
        Case etOmit
            ControlType = pcOmit
        Case etOverride
            ControlType = pcOverride
        Case etBracket
            ControlType = pcLeave
        Case etLeave
            ControlType = pcLeave
    End Select
End Function

Private Function CompileFunction(ByVal oFunctionIn As ParseTree, Optional WorkingObject As IParseObject, Optional sObjectName As String) As IParseObject
    Dim StatementType As Integer
    Dim ObjectType As String
    Dim oObject As IParseObject
    Dim pcParseControl As ParseControl
    Dim iType As Integer
    Dim oObj As ParseTree
    Dim oObj1 As IParseObject
    Dim oParams As Collection
    Dim minimum As Integer
    Dim maximum As Integer
    Dim oFunction As ParseTree
    Dim lObjIndex As Long
    Dim bNullControl As Boolean
    Dim lSkipNo As Long
    
    Select Case oFunctionIn.Index
        Case etIgnore
            pcParseControl = pcHide
        Case etOmit
            pcParseControl = pcOmit
        Case etOverride
            pcParseControl = pcOverride
        Case etBracket
            pcParseControl = pcLeave
        Case etLeave
            pcParseControl = pcLeave
    End Select
    
    Set oFunction = oFunctionIn(1)(1)

    StatementType = CInt(oFunction.Index)
    Select Case StatementType
        Case stLiteral, stAnd, stPerm, stOr, stIn, stNot, stRepeatUntil, stRepeat, stOptional, stList, stSkip, stEOS, stBOS, stPass, stFail, stExternal, stEach
            If WorkingObject Is Nothing Then
                Set CompileFunction = Create(Choose(StatementType, "Literal", "And", "Perm", "Or", "In", "Not", "RepeatUntil", "Repeat", "Optional", "List", "Skip", "EOS", "BOS", "Null", "Null", "External"))
            Else
                Set CompileFunction = WorkingObject
            End If
            
        Case stVariable
            Set CompileFunction = WorkingObject
    End Select
    
    Select Case StatementType
        Case stLiteral
            Dim sLiteral As String
            Dim iCount As Integer
    
            If oFunction(1)(2).Index = 1 Then ' String literal
                sLiteral = sLiteral & ReduceDoubleSingleQuotes(oFunction(1)(2)(1)(2).Text)
            Else ' Numeric literal
                sLiteral = sLiteral & Chr(oFunction(1)(2).Text)
            End If
            
            For iCount = 1 To oFunction(1)(3).Index
                If oFunction(1)(3)(iCount)(2).Index = 1 Then ' String literal
                    sLiteral = sLiteral & ReduceDoubleSingleQuotes(oFunction(1)(3)(iCount)(2)(1)(2).Text)
                Else ' Numeric literal
                    sLiteral = sLiteral & Chr(oFunction(1)(3)(iCount)(2).Text)
                End If
            Next
            
            If oFunction(1)(1).Index = 0 Then ' No case keyword
                CompileFunction.Initialise sObjectName, pcParseControl, False, sLiteral
            Else ' with case keyword
                CompileFunction.Initialise sObjectName, pcParseControl, False, sLiteral, True
            End If

        Case stAnd
            Set oParams = New Collection
            For Each oObj In oFunction(1)(3).SubTree
                oParams.Add CompileFunction(oObj)
                oParams.Add ControlType(oObj)
            Next
            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, oParams
            
        Case stPerm
            Set oParams = New Collection
            For Each oObj In oFunction(1)(3).SubTree
                oParams.Add CompileFunction(oObj)
                oParams.Add ControlType(oObj)
            Next
            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, oParams
                        
        Case stOr
            Set oParams = New Collection
            For Each oObj In oFunction(1)(3).SubTree
                oParams.Add CompileFunction(oObj)
                oParams.Add ControlType(oObj)
            Next
            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, oParams
            
        Case stIn
            CompileFunction.Initialise sObjectName, pcParseControl, False, ResolveString(oFunction(1)(2))
                        
        Case stNot
            CompileFunction.Initialise sObjectName, pcParseControl, False, CompileFunction(oFunction(1)(2))
            
        Case stRepeatUntil
            minimum = 1
            maximum = 0
            
            If oFunction(1)(6).Index = 1 Then
                minimum = CInt(oFunction(1)(6)(1)(2).Text)
            End If
            
            If oFunction(1)(7).Index = 1 Then
                maximum = CInt(oFunction(1)(7)(1)(2).Text)
            End If

            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, CompileFunction(oFunction(1)(3)), ControlType(oFunction(1)(3)), CompileFunction(oFunction(1)(5)), ControlType(oFunction(1)(5)), maximum, minimum
        
        Case stRepeat
            minimum = 1
            maximum = 0
            
            If oFunction(1)(4).Index = 1 Then
                minimum = CInt(oFunction(1)(4)(1)(2).Text)
            End If
            
            If oFunction(1)(5).Index = 1 Then
                maximum = CInt(oFunction(1)(5)(1)(2).Text)
            End If
            
            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, CompileFunction(oFunction(1)(3)), ControlType(oFunction(1)(3)), maximum, minimum
            
        Case stOptional
            CompileFunction.Initialise sObjectName, pcParseControl, False, CompileFunction(oFunction(1)(2)), ControlType(oFunction(1)(2))
        
        Case stList
            minimum = 1
            maximum = 0
            
            If oFunction(1)(6).Index = 1 Then
                minimum = CInt(oFunction(1)(6)(1)(2).Text)
            End If
            
            If oFunction(1)(7).Index = 1 Then
                maximum = CInt(oFunction(1)(7)(1)(2).Text)
            End If

            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, CompileFunction(oFunction(1)(3)), ControlType(oFunction(1)(3)), CompileFunction(oFunction(1)(5)), ControlType(oFunction(1)(5)), maximum, minimum
        
        Case stSkip
            lSkipNo = 1
            If oFunction(1)(2).Index = 1 Then
                lSkipNo = CLng(oFunction(1)(2).Text)
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, lSkipNo
        
        Case stEOS
            CompileFunction.Initialise sObjectName, pcParseControl, False
            
        Case stBOS
            CompileFunction.Initialise sObjectName, pcParseControl, False
                        
        Case stPass
            CompileFunction.Initialise sObjectName, pcParseControl, False, True
            
        Case stFail
            CompileFunction.Initialise sObjectName, pcParseControl, False, False
            
        Case stExternal
            ' CompileFunction
            Select Case oFunction(1)(2)(2).Index
                Case 0 ' No parameters
                    CompileFunction.Initialise sObjectName, pcParseControl, False, oFunction(1)(2)(1).Text, ""
                Case 1
                    CompileFunction.Initialise sObjectName, pcParseControl, False, oFunction(1)(2)(1).Text, oFunction(1)(2)(2)(1)(2).Text
            End Select
            
        Case stEach
            Set oParams = New Collection
            For Each oObj In oFunction(1)(3).SubTree
                oParams.Add CompileFunction(oObj)
                oParams.Add ControlType(oObj)
            Next

            bNullControl = False
            If oFunction(1)(1).Index = 1 Then
                bNullControl = oFunction(1)(1)(1).Index = 1
            End If
            CompileFunction.Initialise sObjectName, pcParseControl, bNullControl, oParams
            
        Case stVariable
            On Error Resume Next
            Set CompileFunction = Objs(oFunction.Text)
            If Err.Number = 5 Then
                On Error GoTo 0
                Err.Clear
                Err.Raise 5, , "Referenced definition '" & oFunction.Text & "' not found."
            End If
    End Select
End Function

Friend Function Create(sType) As IParseObject
    Dim obj As IParseObject
    Set obj = CreateObject("PDLClasses.P" & sType)
    Set Create = obj
End Function

Private Function ResolveString(ByVal oFunction As ParseTree) As String
    Dim oSubFunction As ParseTree
    Dim iSubCount As Integer
    Dim bCharSet(255) As Boolean
    Dim bNot As Boolean
    Dim bCase As Boolean
    Dim sCharList As String
    Dim iPos As Integer
    Dim iStartRange As Integer
    Dim iEndRange As Integer
    Dim iChar As Integer
    Dim iUpperCase As Integer
    Dim iLowerCase As Integer
    
    For iSubCount = 1 To oFunction(2).SubTree.Count + 1
        If iSubCount = 1 Then
            Set oSubFunction = oFunction(1)
        Else
            Set oSubFunction = oFunction(2)(iSubCount - 1)(2)
        End If
        
        bNot = oSubFunction.SubTree(1).Index = 1
        bCase = oSubFunction.SubTree(2).Index = 1
        
        Select Case oSubFunction.SubTree(3).Index
            Case 1 ' Range
                If oSubFunction(3)(1)(1).Index = 1 Then
                    iStartRange = oSubFunction(3)(1)(1).Text
                Else
                    iStartRange = Asc(oSubFunction(3)(1)(1)(1)(2).Text)
                End If
                If oSubFunction(3)(1)(3).Index = 1 Then
                    iEndRange = oSubFunction(3)(1)(3).Text
                Else
                    iEndRange = Asc(oSubFunction(3)(1)(3)(1)(2).Text)
                End If
                
                For iPos = iStartRange To iEndRange
                    bCharSet(iPos) = Not bNot
                                        
                    If bCase Then
                        iLowerCase = Asc(LCase(Chr(iPos)))
                        iUpperCase = Asc(UCase(Chr(iPos)))
                        bCharSet(iLowerCase) = Not bNot
                        bCharSet(iUpperCase) = Not bNot
                    End If
                Next
                
            Case 2 ' Literal
                sCharList = oSubFunction.SubTree(3)(1)(2).Text
                For iPos = 1 To Len(sCharList)
                
                    bCharSet(Asc(Mid(sCharList, iPos, 1))) = Not bNot
                                        
                    If bCase Then
                        iLowerCase = Asc(LCase(Mid(sCharList, iPos, 1)))
                        iUpperCase = Asc(UCase(Mid(sCharList, iPos, 1)))
                        bCharSet(iLowerCase) = Not bNot
                        bCharSet(iUpperCase) = Not bNot
                    End If
                Next
                
            Case 3 ' Figure
                iChar = oSubFunction(3).Text
                bCharSet(iChar) = Not bNot
                                    
                If bCase Then
                    iLowerCase = Asc(LCase(Chr(iChar)))
                    iUpperCase = Asc(UCase(Chr(iChar)))
                    bCharSet(iLowerCase) = Not bNot
                    bCharSet(iUpperCase) = Not bNot
                End If
                    
        End Select
    Next

ExitPoint:
    For iPos = 0 To 255
        If bCharSet(iPos) Then
            ResolveString = ResolveString & Chr(iPos)
        End If
    Next
End Function

Private Function ReduceDoubleSingleQuotes(ByVal sString As String) As String
    Dim pos As Long
    Dim quotes As Long
    
    pos = 1
    quotes = InStr(pos, sString, "''")
    
    While quotes > 0
        ReduceDoubleSingleQuotes = ReduceDoubleSingleQuotes & Mid$(sString, pos, quotes - pos) & "'"
        pos = quotes + 2
        quotes = InStr(pos, sString, "''")
    Wend
    ReduceDoubleSingleQuotes = ReduceDoubleSingleQuotes & Mid$(sString, pos)
End Function
