Imports Scripting
Imports PDLClasses
Imports PDLCompiler
Imports System.Math

Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtEntry As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader()
        Me.txtEntry = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtEntry
        '
        Me.txtEntry.AcceptsReturn = CType(configurationAppSettings.GetValue("txtEntry.AcceptsReturn", GetType(System.Boolean)), Boolean)
        Me.txtEntry.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtEntry.Font = New System.Drawing.Font("Arial", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEntry.Location = New System.Drawing.Point(8, 8)
        Me.txtEntry.Multiline = True
        Me.txtEntry.Name = "txtEntry"
        Me.txtEntry.Size = New System.Drawing.Size(272, 248)
        Me.txtEntry.TabIndex = 0
        Me.txtEntry.Text = ""
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(292, 266)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtEntry})
        Me.Name = "Form1"
        Me.Text = "Calculator"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private oVariables As New Dictionary()
    Private oStatement As IParseObject
    Private oPDLCompiler As New PDLObject()
    Private oPDLClasses As New PDLClasses.ParserTextString()
    Private ErrorSource As String

    Private Enum FunctionTypes
        ftLoop = 1
        ftFactorial
        ftPercent
        ftTime
        ftRadixNumber
        ftNumber
        ftFunctionExpression
        ftConstant
        ftFunctionVariableCall
        ftUnaryExpression
        ftBracketExpression
    End Enum

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        InitialiseDefinition()
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        SaveSetting("Calculator", "Dimensions", "Width", Me.Width)
        SaveSetting("Calculator", "Dimensions", "Height", Me.Height)
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Resize
        txtEntry.Width = Me.ClientSize.Width - 16
        txtEntry.Height = Me.ClientSize.Height - 16
    End Sub

    Private Sub InitialiseDefinition()
        Dim sDefinition As String

        Me.Width = GetSetting("Calculator", "Dimensions", "Width", Me.Width)
        Me.Height = GetSetting("Calculator", "Dimensions", "Height", Me.Height)

        sDefinition = "ws := [REPEAT ' ' MIN 0];" & _
                            "digits := {REPEAT IN '0' TO '9'};" & _
                            "radixnumber := AND {REPEAT IN '0' TO '9', CASE 'A' TO 'F'}, CASE 'R', digits;" & _
                            "fraction := AND '.', digits; " & _
                            "exponent := AND CASE 'E', OPTIONAL IN '+-', digits;" & _
                            "number := {AND OPTIONAL '-', digits, OPTIONAL fraction, {OPTIONAL exponent}}; " & _
                            "minsec := AND digits, ':', digits, OPTIONAL fraction;" & _
                            "hourminsec := AND digits, ':', digits, ':', digits, OPTIONAL fraction;" & _
                            "time := OR hourminsec, minsec;" & _
                            "function := OR CASE 'sin', " & _
                            "                 CASE 'cos', " & _
                            "                 CASE 'tan', " & _
                            "                 CASE 'cot', " & _
                            "                 CASE 'csc', " & _
                            "                 CASE 'sec', " & _
                            "                 CASE 'exp', " & _
                            "                 (AND (CASE 'log'), OPTIONAL digits)," & _
                            "                 (AND (CASE 'radix'), digits)," & _
                            "                 CASE 'rad'," & _
                            "                 CASE 'deg'," & _
                            "                 CASE 'int'," & _
                            "                 CASE 'frac'," & _
                            "                 CASE 'sqr'," & _
                            "                 CASE 'atn'," & _
                            "                 CASE 'asn',"

        sDefinition = sDefinition & _
                            "                 CASE 'acs'," & _
                            "                 CASE 'gam'," & _
                            "                 CASE 'dms'," & _
                            "                 (AND (CASE 'fix'), digits);" & _
                            "loop_operator := AND (OR ^'loop', ^'show'), (OPTIONAL IN '+', '-', '*', '/', '\');" & _
                            "loop_params := AND ['('], ws, level0, ws, [','], ws, level0, ws, [')'];" & _
                            "loop := AND loop_operator, loop_params, ws, level0, ws, [':'], ?variable;" & _
                            "function_variable_assign := AND variable, OPTIONAL (AND ['('], (LIST variable, [AND ws, ',', ws]), [')']);" & _
                            "function_variable_call := AND variable, OPTIONAL(AND ['('], (LIST level0, [AND ws, ',', ws]), [')']);" & _
                            "dummy := IN 0;" & _
                            "level0 := LIST level1, (AND ws, (OR CASE 'and', CASE 'or', CASE 'xor', CASE 'mod'), ws);" & _
                            "level1 := LIST level2, (AND ws, (OR '+','-'), ws);" & _
                            "level2 := LIST level3, (AND ws, (OR '*','/','\'), ws);" & _
                            "level3 := LIST level4, (AND ws, '^', ws);" & _
                            "level4 := OR loop, factorial, percent, time, radixnumber, number, functionexpression, constant, function_variable_call, unaryexpression, bracketexpression;"

        sDefinition = sDefinition & _
                            "factorial := AND number, ['!'];" & _
                            "percent := AND number, ['%'];" & _
                            "functionexpression := AND function, ws, ['('], ws, level0, ws, [')'], ws;" & _
                            "constant := OR CASE 'pi', CASE 'e';" & _
                            "variable := {AND (IN 'A' TO 'Z', 'a' TO 'z'), (REPEAT IN 'A' TO 'Z', 'a' TO 'z', '0' TO '9' MIN 0)};" & _
                            "unaryexpression := AND (OR '+', '-'), level0;" & _
                            "bracketexpression := AND '(', level0, ')';" & _
                            "assignment := AND variable, ws, '=', ws, level0;" & _
                            "definition := AND function_variable_assign, ws, ':=', ws, REPEAT IN 0 TO 255 UNTIL EOS;" & _
                            "program := OR assignment, definition, level0;"

        If Not oPDLCompiler.SetNewDefinition(sDefinition) Then
            System.Diagnostics.Debug.Write(oPDLCompiler.ErrorString)
            Stop
        End If

        oStatement = oPDLCompiler.ParserObjects.Item("program")
    End Sub

    Private Function InsertText(ByVal sText As String)
        Dim iTextPos As Integer

        iTextPos = txtEntry.SelectionStart
        txtEntry.Text = Strings.Left(txtEntry.Text, iTextPos) & sText & Mid$(txtEntry.Text, iTextPos + 1)

        txtEntry.SelectionStart = iTextPos + Len(sText)
        txtEntry.ScrollToCaret()
    End Function

    Private Sub txtEntry_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtEntry.KeyPress
        On Error GoTo ShowError

        Dim sCommand As String
        Dim vLines As Object
        Dim oResult As New ParseTree()
        Dim lCursorPos As Long
        Dim sAnswer As String

        If e.KeyChar = Chr(13) Then
            lCursorPos = txtEntry.SelectionStart
            vLines = Split(Strings.Left(txtEntry.Text, lCursorPos), vbCrLf)
            If UBound(vLines) <> -1 Then
                sCommand = vLines(UBound(vLines))
            End If
            e.Handled = True
            oPDLClasses.ParserText = sCommand
            If oStatement.Parse(oResult) Then
                If oPDLClasses.ParserTextPosition = Len(sCommand) + 1 Then
                    sAnswer = EvaluateCommand(oResult)
                    If sAnswer <> "" Then
                        InsertText(vbCrLf & sAnswer & vbCrLf & vbCrLf)
                    Else
                        InsertText(vbCrLf)
                    End If
                Else
                    InsertText(vbCrLf & "Syntax Error" & vbCrLf & vbCrLf)
                End If
            Else
                If Trim$(sCommand) <> "" Then
                    InsertText(vbCrLf & "Syntax Error" & vbCrLf & vbCrLf)
                Else
                    InsertText(vbCrLf)
                End If
            End If
        End If
        Exit Sub
ShowError:
        'If Err.Description <> "" Then
        '    InsertText(vbCrLf & Err.Description & vbCrLf & vbCrLf)
        'Else
        '    InsertText(vbCrLf & "Error" & vbCrLf & vbCrLf)
        'End If
        'Err.Clear()
        'ErrorSource = ""
    End Sub

    Private Function EvaluateCommand(ByVal oExpression As ParseTree) As String
        On Error GoTo ExitPoint
        Dim sErrorSource As String
        Dim oNewVar As New Variable()

        Select Case oExpression.Index
            Case 1 ' Value assignment
                If oVariables.Exists(oExpression(1)(1).Text) Then
                    oVariables.Remove(CStr(oExpression(1)(1).Text))
                    oNewVar.Value = CStr(EvaluateLevel0(oExpression(1)(3), ))
                    oVariables.Add(CStr(oExpression(1)(1).Text), oNewVar)
                Else
                    oNewVar.Value = CStr(EvaluateLevel0(oExpression(1)(3), ))
                    oVariables.Add(CStr(oExpression(1)(1).Text), oNewVar)
                End If

            Case 2 ' Function Assigment
                Dim sVariableName As String
                sVariableName = oExpression(1)(1)(1).Text
                If oVariables.Exists(sVariableName) Then
                    oVariables.Remove(CStr(sVariableName))
                End If

                ' Check Expression
                Dim sExpression As String
                Dim oLevel0 As IParseObject
                Dim oResult As New ParseTree()

                oLevel0 = oPDLCompiler.ParserObjects.Item("level0")
                sExpression = oExpression(1)(3).Text
                oPDLClasses.ParserText = sExpression

                If Not oLevel0.Parse(oResult) Then
                    Err.Raise(-1)
                    Exit Function
                ElseIf oPDLClasses.ParserTextPosition <> Len(sExpression) + 1 Then
                    Err.Raise(-1)
                    Exit Function
                End If

                oNewVar.Expression = oExpression(1)(3).Text

                ' Any parameters?
                Dim oParameterVar As Variable
                Dim oParameter As ParseTree
                If oExpression(1)(1)(2).Index > 0 Then
                    oNewVar.Parameters = New Dictionary()
                    For Each oParameter In oExpression(1)(1)(2)(1)(1).SubTree
                        oParameterVar = New Variable()
                        oNewVar.Parameters.Add(CStr(oParameter.Text), oParameterVar)
                    Next
                End If

                oVariables.Add(CStr(sVariableName), oNewVar)

            Case 3
                EvaluateCommand = EvaluateLevel0(oExpression(1))
        End Select
        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function EvaluateLevel0(ByVal oExpression As ParseTree, Optional ByVal oParameters As Dictionary = Nothing) As String
        Dim runningvalue As String
        Dim vPart As Object
        Dim lPartIndex As Long
        Dim tempvalue As String

        On Error GoTo ExitPoint
        Dim sErrorSource As String

        runningvalue = EvaluateLevel1(oExpression(1), oParameters)
        For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
            vPart = oExpression(lPartIndex)
            Select Case UCase(vPart(1).Text)
                Case "AND"
                    runningvalue = runningvalue And EvaluateLevel1(oExpression(lPartIndex + 1), oParameters)
                Case "OR"
                    runningvalue = runningvalue Or EvaluateLevel1(oExpression(lPartIndex + 1), oParameters)
                Case "XOR"
                    runningvalue = runningvalue Xor EvaluateLevel1(oExpression(lPartIndex + 1), oParameters)
                Case "MOD"
                    tempvalue = EvaluateLevel1(oExpression(lPartIndex + 1), oParameters)
                    runningvalue = ((runningvalue / tempvalue) - Int(runningvalue / tempvalue)) * tempvalue
            End Select
        Next

        EvaluateLevel0 = runningvalue

        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function EvaluateLevel1(ByVal oExpression As ParseTree, Optional ByVal oParameters As Dictionary = Nothing) As String
        Dim runningvalue As String
        Dim vPart As Object
        Dim lPartIndex As Long

        On Error GoTo ExitPoint
        Dim sErrorSource As String

        runningvalue = EvaluateLevel2(oExpression(1), oParameters)
        For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
            vPart = oExpression(lPartIndex)
            Select Case UCase(vPart(1).Text)
                Case "+"
                    runningvalue = runningvalue + CDbl(EvaluateLevel2(oExpression(lPartIndex + 1), oParameters))
                Case "-"
                    runningvalue = runningvalue - CDbl(EvaluateLevel2(oExpression(lPartIndex + 1), oParameters))
            End Select
        Next

        EvaluateLevel1 = runningvalue

        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function EvaluateLevel2(ByVal oExpression As ParseTree, Optional ByVal oParameters As Dictionary = Nothing) As String
        Dim runningvalue As String
        Dim vPart As Object
        Dim lPartIndex As Long

        On Error GoTo ExitPoint
        Dim sErrorSource As String

        runningvalue = EvaluateLevel3(oExpression(1), oParameters)
        For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
            vPart = oExpression(lPartIndex)
            Select Case UCase(vPart(1).Text)
                Case "*"
                    runningvalue = runningvalue * EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)
                Case "/"
                    sErrorSource = "DIVIDE"
                    runningvalue = runningvalue / EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)
                Case "\"
                    runningvalue = runningvalue \ EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)
            End Select
        Next

        EvaluateLevel2 = runningvalue

        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function EvaluateLevel3(ByVal oExpression As ParseTree, Optional ByVal oParameters As Dictionary = Nothing) As String
        Dim runningvalue As String
        Dim vPart As Object
        Dim lPartIndex As Long

        On Error GoTo ExitPoint
        Dim sErrorSource As String

        runningvalue = EvaluateLevel4(oExpression(1), oParameters)
        For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
            vPart = oExpression(lPartIndex)
            Select Case UCase(vPart(1).Text)
                Case "^"
                    runningvalue = runningvalue ^ EvaluateLevel4(oExpression(lPartIndex + 1), oParameters)
            End Select
        Next

        EvaluateLevel3 = runningvalue

        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function EvaluateLevel4(ByVal oExpression As ParseTree, Optional ByVal oParameters As Dictionary = Nothing) As String
        Dim runningvalue As String
        Dim vPart As Object
        Dim temp As Double

        Dim hour As Double
        Dim minute As Integer
        Dim second As Integer
        Dim fraction As Double

        On Error GoTo ExitPoint
        Dim sErrorSource As String

        Select Case oExpression.Index
            Case FunctionTypes.ftLoop
                Dim iLower As Double
                Dim iUpper As Double
                Dim sOperator As String
                Dim sVariable As String
                Dim iLoopIndex As Double
                Dim bShowResult As Boolean
                Dim oVariable As Variable

                iLower = EvaluateLevel0(oExpression(1)(2)(1), oParameters)
                iUpper = EvaluateLevel0(oExpression(1)(2)(2), oParameters)
                sOperator = oExpression(1)(1)(2).Text
                sVariable = oExpression(1)(4).Text
                Select Case LCase$(oExpression(1)(1)(1).Text)
                    Case "loop"
                        bShowResult = False
                    Case "show"
                        bShowResult = True
                End Select

                If oVariables.Exists(sVariable) Then
                    oVariables.Remove(sVariable)
                End If

                oVariable = New Variable()

                oVariables.Add(sVariable, oVariable)

                Select Case sOperator
                    Case "+", "-"
                        runningvalue = 0
                    Case "*", "/"
                        runningvalue = 1
                End Select

                For iLoopIndex = iLower To iUpper
                    oVariables.Item(sVariable).Value = iLoopIndex
                    Select Case sOperator
                        Case ""
                            runningvalue = EvaluateLevel0(oExpression(1)(3), oParameters)
                        Case "+"
                            runningvalue = Val(runningvalue) + EvaluateLevel0(oExpression(1)(3), oParameters)
                        Case "-"
                            runningvalue = Val(runningvalue) - EvaluateLevel0(oExpression(1)(3), oParameters)
                        Case "*"
                            runningvalue = Val(runningvalue) * EvaluateLevel0(oExpression(1)(3), oParameters)
                        Case "/"
                            runningvalue = Val(runningvalue) / EvaluateLevel0(oExpression(1)(3), oParameters)
                    End Select
                    If bShowResult Then
                        InsertText(vbCrLf & runningvalue)
                    End If
                Next
                If bShowResult Then
                    runningvalue = ""
                End If

            Case FunctionTypes.ftFactorial
                runningvalue = Factorial(CDbl(oExpression.Text))

            Case FunctionTypes.ftPercent
                runningvalue = CDbl(oExpression.Text) / 100

            Case FunctionTypes.ftTime
                hour = oExpression(1)(1)(1).Text
                minute = oExpression(1)(1)(3).Text
                'second = oExpression(3).Text
                'fraction = oExpression(4).Text
                runningvalue = hour + minute / 60

            Case FunctionTypes.ftRadixNumber
                runningvalue = Base(oExpression(1)(1).Text, CLng(oExpression(1)(3).Text), 10)

            Case FunctionTypes.ftNumber
                sErrorSource = "NUMBER"
                runningvalue = CDbl(oExpression.Text)

            Case FunctionTypes.ftFunctionExpression
                Select Case UCase(oExpression(1)(1).Text)
                    Case "SIN"
                        runningvalue = Sin(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "COS"
                        runningvalue = Cos(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "TAN"
                        runningvalue = Tan(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "EXP"
                        runningvalue = Exp(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "LOG"
                        runningvalue = Log(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "DEG"
                        runningvalue = 360 * (EvaluateLevel0(oExpression(1)(2), oParameters)) / 8 / Atan(1)
                    Case "RAD"
                        runningvalue = (EvaluateLevel0(oExpression(1)(2), oParameters)) * 8 * Atan(1) / 360
                    Case "INT"
                        runningvalue = Int(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "FRAC"
                        temp = EvaluateLevel0(oExpression(1)(2), oParameters)
                        runningvalue = temp - Int(temp)
                    Case "SQR"
                        sErrorSource = "SQR"
                        runningvalue = Sqrt(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "NOT"
                        runningvalue = Not (EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "ATN"
                        runningvalue = Atan(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "ACS"
                        temp = (EvaluateLevel0(oExpression(1)(2), oParameters))
                        runningvalue = 2 * Atan(1) - Atan(temp / Sqrt(1 - temp * temp))
                    Case "ASN"
                        temp = (EvaluateLevel0(oExpression(1)(2), oParameters))
                        runningvalue = Atan(temp / Sqrt(1 - temp * temp))
                    Case "GAM"
                        '   G[z] = Integral[t^(z-1) Exp[-t] dt, {t, 0, Infinity}]
                        Dim oldrunningvalue As Double
                        Dim t As Double
                        temp = (EvaluateLevel0(oExpression(1)(2), oParameters))
                        runningvalue = 0
                        oldrunningvalue = -1
                        t = 0.0001
                        While oldrunningvalue <> runningvalue
                            oldrunningvalue = runningvalue
                            runningvalue = runningvalue + CDbl((t ^ temp * Exp(-t)) * 0.0001)
                            t = t + 0.0001
                        End While
                    Case "COT"
                        runningvalue = 1 / Tan(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "CSC"
                        runningvalue = 1 / Sin(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "SEC"
                        runningvalue = 1 / Cos(EvaluateLevel0(oExpression(1)(2), oParameters))
                    Case "DMS"
                        runningvalue = EvaluateLevel0(oExpression(1)(2), oParameters)
                        hour = Int(runningvalue) Mod 24
                        hour = IIf(runningvalue < 0, 24 * (1 - (runningvalue + 1) \ 24) + runningvalue, (Int(runningvalue)) Mod 24)
                        runningvalue = Format(hour, "00") & ":" & Format((runningvalue - Int(runningvalue)) * 60, "00")
                    Case Else
                        If Strings.Left(UCase(oExpression(1)(1).Text), 3) = "LOG" Then
                            runningvalue = Log(EvaluateLevel0(oExpression(1)(2), oParameters)) / Log(oExpression(1)(1)(1)(2).Text)
                        End If
                        If Strings.Left(UCase(oExpression(1)(1).Text), 5) = "RADIX" Then
                            'runningvalue = Log(EvaluateLevel0(oExpression(1)(2))) / Log(oExpression(1)(1)(1)(2).Text)
                            runningvalue = Base(EvaluateLevel0(oExpression(1)(2), oParameters), 10, Val(oExpression(1)(1)(1)(2).Text))
                        End If
                        If Strings.Left(UCase(oExpression(1)(1).Text), 3) = "FIX" Then
                            Dim iDecimalPlaces As Long
                            iDecimalPlaces = Val(oExpression(1)(1)(1)(2).Text)
                            runningvalue = Int(EvaluateLevel0(oExpression(1)(2), oParameters) * 10 ^ iDecimalPlaces + 0.5) / 10 ^ iDecimalPlaces
                        End If
                End Select

                'Secant Sec(X) = 1 / Cos(X)
                'Cosecant Cosec(X) = 1 / Sin(X)
                'Cotangent Cotan(X) = 1 / Tan(X)
                'Inverse Sine Arcsin(X) = Atn(X / Sqr(-X * X + 1))
                'Inverse Cosine Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
                'Inverse Secant Arcsec(X) = 2 * Atn(1) – Atn(Sgn(X) / Sqr(X * X – 1))
                'Inverse Cosecant Arccosec(X) = Atn(Sgn(X) / Sqr(X * X – 1))
                'Inverse Cotangent Arccotan(X) = 2 * Atn(1) - Atn(X)
                'Hyperbolic Sine HSin(X) = (Exp(X) – Exp(-X)) / 2
                'Hyperbolic Cosine HCos(X) = (Exp(X) + Exp(-X)) / 2
                'Hyperbolic Tangent HTan(X) = (Exp(X) – Exp(-X)) / (Exp(X) + Exp(-X))
                'Hyperbolic Secant HSec(X) = 2 / (Exp(X) + Exp(-X))
                'Hyperbolic Cosecant HCosec(X) = 2 / (Exp(X) – Exp(-X))
                'Hyperbolic Cotangent HCotan(X) = (Exp(X) + Exp(-X)) / (Exp(X) – Exp(-X))
                'Inverse Hyperbolic Sine HArcsin(X) = Log(X + Sqr(X * X + 1))
                'Inverse Hyperbolic Cosine HArccos(X) = Log(X + Sqr(X * X – 1))
                'Inverse Hyperbolic Tangent HArctan(X) = Log((1 + X) / (1 – X)) / 2
                'Inverse Hyperbolic Secant HArcsec(X) = Log((Sqr(-X * X + 1) + 1) / X)
                'Inverse Hyperbolic Cosecant HArccosec(X) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
                'Inverse Hyperbolic Cotangent HArccotan(X) = Log((X + 1) / (X – 1)) / 2
                'Logarithm to base N LogN(X) = Log(X) / Log(N)

            Case FunctionTypes.ftConstant
                Select Case UCase(oExpression.Text)
                    Case "PI"
                        runningvalue = Atan(1) * 4
                    Case "E"
                        runningvalue = Exp(1)
                End Select

            Case FunctionTypes.ftFunctionVariableCall
                Dim sVariableName As String
                Dim iParameterIndex As Long
                Dim oParameterValue As ParseTree
                Dim oParameterVariable As Dictionary

                sErrorSource = "VARIABLE"

                sVariableName = oExpression(1)(1).Text

                If oExpression(1)(2).Index = 0 Then ' Has no parameters
                    If Not oParameters Is Nothing Then
                        If oParameters.Exists(sVariableName) Then
                            runningvalue = oParameters.Item(sVariableName).Value
                            EvaluateLevel4 = runningvalue
                            Exit Function
                        End If
                    End If
                End If

                If oVariables.Exists(sVariableName) Then
                    If oExpression(1)(2).Index = 0 Then
                        If Not oVariables.Item(sVariableName).Parameters Is Nothing Then
                            Err.Raise(-1)
                        End If
                    Else
                        If oExpression(1)(2)(1)(1).Index <> oVariables.Item(sVariableName).Parameters.Count Then
                            Err.Raise(-1)
                        End If

                        iParameterIndex = 0
                        oParameterVariable = oVariables.Item(sVariableName).Parameters
                        For Each oParameterValue In oExpression(1)(2)(1)(1).SubTree
                            oParameterVariable.Items(iParameterIndex).Value = EvaluateLevel0(oParameterValue)
                            iParameterIndex = iParameterIndex + 1
                        Next
                    End If

                    Dim thevar As Variable
                    thevar = oVariables.Item(sVariableName)
                    If thevar.Expression = "" Then
                        runningvalue = thevar.Value
                    Else
                        Dim subdecode As IParseObject
                        Dim othisResult As New ParseTree()
                        Dim savetext As String
                        Dim savepos As String

                        savetext = oPDLClasses.ParserText
                        savepos = oPDLClasses.ParserTextPosition

                        subdecode = oPDLCompiler.ParserObjects.Item("level0")

                        oPDLClasses.ParserText = thevar.Expression
                        subdecode.Parse(othisResult)
                        runningvalue = EvaluateLevel0(othisResult, oVariables.Item(sVariableName).Parameters)
                        oPDLClasses.ParserText = savetext
                        oPDLClasses.ParserTextPosition = savepos
                    End If
                Else
                    sErrorSource = "VARIABLE"
                    Err.Raise(-3)
                End If

            Case FunctionTypes.ftUnaryExpression
                Select Case oExpression(1)(1).Text
                    Case "+"
                        runningvalue = EvaluateLevel0(oExpression(1)(2), oParameters)

                    Case "-"
                        runningvalue = -EvaluateLevel0(oExpression(1)(2), oParameters)
                End Select

            Case FunctionTypes.ftBracketExpression
                runningvalue = EvaluateLevel0(oExpression(1)(2), oParameters)

        End Select

        EvaluateLevel4 = runningvalue

        Exit Function
ExitPoint:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Sub Form_QueryUnload(ByVal Cancel As Integer, ByVal UnloadMode As Integer)
        SaveSetting("Calculator", "Dimensions", "Width", Me.Width)
        SaveSetting("Calculator", "Dimensions", "Height", Me.Height)
    End Sub

    Private Sub Form_Resize()
        On Error Resume Next
        txtEntry.Width = Me.ClientSize.Width - 16
        txtEntry.Height = Me.ClientSize.Height - 16
    End Sub

    Private Function Factorial(ByVal lnum As Double) As Double
        On Error GoTo ExitFunction
        Dim sErrorSource As String

        Dim q As Long

        sErrorSource = "FACTORIAL"

        If (lnum < 0) Or (Int(lnum) <> lnum) Then
            Err.Raise(-1)
        End If

        Factorial = 1
        For q = 2 To lnum
            Factorial = Factorial * q
        Next
        Exit Function
ExitFunction:
        Err.Raise(-1, , ErrorHandler(sErrorSource))
    End Function

    Private Function Factor(ByVal lnum As Double) As String
        Dim f As Double
        lnum = Int(lnum)
        f = 2
        While lnum > 1
            If (lnum / f) = Int(lnum / f) Then
                Factor = Factor & CStr(f) & " "
                lnum = lnum / Factor
                f = 2
            Else
                f = f + 1
            End If
        End While
    End Function

    Private Function Base(ByVal lnum As String, ByVal frombase As Integer, ByVal tobase As Integer) As String
        Const digits As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim digitpos As Long
        Dim power As Long
        Dim result As Long
        Dim digitvalue As Long

        power = 1
        For digitpos = Len(lnum) To 1 Step -1
            digitvalue = InStr(digits, Mid$(lnum, digitpos, 1)) - 1
            result = result + digitvalue * power
            power = power * frombase
        Next

        power = tobase
        While result > 0
            digitvalue = ((result / power) - Int(result / power)) * power
            Base = Mid$(digits, digitvalue + 1, 1) & Base
            result = Int(result / power)
        End While

    End Function

    Private Function ErrorHandler(ByVal sSource As String) As String
        If ErrorSource = "" Then
            ErrorSource = sSource
        Else
            ErrorHandler = Err.Description
            Exit Function
        End If

        Select Case sSource
            Case "DIVIDE"
                Select Case Err.Number
                    Case 11
                        ErrorHandler = "Division by Zero"
                End Select
            Case "SQR"
                Select Case Err.Number
                    Case 5
                        ErrorHandler = "Negative Square Root"
                End Select
            Case "NUMBER"
                Select Case Err.Number
                    Case 6
                        ErrorHandler = "Number too Large"
                End Select
            Case "VARIABLE"
                Select Case Err.Number
                    Case -1
                        ErrorHandler = "Wrong Number of Arguments"
                    Case -2
                        ErrorHandler = "Variable not Assigned"
                    Case -3
                        ErrorHandler = "Variable not Assigned"
                End Select
            Case "FACTORIAL"
                Select Case Err.Number
                    Case -1
                        ErrorHandler = "Factorial not Positive Integer"
                    Case Else
                        ErrorHandler = "Factorial too Large"
                End Select
            Case Else
                ErrorHandler = Err.Description
        End Select
    End Function

End Class
