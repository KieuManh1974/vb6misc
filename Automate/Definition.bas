Attribute VB_Name = "Definition"
Option Explicit

Public oProcesses As Dictionary

Public oParser As IParseObject

Public Sub Initialise()
    Dim sDef As String
    
    sDef = sDef & "ws := [#(:0-32)-0];"
    sDef = sDef & "identifier := {&(#:'0'-'9',^'A'-'Z'),['.']};"
    sDef = sDef & "path := &[39],#(:0-255,!39),[39];"
    sDef = sDef & "digits := #:'0'-'9';"
    sDef = sDef & "open := & ws,identifier, ^'open', ws, path,ws;"
    sDef = sDef & "close := & ws,identifier, ^'close',ws;"
    sDef = sDef & "wait := & ws,identifier,^'wait',ws,digits,ws;"
    sDef = sDef & "pause := & ws,^'pause',ws,digits,ws;"
    sDef = sDef & "text := &ws,[39],(#OR(:0-255,!39),39+39),[39],ws;"
    sDef = sDef & "key := &ws,(|^'tab', ^'caps',^'escape',^'shiftdown',^'shiftup',^'ctrldown',^'ctrlup',^'alt',^'delete',^'return',^'enter',^'back'),ws;"
    sDef = sDef & "settime := & ws,^'settime',ws,{other};"
    sDef = sDef & "restoretime := & ws,^'restoretime',ws;"
    sDef = sDef & "sequence := @(|text, key), [&ws,',',ws];"
    sDef = sDef & "other := &(#:0-255:|13+10,eos),13+10;"
    sDef = sDef & "command := |open,close,wait,pause,sequence,settime,restoretime,other;"
    sDef = sDef & "script := #command;"
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParser = ParserObjects("script")
    Set oProcesses = New Dictionary
End Sub
