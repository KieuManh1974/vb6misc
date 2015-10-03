Attribute VB_Name = "AutomateScript"
Option Explicit

'E:\Development\automate\test.aut
Public Sub ExecuteScript(oTree As ParseTree)
    Dim oCommand As ParseTree
    Dim lTime As Long
    
    For Each oCommand In oTree.SubTree
        Select Case oCommand.Index
            Case 1 ' open
                If Not oProcesses.Exists(LCase$(oCommand(1)(1).Text)) Then
                    oProcesses.Add LCase$(oCommand(1)(1).Text), Process.CreateProcessX(oCommand(1)(3).Text)
                End If
            Case 2 ' close
                If oProcesses.Exists(LCase$(oCommand(1)(1).Text)) Then
                    TerminateProcessX oProcesses.Item(LCase$(oCommand(1)(1).Text))
                End If
            Case 3 ' wait
                If oProcesses.Exists(LCase$(oCommand(1)(1).Text)) Then
                    WaitForSingleObject oProcesses.Item(LCase$(oCommand(1)(1).Text)), CLng(oCommand(1)(3).Text) * 1000
                End If
            Case 4 ' pause
                lTime = CLng(oCommand(1)(2).Text)
                StartCounter
                While GetCounter < lTime
                    DoEvents
                Wend
            Case 5 ' sequence
                ExecuteSequence oCommand(1)
            Case 6 ' settime
                If IsDate(oCommand(1)(2).Text) Then
                    SetTime CDate(oCommand(1)(2).Text)
                Else
                    MsgBox "Date format not recognised when setting time."
                    End
                End If
            Case 7 ' restoretime
                RestoreTime
            Case 8 ' other
                ' do nothing
        End Select
    Next
End Sub

Public Sub ExecuteSequence(oTree As ParseTree)
    Dim oSequence As ParseTree
    Dim iIndex As Integer
    Dim sText As String
    Dim sChar As String
    
    For Each oSequence In oTree.SubTree
        Select Case oSequence.Index
            Case 1 'text
                sText = Replace$(oSequence.Text, "''", "'")
                sText = Replace$(sText, vbCrLf, vbCr)
                
                For iIndex = 1 To Len(sText)
                    sChar = Mid$(sText, iIndex, 1)
                    If InStr(")!""£$%^&*(", sChar) > 0 Then
                        KeyDown vbKeyShift
                        KeyPress InStr(")!""£$%^&*(", sChar) + vbKey0 - 1
                        KeyUp vbKeyShift
                    ElseIf InStr("-=[];'#,./\", sChar) > 0 Then
                        KeyPress KeyCode(sChar)
                    ElseIf InStr("_+{}:@~<>?|", sChar) > 0 Then
                        KeyDown vbKeyShift
                        KeyPress KeyCode(Mid$("-=[];'#,./\", InStr("_+{}:@~<>?|", sChar), 1))
                        KeyUp vbKeyShift
                    Else
                        If CapsLock Then
                            If UCase$(sChar) <> sChar Then
                                KeyDown vbKeyShift
                                KeyPress Asc(UCase$(Mid$(sText, iIndex, 1)))
                                KeyUp vbKeyShift
                            Else
                                KeyPress Asc(UCase$(Mid$(sText, iIndex, 1)))
                            End If
                        Else
                            If LCase$(sChar) <> sChar Then
                                KeyDown vbKeyShift
                                KeyPress Asc(UCase$(Mid$(sText, iIndex, 1)))
                                KeyUp vbKeyShift
                            Else
                                KeyPress Asc(UCase$(Mid$(sText, iIndex, 1)))
                            End If
                        End If
                    End If
                Next
            Case 2 'key
                Select Case UCase$(oSequence.Text)
                    Case "TAB"
                        KeyPress vbKeyTab
                    Case "CAPS"
                        KeyPress vbKeyCapital
                    Case "ESCAPE"
                        KeyPress vbKeyEscape
                    Case "SHIFTDOWN"
                        KeyDown vbKeyShift
                    Case "SHIFTUP"
                        KeyUp vbKeyShift
                    Case "CTRLDOWN"
                        KeyDown vbKeyControl
                    Case "CTRLUP"
                        KeyUp vbKeyControl
                    Case "ALT"
                        KeyPress vbKeyMenu
                    Case "DELETE"
                        KeyPress vbKeyDelete
                    Case "RETURN"
                        KeyPress vbKeyReturn
                    Case "LEFT"
                        KeyPress vbKeyLeft
                    Case "RIGHT"
                        KeyPress vbKeyRight
                    Case "UP"
                        KeyPress vbKeyUp
                    Case "DOWN"
                        KeyPress vbKeyDown
                    Case "BACK"
                        KeyPress vbKeyBack
                End Select
        End Select
    Next
End Sub

