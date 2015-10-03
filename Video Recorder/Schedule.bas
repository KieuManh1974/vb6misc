Attribute VB_Name = "Schedule"
Option Explicit

Private vTVChannels As Variant
Public vTVChannelNames As Variant
Private vRadioChannels As Variant
Public vRadioChannelNames As Variant

Public Const MaxSlots As Long = 7
Public oProgrammes(0 To MaxSlots) As Programme
Private Const StartupTime As Date = #12:00:52 AM#
Private Const LeadInTime As Date = #12:04:00 AM#
Private Const LeadOutTime As Date = #12:05:00 AM#

Public Function CheckProgrammes() As Boolean
    Dim oProgramme As Programme
    Dim dDate As Date
    Dim sFileName As String
    Dim iProgrammeIndex As Integer
    
    For iProgrammeIndex = 0 To MaxSlots
        Set oProgramme = oProgrammes(iProgrammeIndex)
        If oProgramme.Valid Then
            If (oProgramme.mStartTime - StartupTime - LeadInTime) <= Now And oProgramme.mStopTime >= Now Then
                RecordChannel oProgramme.mChannel, (oProgramme.mStopTime - Now - StartupTime + LeadInTime + LeadOutTime) * CLng(86400), oProgramme.Radio

                If Not oProgramme.mRadio Then
                    sFileName = vTVChannelNames(oProgramme.mChannel) & " " & Format$(oProgramme.mStartTime, "YYYY-MM-DD HHMM") & "-" & Format$(oProgramme.mStopTime, "HHMM") & ".avi"
                Else
                    sFileName = vRadioChannelNames(oProgramme.mChannel) & " " & Format$(oProgramme.mStartTime, "YYYY-MM-DD HHMM") & "-" & Format$(oProgramme.mStopTime, "HHMM") & ".avi"
                End If
                Debug.Print sFileName
                With New FileSystemObject
                    If .FileExists("D:\Media\Video\Captured\capture.avi") Then
                        On Error Resume Next
                        .GetFile("D:\Media\Video\Captured\capture.avi").Name = sFileName
                    End If
                End With
                oProgramme.Recorded = True
                CheckProgrammes = SetNextProgramme(oProgramme)
            End If
        End If
    Next
End Function

Private Function SetNextProgramme(oCopyProgramme As Programme) As Boolean
    Dim iProgrammeIndex As Long
    Dim oProgramme As Programme
    Dim dNextDate As Date
    Dim dNextDay As Integer
    Dim dNextMonth As Integer
    Dim dNextYear As Integer
    Dim dNextStartTime As Date
    Dim dNextStopTime As Date
    Dim iOffset As Integer
    Dim bCreateNew As Boolean
    
    With oCopyProgramme
        If oCopyProgramme.mDaily Then
            iOffset = 1
        ElseIf oCopyProgramme.mMonFri Then
            iOffset = 1
            If .mWeekday = "FRIDAY" Then
                iOffset = 3
            End If
        ElseIf oCopyProgramme.mWeekly Then
            iOffset = 7
        End If
    
        If oCopyProgramme.mDaily Or oCopyProgramme.mMonFri Or oCopyProgramme.mWeekly Then
            dNextDate = oCopyProgramme.mDate + iOffset
            dNextDay = Val(Format$(dNextDate, "DD"))
            dNextMonth = Val(Format$(dNextDate, "MM"))
            dNextYear = Val(Format$(dNextDate, "YYYY"))
            dNextStartTime = oCopyProgramme.mStartTime + iOffset
            dNextStopTime = oCopyProgramme.mStopTime + iOffset
            bCreateNew = True
        End If
    End With
    
    If bCreateNew Then
        For iProgrammeIndex = 0 To MaxSlots
            Set oProgramme = oProgrammes(iProgrammeIndex)
            If Not oProgramme.Valid Then
                With oProgramme
                    .mCurrentDay = oCopyProgramme.mCurrentDay
                    .mCurrentMonth = oCopyProgramme.mCurrentMonth
                    .mCurrentYear = oCopyProgramme.mCurrentYear
                    .mPlusCode = oCopyProgramme.mPlusCode
                    .mWeekday = oCopyProgramme.mWeekday
                    .mDate = dNextDate
                    .mDay = dNextDay
                    .mMonth = dNextMonth
                    .mYear = dNextYear
                    .mChannel = oCopyProgramme.mChannel
                    .mStartTime = dNextStartTime
                    .mStopTime = dNextStopTime
                    .mDuration = oCopyProgramme.mDuration
                    .mRadio = oCopyProgramme.mRadio
                    .mDaily = oCopyProgramme.mDaily
                    .mWeekly = oCopyProgramme.mWeekly
                    .mMonFri = oCopyProgramme.mMonFri
                    .mRecorded = False
                    .mStatus = Ready
                    .mValid = True
                End With
                SetNextProgramme = True
                WriteFile
                Exit Function
            End If
        Next
    End If
End Function

Public Sub RecordChannel(iChannelNo As Long, iDuration As Long, bRadio As Boolean)
    Dim sFileName As String
    Show.tmrTime.Enabled = False
    If Not bRadio Then
        Record vTVChannels(iChannelNo), iDuration, bRadio
    Else
        Record vRadioChannels(iChannelNo), iDuration, bRadio
    End If
    Show.tmrTime.Enabled = True
End Sub

Public Sub Initialise()
    Dim iIndex As Integer
    Dim oTS As TextStream
    Dim oFSO As New FileSystemObject
    
    vTVChannelNames = Array("", "BBC1", "BBC2", "ITV", "CHANNEL4", "CHANNEL5", "FREEVIEW", "VIDEO")
    vTVChannels = Array(0, 57, 63, 60, 53, 35, 45, 65)
    vRadioChannelNames = Array("", "Radio1", "Radio2", "Radio3", "Radio4")
    vRadioChannels = Array(0, "9820", "0000", "9010", "9450")
    For iIndex = 0 To MaxSlots
        Set oProgrammes(iIndex) = New Programme
    Next
    
    ReadFile
End Sub
