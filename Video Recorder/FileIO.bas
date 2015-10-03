Attribute VB_Name = "FileIO"
Option Explicit

Public Sub ReadFile()
    Dim oTS As TextStream
    Dim oFSO As New FileSystemObject
    Dim iProgrammeIndex As Integer
    Dim vDetails As Variant
    Dim sDate As String
    
    If oFSO.FileExists(App.Path & "\programmes.ini") Then
        Set oTS = oFSO.OpenTextFile(App.Path & "\programmes.ini", ForReading)
        
        While Not oTS.AtEndOfStream And iProgrammeIndex <= MaxSlots
            vDetails = Split(oTS.ReadLine, "|", , vbTextCompare)
            With oProgrammes(iProgrammeIndex)
                .mStatus = Loading
                .PlusCode = vDetails(0)
                .mDate = CDate(vDetails(1))
                .mDay = Val(Format$(.mDate, "DD"))
                .mMonth = Val(Format$(.mDate, "MM"))
                .mYear = Val(Format$(.mDate, "YYYY"))
                .mWeekday = Format$(.mDate, "DDDD")
                .mStartTime = .mDate + CDate(vDetails(2))
                .mStopTime = .mDate + CDate(vDetails(3))
                If .mStopTime < .mStartTime Then
                    .mStopTime = .mStopTime + 1
                End If
                .mDuration = Int((.mStopTime - .mStartTime) * 24 * 60 + 0.1)
                .mChannel = vDetails(4)
                .mRadio = vDetails(6)
                .mValid = vDetails(5)
                .mDaily = vDetails(7)
                .mMonFri = vDetails(8)
                .mWeekly = vDetails(9)
                .mStatus = Ready
            End With
            iProgrammeIndex = iProgrammeIndex + 1
        Wend
        oTS.Close
    End If
    
End Sub

Public Sub WriteFile()
    Dim oTS As TextStream
    Dim oFSO As New FileSystemObject
    Dim vProgramme As Variant
    
    Set oTS = oFSO.CreateTextFile(App.Path & "\programmes.ini", True)
    
    For Each vProgramme In oProgrammes
        With vProgramme
            If .Valid Then
                oTS.WriteLine .mPlusCode & "|" & Format$(.mStartTime, "DD/MM/YYYY") & "|" & Format$(.mStartTime, "HH:MM") & "|" & Format$(.mStopTime, "HH:MM") & "|" & .mChannel & "|" & .mValid & "|" & .mRadio & "|" & .mDaily & "|" & .mMonFri & "|" & .mWeekly
            End If
        End With
    Next
    oTS.Close
End Sub
