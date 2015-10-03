VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Programme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public mStartTime As Date
Public mStopTime As Date
Public mChannel As Long
Public mPlusCode As Long
Public mValid As Boolean
Public mRecorded As Boolean
Public mRadio As Boolean
Public mDaily As Boolean
Public mMonFri As Boolean
Public mWeekly As Boolean
Public mWeekday As String
Public mDay As Integer
Public mMonth As Integer
Public mDate As Date
Public mYear As Integer
Public mCurrentYear As Integer
Public mCurrentMonth As Integer
Public mCurrentDay As Integer
Public mDuration As Integer

Public mStatus As StatusTypes

Public Enum StatusTypes
    Loading
    Ready
End Enum

Public Property Let CurrentDay(sValue As String)
    mCurrentDay = Val(sValue)
End Property
Public Property Let CurrentMonth(sValue As String)
    mCurrentMonth = Val(sValue)
    SetYear
End Property
Public Property Let CurrentYear(sValue As String)
    mCurrentYear = Val(sValue)
    SetYear
End Property

Private Sub SetYear()
    If mMonth <> 0 Then
        If mCurrentMonth <> 0 Then
            If mMonth < mCurrentMonth Then
                mYear = mCurrentYear + 1
            Else
                mYear = mCurrentYear
            End If
        End If
    End If
End Sub
Public Property Let day(sValue As String)
    Dim sDate As String
    
    sValue = Normal(sValue)
    mDay = Val(sValue)
    SetDate
    FindPlusCode
    TestValid
End Property

Public Property Let month(sValue As String)
    Dim sDate As String
    
    sValue = Normal(sValue)
    mMonth = Val(sValue)
    SetDate
    FindPlusCode
    TestValid
End Property

Private Sub SetDate()
    Dim sDate As String
    SetYear
    sDate = mDay & "/" & mMonth & "/" & mYear
    If IsDate(sDate) Then
        mDate = CDate(sDate)
        mWeekday = Format$(mDate, "dddd")
    End If
End Sub

Public Property Let StartTime(sValue As String)
    sValue = Normal(sValue)
    If (sValue) <> "" Then
        sValue = Format$(sValue, "0000")
        If IsDate(Left$(sValue, 2) & ":" & Right$(sValue, 2)) Then
            mStartTime = mDate + CDate(Left$(sValue, 2) & ":" & Right$(sValue, 2))
            FindPlusCode
        End If
    End If
    TestValid
End Property

Private Function Normal(sValue As String) As String
    Normal = Replace$(sValue, "-", "")
End Function
Public Property Let StopTime(sValue As String)
    sValue = Normal(sValue)
    If (sValue) <> "" Then
        sValue = Format$(sValue, "0000")
        mStopTime = mDate + CDate(Left$(sValue, 2) & ":" & Right$(sValue, 2))
        If mStartTime <> CDate(0) Then
            If mStopTime < mStartTime Then
                mDuration = Int((mStopTime - mStartTime) * 24 * 60 + 0.1) + 1440
            Else
                mDuration = Int((mStopTime - mStartTime) * 24 * 60 + 0.1)
            End If
        End If
        FindPlusCode
    End If
    TestValid
End Property
Public Property Let Channel(sValue As String)
    sValue = Normal(sValue)
    mChannel = Val(sValue)
    FindPlusCode
    TestValid
End Property
Public Property Let PlusCode(sValue As String)
    sValue = Normal(sValue)
    mPlusCode = Val(sValue)
    FindDetails
    TestValid
End Property

Public Property Let Recorded(sValue As Integer)
    mRecorded = IIf(sValue = vbChecked, True, False)
End Property
Public Property Let Radio(sValue As Integer)
    mRadio = IIf(sValue = vbChecked, True, False)
    TestValid
End Property
Public Property Let Daily(sValue As Integer)
    mDaily = IIf(sValue = vbChecked, True, False)
    TestValid
End Property
Public Property Let MonFri(sValue As Integer)
    mMonFri = IIf(sValue = vbChecked, True, False)
    TestValid
End Property
Public Property Let Weekly(sValue As Integer)
    mWeekly = IIf(sValue = vbChecked, True, False)
    TestValid
End Property

Public Property Get month() As String
    If mMonth <> 0 Then
        month = Format$(mMonth, "00")
    Else
        month = "--"
    End If
End Property
Public Property Get day() As String
    If mDay <> 0 Then
        day = Format$(mDay, "00")
    Else
        day = "--"
    End If
End Property

Public Property Get StartTime() As String
    If mStartTime <> CDate(0) Then
        StartTime = Format$(mStartTime, "HHMM")
    Else
        StartTime = "----"
    End If
End Property
Public Property Get StopTime() As String
    If mStopTime <> CDate(0) Then
        StopTime = Format$(mStopTime, "HHMM")
    Else
        StopTime = "----"
    End If
End Property
Public Property Get Channel() As String
    If mChannel <> 0 Then
        Channel = Format$(mChannel, "00")
    Else
        Channel = "--"
    End If
End Property
Public Property Get PlusCode() As String
    If mPlusCode <> 0 Then
        PlusCode = CStr(mPlusCode) & String$(8 - Len(CStr(mPlusCode)), "-")
    Else
        PlusCode = "--------"
    End If
End Property
Public Property Get Valid() As Boolean
    Valid = mValid
End Property
Public Property Get Recorded() As Integer
    Recorded = IIf(mRecorded, vbChecked, vbUnchecked)
End Property
Public Property Get Radio() As Integer
    Radio = IIf(mRadio, vbChecked, vbUnchecked)
End Property
Public Property Get Daily() As Integer
    Daily = IIf(mDaily, vbChecked, vbUnchecked)
End Property
Public Property Get MonFri() As Integer
    MonFri = IIf(mMonFri, vbChecked, vbUnchecked)
End Property
Public Property Get Weekly() As Integer
    Weekly = IIf(mWeekly, vbChecked, vbUnchecked)
End Property

Public Property Get Weekday() As String
    If mDate <> CDate(0) Then
        Weekday = mWeekday
    Else
        Weekday = ""
    End If
End Property

Public Property Get Duration() As String
    If mDuration > 0 And mDuration <= 999 Then
        Duration = Format$(mDuration, "000")
    Else
        Duration = "---"
    End If
End Property
Public Property Let Duration(sValue As String)
    mDuration = Val(sValue)
    If mDuration <> 0 Then
        mStopTime = mStartTime + mDuration / 24 / 60
        FindPlusCode
    End If
End Property
Private Sub TestValid()
    Dim bIsValid As Boolean
    Dim sYear As String
    
    If mStatus <> Ready Then
        Exit Sub
    End If
    
    sYear = CStr(mCurrentYear)
    If mMonth <> 0 Then
'        If mMonth < CurrentMonth Then
'            sYear = CStr(CurrentYear + 1)
'        End If
    End If
        
    bIsValid = True
    If mDay = 0 Then bIsValid = False
    If mMonth = 0 Then bIsValid = False
    If mChannel = 0 Then bIsValid = False
    If mStartTime = CDate(0) Then bIsValid = False
    If mStopTime = CDate(0) Then bIsValid = False
    
    mValid = bIsValid
    WriteFile
End Sub

Public Function Clear()
    mStartTime = CDate(0)
    mStopTime = CDate(0)
    mChannel = 0
    mPlusCode = 0
    mValid = False
    mRecorded = False
    mRadio = False
    mDaily = False
    mMonFri = False
    mWeekly = False
    mWeekday = ""
    mDay = 0
    mMonth = 0
    mDate = CDate(0)
    mYear = 0
    mCurrentYear = 0
    mCurrentMonth = 0
    mDuration = 0
    WriteFile
End Function

Private Sub FindPlusCode()
    Dim out As Variant
    Dim sYear As String
    
    Exit Sub
    sYear = CStr(mCurrentYear)
    If mMonth <> 0 Then
        If mMonth < mCurrentMonth Then
            sYear = CStr(mCurrentYear + 1)
        End If
    End If
    
    If IsDate(mDay & "/" & mMonth & "/" & sYear) Then
        If mChannel <> 0 Then
            If mStartTime <> CDate(0) Then
                If mDuration <> 0 Then
                    mPlusCode = Encode(day, month, CInt(sYear), Channel, StartTime, Duration)
                End If
            End If
        End If
    End If
End Sub

Private Sub FindDetails()
    Dim vDetail As Variant
    
    vDetail = Decode(mCurrentMonth, mCurrentDay, mCurrentYear, mPlusCode)
    
    If Not IsEmpty(vDetail) Then
        day = Format$(vDetail(0), "00")
        month = Format$(vDetail(1), "00")
        ' IF YEAR <50 THEN 20XX ELSE 19XX
        Channel = Format$(vDetail(3), "00")
        StartTime = Format$(vDetail(4), "0000")
        StopTime = Format$(vDetail(6), "0000")
        'Duration = Format$(vDetail(5), "000")
    Else
        mPlusCode = 0
    End If
End Sub

Private Sub Class_Initialize()
    mStatus = Ready
End Sub
