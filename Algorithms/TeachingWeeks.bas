Attribute VB_Name = "TeachingWeeks"
Option Explicit

Public Function GetTeachingWeeks(ByVal dStartDate As Date, dEndDate As Date) As Long
    Dim lYear As Long
    Dim dWeekPatternStart As Date
    Dim dWeekPatternStartNextYear As Date
    Dim dTerm1 As Date
    Dim dTerm2 As Date
    Dim dTerm3 As Date
    Dim dTerm3End As Date
    Dim lTerm1WeekStart As Long
    Dim lTerm2WeekStart As Long
    Dim lTerm3WeekStart As Long
    Dim lTerm3WeekEnd As Long
    Dim yPattern(53) As Byte
    Dim lIndex As Long
    Dim lStartWeek As Long
    Dim lEndWeek As Long
    
    ' Work out which year the Start Date lies in
    lYear = Year(dStartDate)
    
    dWeekPatternStart = FirstFullWeek("1 Aug " & lYear) ' First Full Week in August
    If dStartDate < dWeekPatternStart Then
        lYear = lYear - 1
        dWeekPatternStart = FirstFullWeek("1 Aug " & lYear) ' First Full Week in August
    End If
        
    dTerm1 = FirstFullWeek("1 Sep " & lYear) + 7 ' Second Full Week in September
    dTerm2 = FirstFullWeek("1 Jan " & lYear + 1) ' First Full Week in January
    If lYear < 2005 Then
        dTerm3 = FirstFullWeek(EasterDay(lYear + 1)) + 7 ' Second Week Following Easter
    Else
        dTerm3 = FirstFullWeek("1 apr " & lYear + 1) + 14 ' Third Full Week in Apr
    End If
    dTerm3End = Array(#6/27/2005#, #7/3/2006#)(lYear - 2004) ' Hard-coded Summer Term end dates
    
    ' Find week numbers
    lTerm1WeekStart = (dTerm1 - dWeekPatternStart) / 7 + 1
    lTerm2WeekStart = (dTerm2 - dWeekPatternStart) / 7 + 1
    lTerm3WeekStart = (dTerm3 - dWeekPatternStart) / 7 + 1
    lTerm3WeekEnd = (dTerm3End - dWeekPatternStart) / 7 + 1
    
    ' Fill in all Terms
    For lIndex = lTerm1WeekStart To lTerm3WeekEnd
        yPattern(lIndex) = 1
    Next
    
    ' Exclude Holidays
    yPattern(lTerm1WeekStart + 6) = 0 ' Autumn Half Term
    yPattern(lTerm2WeekStart - 2) = 0  ' Christmas
    yPattern(lTerm2WeekStart - 1) = 0 ' Christmas
    yPattern(lTerm2WeekStart + 6) = 0 ' Sprint Half Term
    yPattern(lTerm3WeekStart - 2) = 0 ' Easter
    yPattern(lTerm3WeekStart - 1) = 0 ' Easter
    
    ' Find Week Nos for Start and End dates
    lStartWeek = (WeekStart(dStartDate) - dWeekPatternStart) / 7 + 1
    lEndWeek = (WeekStart(dEndDate) - dWeekPatternStart) / 7 + 1
    
    ' Count the number of Teaching Weeks
    For lIndex = lStartWeek To lEndWeek
         GetTeachingWeeks = GetTeachingWeeks + yPattern(lIndex)
    Next
End Function

' This function will return the first full week (starting Monday) following the given date
Public Function FirstFullWeek(dDate As Date) As Date
    FirstFullWeek = dDate - Weekday(dDate, vbTuesday) + 7
End Function

Public Function WeekStart(dDate As Date) As Date
    WeekStart = dDate - Weekday(dDate, vbMonday) + 1
End Function

' This function will provide the date of Easter for the given year
Public Function EasterDay(ByVal dblYear As Double) As Date
    Dim C As Long
    Dim N As Long
    Dim K As Long
    Dim I As Long
    Dim J As Long
    Dim L As Long
    Dim M As Long
    Dim D As Long
    Dim Y As Long
    
    Y = dblYear
    
    C = Int(Y / 100)
    N = Y - 19 * Int(Y / 19)
    K = Int((C - 17) / 25)
    I = C - Int(C / 4) - Int((C - K) / 3) + 19 * N + 15
    I = I - 30 * Int(I / 30)
    I = I - Int(I / 28) * (1 - Int(I / 28) * Int(29 / (I + 1)) * Int((21 - N) / 11))
    J = Y + Int(Y / 4) + I + 2 - C + Int(C / 4)
    J = J - 7 * Int(J / 7)
    L = I - J
    M = 3 + Int((L + 40) / 44)
    D = L + 28 - 31 * Int(M / 4)
    
    EasterDay = CDate(D & "/" & M & "/" & Y)
End Function



