Attribute VB_Name = "Module1"
Option Explicit

Public Function AgeAtStartOfQual(dQualStartDate As Date, dDateOfBirth As Date) As Integer
    Dim dNearestJulyDate As Date
    Dim iMonth As Integer
    Dim iYear As Integer
    Dim iNearestJulyDateYear As Integer
    Dim dBirthDate As Date
    Dim iYearToUse As Integer
    Dim iBirthYear As Integer
    
    ' Find nearest 31 July before the Qual Date
    iMonth = Format$(dQualStartDate, "MM")
    iYear = Format$(dQualStartDate, "YYYY")
    If iMonth <= 7 Then
       iYearToUse = iYear - 1
    Else
       iYearToUse = iYear
    End If
    dNearestJulyDate = CDate("31 Jul " & iYearToUse)
    
    ' Find age on this particular 31 July
    dBirthDate = CDate(Format$(dDateOfBirth, "DD/MM/") & iYearToUse)
    
    iBirthYear = Format$(dDateOfBirth, "YYYY")
    AgeAtStartOfQual = iYearToUse - iBirthYear
    If dNearestJulyDate < dBirthDate Then
        AgeAtStartOfQual = AgeAtStartOfQual - 1
    End If
End Function

Public Function Age(dDate As Date, dDateOfBirth As Date) As Integer
    Dim lAge As Long
    Dim dBirthday As Date
    
    Age = Year(dDate) - Year(dDateOfBirth)
    dBirthday = DateSerial(Year(dDate), Month(dDateOfBirth), Day(dDateOfBirth))
    If dBirthday > dDate Then
        Age = Age - 1
    End If
End Function

