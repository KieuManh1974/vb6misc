Attribute VB_Name = "DeltaCompare"
Option Explicit

Public Function Compare(ByVal sFile1 As String, ByVal sFile2 As String) As Collection
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim lFPos1 As Long
    Dim lFPos2 As Long
    Dim lBPos1 As Long
    Dim lBPos2 As Long
    Dim oInstruction As clsInstruction
    
    Dim vData As Variant
    
    Set Compare = New Collection
    
    OpenFile sFile1, 1
    OpenFile sFile2, 2
    
    lPos1 = 1
    lPos2 = 1
    
Point1:
    If FetchChar(1, lPos1) <> FetchChar(2, lPos2) Then
        GoTo Point3
    End If
    
Point2:
    lPos1 = lPos1 + 1
    lPos2 = lPos2 + 1
    
    If lPos1 > mlLen(1) Then
        If lPos2 <= mlLen(2) Then
            Set oInstruction = CreateInstruction2(1, lPos2, mlLen(2) - lPos2 + 1)
            FetchBytes 2, lPos2, mlLen(2) - lPos2 + 1, vData
            oInstruction.ByteData = vData
            Compare.Add oInstruction
'            Debug.Print "Inserted @ " & lPos2 & " length " & lFile2Len - lPos2 + 1 & " data " & FetchChar(2, lPos2, lFile2Len - lPos2 + 1)
        End If
        Close #1
        Close #2
        Exit Function
    End If
    
    If lPos2 > mlLen(2) Then
        If lPos1 <= mlLen(1) Then
            Set oInstruction = CreateInstruction2(0, lPos2, mlLen(1) - lPos1 + 1)
            Compare.Add oInstruction
'            Debug.Print "Deleted @ " & lPos2 & " length " & lFile1Len - lPos1 + 1
        End If
        Close #1
        Close #2
        Exit Function
    End If
    GoTo Point1
    
Point3:
    lFPos1 = lPos1 + 1
    lFPos2 = lPos2 + 1
    
Point4:
    If FetchChar(1, lFPos1) = FetchChar(2, lFPos2) Then
        If lFPos1 > mlLen(1) Then
            lFPos1 = mlLen(1) + 1
        End If
        Set oInstruction = CreateInstruction2(0, lPos2, lFPos1 - lPos1)
        Compare.Add oInstruction
'        Debug.Print "Deleted @ " & lPos2 & " length " & lFPos1 - lPos1
        If lFPos2 > mlLen(2) Then
            lFPos2 = mlLen(2) + 1
        End If
        Set oInstruction = CreateInstruction2(1, lPos2, lFPos2 - lPos2)
        FetchBytes 2, lPos2, lFPos2 - lPos2, vData
        oInstruction.ByteData = vData
        Compare.Add oInstruction
'        Debug.Print "Inserted @ " & lPos2 & " length " & lFPos2 - lPos2 & " data " & FetchChar(2, lPos2, lFPos2 - lPos2)
        lPos1 = lFPos1
        lPos2 = lFPos2
        GoTo Point2
    End If
    
    lBPos1 = lFPos1 - 1
    lBPos2 = lFPos2 - 1
    
Point5:
    If FetchChar(1, lFPos1) = FetchChar(2, lBPos2) Then
        Set oInstruction = CreateInstruction2(0, lPos2, lFPos1 - lPos1)
        Compare.Add oInstruction
'        Debug.Print "Deleted @ " & lPos2 & " length " & lFPos1 - lPos1
        If lBPos2 > lPos2 Then
            Set oInstruction = CreateInstruction2(1, lPos2, lBPos2 - lPos2)
            FetchBytes 2, lPos2, lBPos2 - lPos2, vData
            oInstruction.ByteData = vData
            Compare.Add oInstruction
'            Debug.Print "Inserted @ " & lPos2 & " length " & lBPos2 - lPos2 & " data " & FetchChar(2, lPos2, lBPos2 - lPos2)
        End If
        lPos1 = lFPos1
        lPos2 = lBPos2
        GoTo Point2
    ElseIf FetchChar(2, lFPos2) = FetchChar(1, lBPos1) Then
        If lBPos1 > lPos1 Then
            Set oInstruction = CreateInstruction2(0, lPos2, lBPos1 - lPos1)
            Compare.Add oInstruction
'            Debug.Print "Deleted @ " & lPos2 & " length " & lBPos1 - lPos1
        End If
        Set oInstruction = CreateInstruction2(1, lPos2, lFPos2 - lPos2)
        FetchBytes 2, lPos2, lFPos2 - lPos2, vData
        oInstruction.ByteData = vData
        Compare.Add oInstruction
'        Debug.Print "Inserted @ " & lPos2 & " length " & lFPos2 - lPos2 & " data " & FetchChar(2, lPos2, lFPos2 - lPos2)
        lPos1 = lBPos1
        lPos2 = lFPos2
        GoTo Point2
    End If
    
    If lBPos1 = lPos1 Then
        lFPos1 = lFPos1 + 1
        lFPos2 = lFPos2 + 1
        GoTo Point4
    End If
    
    lBPos1 = lBPos1 - 1
    lBPos2 = lBPos2 - 1
    
    GoTo Point5
End Function
