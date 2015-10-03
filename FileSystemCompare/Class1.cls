VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Snapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type FileInfo
    Name As String
    UpdatedDate As Date
    FileType As Boolean
    Level As Byte
End Type

Private FileStruct0() As FileInfo
Private FileStruct1() As FileInfo

Private mlPosition As Long
Private mlScanIndex As Long
Private oFS As New FileSystemObject

Public Sub Scan(ByVal lScanIndex)
    mlScanIndex = lScanIndex - 1
    mlPosition = 0
    
    Select Case mlScanIndex
        Case 0
            Erase FileStruct0
            ScanFolder0 oFS.GetFolder("c:\"), 0
        Case 1
            Erase FileStruct1
            ScanFolder1 oFS.GetFolder("c:\"), 0
    End Select
End Sub

Private Sub ScanFolder0(oFolder As Folder, ByVal lLevel As Long)
    Dim oSubFolder As Folder
    Dim oFile As File
    
    On Error Resume Next
    
    If UCase$(oFolder.Path) = "C:\MAIN" Then
        Exit Sub
    End If
    
    For Each oSubFolder In oFolder.SubFolders
        If Err.Number = 70 Then
            Err.Clear
            Exit Sub
        End If
        ReDim Preserve FileStruct0(mlPosition) As FileInfo
        With FileStruct0(mlPosition)
            .FileType = True
            .Name = oSubFolder.Name
            .UpdatedDate = oSubFolder.DateLastModified
            .Level = lLevel
        End With
        mlPosition = mlPosition + 1
        ScanFolder0 oSubFolder, lLevel + 1
    Next
    
    For Each oFile In oFolder.Files
        ReDim Preserve FileStruct0(mlPosition) As FileInfo
        With FileStruct0(mlPosition)
            .FileType = False
            .Name = oFile.Name
            .UpdatedDate = oFile.DateLastModified
            .Level = lLevel
        End With
        mlPosition = mlPosition + 1
    Next
End Sub

Private Sub ScanFolder1(oFolder As Folder, ByVal lLevel As Long)
    Dim oSubFolder As Folder
    Dim oFile As File
    
    On Error Resume Next
    
    If UCase$(oFolder.Path) = "C:\MAIN" Then
        Exit Sub
    End If
    
    For Each oSubFolder In oFolder.SubFolders
        If Err.Number = 70 Then
            Err.Clear
            Exit Sub
        End If
        ReDim Preserve FileStruct1(mlPosition) As FileInfo
        With FileStruct1(mlPosition)
            .FileType = True
            .Name = oSubFolder.Name
            .UpdatedDate = oSubFolder.DateLastModified
            .Level = lLevel
        End With
        mlPosition = mlPosition + 1
        ScanFolder1 oSubFolder, lLevel + 1
    Next
    
    For Each oFile In oFolder.Files
        ReDim Preserve FileStruct1(mlPosition) As FileInfo
        With FileStruct1(mlPosition)
            .FileType = False
            .Name = oFile.Name
            .UpdatedDate = oFile.DateLastModified
            .Level = lLevel
        End With
        mlPosition = mlPosition + 1
    Next
End Sub

Public Sub Compare()
    Dim lPosition0 As Long
    Dim lPosition1 As Long
    Dim lSearch0 As Long
    Dim lSearch1 As Long
    Dim bFinished As Boolean
    Dim sAdded As String
    Dim sRemoved As String
    Dim sModified As String
    Dim bMatchFound As Boolean
    Dim vFilename0 As Variant
    Dim vFilename1 As Variant
    
    vFilename0 = Array()
    vFilename1 = Array()
    
    While Not bFinished
        ReDim Preserve vFilename0(FileStruct0(lPosition0).Level) As Variant
        ReDim Preserve vFilename1(FileStruct1(lPosition1).Level) As Variant
        
        vFilename0(UBound(vFilename0)) = FileStruct0(lPosition0).Name
        vFilename1(UBound(vFilename1)) = FileStruct1(lPosition1).Name
                
        If FileStruct0(lPosition0).Name <> FileStruct1(lPosition1).Name Or FileStruct0(lPosition0).Level <> FileStruct1(lPosition1).Level Or FileStruct0(lPosition0).FileType <> FileStruct1(lPosition1).FileType Then
            bMatchFound = False
            lSearch0 = lPosition0
            lSearch1 = lPosition1
            Do While Not bMatchFound
                If FileStruct0(lSearch0).Level = FileStruct1(lPosition1).Level And FileStruct0(lSearch0).FileType = FileStruct1(lPosition1).FileType And FileStruct0(lSearch0).Name = FileStruct1(lPosition1).Name Then
                    sRemoved = sRemoved & Join(vFilename0, "\") & vbCrLf
                    lPosition0 = lSearch0 + 1
                    lPosition1 = lPosition1 + 1
                    bMatchFound = True
                    Exit Do
                ElseIf FileStruct0(lPosition0).Level = FileStruct1(lSearch1).Level And FileStruct0(lPosition0).FileType = FileStruct1(lSearch1).FileType And FileStruct0(lPosition0).Name = FileStruct1(lSearch1).Name Then
                    sAdded = sAdded & Join(vFilename1, "\") & vbCrLf
                    bMatchFound = True
                    lPosition0 = lPosition0 + 1
                    lPosition1 = lSearch1 + 1
                    Exit Do
                End If
                lSearch0 = lSearch0 + 1
                lSearch1 = lSearch1 + 1
                If lSearch0 > UBound(FileStruct0) Then
                    sRemoved = sRemoved & Join(vFilename0, "\") & vbCrLf
                    sAdded = sAdded & Join(vFilename1, "\") & vbCrLf
                    lPosition0 = lPosition0 + 1
                    lPosition1 = lPosition1 + 1
                    Exit Do
                ElseIf lSearch1 > UBound(FileStruct1) Then
                    sRemoved = sRemoved & Join(vFilename0, "\") & vbCrLf
                    sAdded = sAdded & Join(vFilename1, "\") & vbCrLf
                    lPosition0 = lPosition0 + 1
                    lPosition1 = lPosition1 + 1
                    Exit Do
                ElseIf FileStruct0(lSearch0).Level < FileStruct0(lSearch0 - 1).Level Or FileStruct1(lSearch1).Level < FileStruct1(lSearch1 - 1).Level Then
                    sRemoved = sRemoved & Join(vFilename0, "\") & vbCrLf
                    sAdded = sAdded & Join(vFilename1, "\") & vbCrLf
                    lPosition0 = lPosition0 + 1
                    lPosition1 = lPosition1 + 1
                    Exit Do
                End If
            Loop
        ElseIf FileStruct0(lPosition0).UpdatedDate <> FileStruct1(lPosition1).UpdatedDate Then
            sModified = sModified & Join(vFilename0, "\") & vbCrLf
            lPosition0 = lPosition0 + 1
            lPosition1 = lPosition1 + 1
        Else
            lPosition0 = lPosition0 + 1
            lPosition1 = lPosition1 + 1
        End If
        
        If lPosition0 > UBound(FileStruct0) Then
            bFinished = True
        ElseIf lPosition1 > UBound(FileStruct1) Then
            bFinished = True
        End If
        
    Wend
    
    Debug.Print vbCrLf & "[Added]"
    Debug.Print sAdded
    Debug.Print vbCrLf & "[Removed]"
    Debug.Print sRemoved
    Debug.Print vbCrLf & "[Modified]"
    Debug.Print sModified
    
    Dim oTS As TextStream
    
    Set oTS = oFS.CreateTextFile("c:\temp\folderdiff.txt", True)
    oTS.WriteLine "[Added]"
    oTS.WriteLine sAdded
    oTS.WriteLine "[Removed]"
    oTS.WriteLine sRemoved
    oTS.WriteLine "[Modified]"
    oTS.WriteLine sModified
End Sub

Private Sub Class_Initialize()
    'ReDim FileStruct(1) As FileInfo
End Sub
