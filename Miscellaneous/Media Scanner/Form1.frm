VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Media Scanner"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2865
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   2865
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstMediaType 
      Height          =   645
      ItemData        =   "Form1.frx":0000
      Left            =   120
      List            =   "Form1.frx":000A
      TabIndex        =   3
      Top             =   945
      Width           =   915
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   330
      Left            =   1800
      TabIndex        =   2
      Top             =   1395
      Width           =   975
   End
   Begin VB.TextBox txtIndex 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Text            =   "0"
      Top             =   135
      Width           =   915
   End
   Begin VB.ListBox lstDrive 
      Height          =   645
      ItemData        =   "Form1.frx":0024
      Left            =   120
      List            =   "Form1.frx":0034
      TabIndex        =   0
      Top             =   135
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private oFS As New FileSystemObject
Private oCon As New Connection
Private oRS As New Recordset

Private Sub cmdScan_Click()
    Dim oDelete As Recordset
    
    cmdScan.Enabled = False
    oCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & App.Path & "\MediaScanner.mdb;"
    oRS.Open "SELECT * FROM tblMedia", oCon, , adLockOptimistic, adCmdText
    
    Set oDelete = New Recordset
    oDelete.Open "DELETE FROM tblMedia WHERE Index = " & txtIndex.Text & " AND MediaType ='" & lstMediaType.List(lstMediaType.ListIndex) & "'", oCon, , adLockOptimistic, adCmdText
    
    Select Case lstDrive.List(lstDrive.ListIndex)
        Case "C Drive"
            ScanFolder oFS.GetFolder("C:\Media")
        Case "D Drive"
            ScanFolder oFS.GetFolder("D:\Media")
        Case "I Drive"
            ScanFolder oFS.GetFolder("I:\")
        Case "Disc"
            ScanFolder oFS.GetFolder("F:")
    End Select
    
    oRS.Close
    oCon.Close
    cmdScan.Enabled = True
End Sub

Private Sub ScanFolder(oFolder As Folder)
    Dim oFile As File
    Dim oSubFolder As Folder
    Dim sExtension As String
    Dim iDotPos As Integer
    
    For Each oFile In oFolder.Files
        iDotPos = InStrRev(oFile.Name, ".")
        If iDotPos > 0 Then
            sExtension = UCase$(Mid$(oFile.Name, iDotPos + 1))
            
            If Len(sExtension) < 5 Then
                oRS.AddNew
                oRS!Index = txtIndex.Text
                oRS!MediaType = lstMediaType.List(lstMediaType.ListIndex)
                oRS!FileName = Left$(oFile.Name, iDotPos - 1)
                oRS!Extension = sExtension
                oRS!Location = oFolder.Path
                oRS!Size = oFile.Size
                oRS.Update
            End If
        End If
    Next
    
    For Each oSubFolder In oFolder.SubFolders
        ScanFolder oSubFolder
    Next
End Sub

