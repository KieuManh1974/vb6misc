VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EZPMsgLevel
   ezpAllMessages = 0
   ezpPartialMessages = 1
   ezpNoMessages = 2
End Enum

Public Event Cancel(ByVal sMsg As String, ByRef bCancel As Boolean)
Public Event PasswordRequest(ByRef sPassword As String, ByVal lMaxPasswordLength As Long, ByVal bConfirm As Boolean, ByRef bCancel As Boolean)
Public Event CommentRequest(ByRef sComment As String, ByRef bCancel As Boolean)
Public Event Progress(ByVal lCount As Long, ByVal sMsg As String)

Private m_tZPOPT As ZPOPT
Private m_sFileName As String
Private m_sFileSpecs() As String
Private m_iCount As Long
Private m_lR As Long

Public Property Get ZipFile() As String
   ZipFile = m_sFileName
End Property

Public Property Let ZipFile(ByVal sFileName As String)
   m_sFileName = sFileName
End Property

Public Property Get BasePath() As String
   BasePath = m_tZPOPT.szRootDir
End Property

Public Property Let BasePath(ByVal sBasePath As String)
   m_tZPOPT.szRootDir = sBasePath
End Property

Public Property Let Encrypt(ByVal bState As Boolean)
   m_tZPOPT.fEncrypt = Abs(bState)
End Property

Public Property Let AddComment(ByVal bState As Boolean)
   m_tZPOPT.fComment = Abs(bState)
End Property

Public Property Let StoreDirectories(ByVal bState As Boolean)
   m_tZPOPT.fNoDirEntries = Abs(Not (bState))
End Property

Public Property Let StoreFolderNames(ByVal bState As Boolean)
   m_tZPOPT.fJunkDir = Abs(Not (bState))
End Property

Public Property Let Update(ByVal bState As Boolean)
   m_tZPOPT.fUpdate = Abs(Not (bState))
End Property

Public Property Let RecurseSubDirs(ByVal bState As Boolean)
   If bState Then
      m_tZPOPT.fRecurse = 2
   Else
      m_tZPOPT.fRecurse = 0
   End If
End Property

Public Sub ClearFileSpecs()
   m_iCount = 0
   Erase m_sFileSpecs()
End Sub

Public Function AddFileSpec(ByVal sSpec As String) As Long
   m_iCount = m_iCount + 1
   ReDim Preserve m_sFileSpecs(1 To m_iCount) As String
   m_sFileSpecs(m_iCount) = sSpec
End Function

Public Sub Zip()
   mZip.VBZip Me, m_tZPOPT, m_sFileSpecs(), m_iCount
End Sub

Private Sub Class_Initialize()
   StoreDirectories = False
   StoreFolderNames = False
   RecurseSubDirs = False
End Sub

 

