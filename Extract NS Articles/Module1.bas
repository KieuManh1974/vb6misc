Attribute VB_Name = "mZip"
Option Explicit


''
' ===============================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal path As String) As Long

Private Type ZIPnames
    s(0 To 1023) As String
End Type

Private Type CBChar
    ch(0 To 4096) As Byte
End Type

Private Type CBCh
    ch(0 To 255) As Byte
End Type


Private Type ZIPUSERFUNCTIONS
    lPtrPrint As Long          ' Pointer to application's print routine
    lptrComment As Long        ' Pointer to application's comment routine
    lptrPassword As Long       ' Pointer to application's password routine.
    lptrService As Long        ' callback function designed to be used for     allowing the
End Type

Public Type ZPOPT
  Date As String ' US Date (8 Bytes Long) "12/31/98"?
  szRootDir As String ' Root Directory Pathname (Up To 256 Bytes Long)
  szTempDir As String ' Temp Directory Pathname (Up To 256 Bytes Long)
  fTemp As Long   ' 1 If Temp dir Wanted, Else 0
  fSuffix As Long   ' Include Suffixes (Not Yet Implemented!)
  fEncrypt As Long   ' 1 If Encryption Wanted, Else 0
  fSystem As Long   ' 1 To Include System/Hidden Files, Else 0
  fVolume As Long   ' 1 If Storing Volume Label, Else 0
  fExtra As Long   ' 1 If Excluding Extra Attributes, Else 0
  fNoDirEntries As Long   ' 1 If Ignoring Directory Entries, Else 0
  fExcludeDate As Long   ' 1 If Excluding Files Earlier Than Specified Date,   Else 0
  fIncludeDate As Long   ' 1 If Including Files Earlier Than Specified Date,   Else 0
  fVerbose As Long   ' 1 If Full Messages Wanted, Else 0
  fQuiet As Long   ' 1 If Minimum Messages Wanted, Else 0
  fCRLF_LF As Long   ' 1 If Translate CR/LF To LF, Else 0
  fLF_CRLF As Long   ' 1 If Translate LF To CR/LF, Else 0
  fJunkDir As Long   ' 1 If Junking Directory Names, Else 0
  fGrow As Long   ' 1 If Allow Appending To Zip File, Else 0
  fForce As Long   ' 1 If Making Entries Using DOS File Names, Else 0
  fMove As Long   ' 1 If Deleting Files Added Or Updated, Else 0
  fDeleteEntries As Long   ' 1 If Files Passed Have To Be Deleted, Else 0
  fUpdate As Long   ' 1 If Updating Zip File-Overwrite Only If Newer,   Else 0
  fFreshen As Long   ' 1 If Freshing Zip File-Overwrite Only, Else 0
  fJunkSFX As Long   ' 1 If Junking SFX Prefix, Else 0
  fLatestTime As Long   ' 1 If Setting Zip File Time To Time Of Latest File   In Archive, Else 0
  fComment As Long   ' 1 If Putting Comment In Zip File, Else 0
  fOffsets As Long   ' 1 If Updating Archive Offsets For SFX Files, Else 0
  fPrivilege As Long   ' 1 If Not Saving Privileges, Else 0
  fEncryption As Long   ' Read Only Property!!!
  fRecurse As Long   ' 1 (-r), 2 (-R) If Recursing Into Sub-Directories,   Else 0
  fRepair As Long   ' 1 = Fix Archive, 2 = Try Harder To Fix, Else 0
  flevel As Byte   ' Compression Level - 0 = Stored 6 = Default 9 = Max
End Type

Private Declare Function ZpInit Lib "vbzip10.dll" (ByRef tUserFn As ZIPUSERFUNCTIONS) As Long ' Set Zip Callbacks
Private Declare Function ZpSetOptions Lib "vbzip10.dll" (ByRef tOpts As ZPOPT) As Long ' Set Zip options
Private Declare Function ZpArchive Lib "vbzip10.dll" (ByVal argc As Long, ByVal funame As String, ByRef argv As ZIPnames) As Long ' Real zipping action

Public Sub VBZip(oZip As clsCZip, zpOptions As ZPOPT, sFileSpecs() As String, iFileCount As Long)
    Dim zufUser As ZIPUSERFUNCTIONS
    Dim lR As Long
    Dim lFileTypesIndex As Long
    Dim sZipFile As String
    Dim znNames As ZIPnames

    SetDllDirectory App.path
    
    If Not Len(Trim$(oZip.BasePath)) = 0 Then
        ChDir oZip.BasePath
    End If
    
    ' Set address of callback functions
    With zufUser
        .lPtrPrint = plAddressOf(AddressOf ZipPrintCallback)
        .lptrPassword = plAddressOf(AddressOf ZipPasswordCallback)
        .lptrComment = plAddressOf(AddressOf ZipCommentCallback)
        .lptrService = plAddressOf(AddressOf ZipServiceCallback)
    End With
    
    ZpInit zufUser
    ZpSetOptions zpOptions
    
    ' Go for it!
    For lFileTypesIndex = 1 To iFileCount
        znNames.s(lFileTypesIndex - 1) = sFileSpecs(lFileTypesIndex)
    Next lFileTypesIndex
    znNames.s(lFileTypesIndex) = vbNullChar

    ZpArchive iFileCount, oZip.ZipFile, znNames
End Sub

Private Function ZipServiceCallback(ByRef mname As CBChar, ByVal x As Long) As Long
End Function

Private Function ZipPrintCallback(ByRef fname As CBChar, ByVal x As Long) As Long
End Function

Private Function ZipCommentCallback(ByRef comm As CBChar) As Integer
End Function

Private Function ZipPasswordCallback(ByRef pwd As CBCh, ByVal maxPasswordLength As Long, ByRef s2 As CBCh, ByRef Name As CBCh) As Integer
End Function

Private Function plAddressOf(ByVal lPtr As Long) As Long
   plAddressOf = lPtr
End Function



