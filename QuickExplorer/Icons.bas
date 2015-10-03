Attribute VB_Name = "Icons"
Option Explicit

Private Type TypeIcon
    cbSize As Long
    picType As PictureTypeConstants
    hIcon As Long
End Type

Private Type CLSID
    id(16) As Byte
End Type

Private Const MAX_PATH = 260
Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

Const SHGFI_ICON = &H100                         '  get icon
Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Const SHGFI_TYPENAME = &H400                     '  get type name
Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Const SHGFI_EXETYPE = &H2000                     '  return exe type
Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Const SHGFI_LARGEICON = &H0                      '  get large icon
Const SHGFI_SMALLICON = &H1                      '  get small icon
Const SHGFI_OPENICON = &H2                       '  get open icon
Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute


' Return a file's icon.
Public Function GetIcon(filename As String, icon_size As Long) As IPictureDisp
    Dim index As Integer
    Dim hIcon As Long
    Dim item_num As Long
    Dim icon_pic As IPictureDisp
    Dim sh_info As SHFILEINFO
    
    SHGetFileInfo filename, 0, sh_info, Len(sh_info), SHGFI_ICON + icon_size
    hIcon = sh_info.hIcon
    Set icon_pic = IconToPicture(hIcon)
    Set GetIcon = icon_pic
End Function

' Convert an icon handle into an IPictureDisp.
Private Function IconToPicture(hIcon As Long) As IPictureDisp
    Dim cls_id As CLSID
    Dim hRes As Long
    Dim new_icon As TypeIcon
    Dim lpUnk As IUnknown
    
    With new_icon
        .cbSize = Len(new_icon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    
    With cls_id
        .id(8) = &HC0
        .id(15) = &H46
    End With
    hRes = OleCreatePictureIndirect(new_icon, cls_id, 1, lpUnk)
    
    If hRes = 0 Then
        Set IconToPicture = lpUnk
    End If
End Function

