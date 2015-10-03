VERSION 5.00
Begin VB.UserControl PictureChanger 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   ScaleHeight     =   1560
   ScaleWidth      =   1755
   ToolboxBitmap   =   "PicChg.ctx":0000
   Begin VB.PictureBox picHidden 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "PictureChanger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Indicates how much we have copied so far.
Private m_PercentDisplayed As Integer

' List of change styles.
Public Enum ChangeStyleValues
    chg_Slats               ' Bring new image in, in slats
    chg_SlideLeftToRight    ' Slide the new image over the old one from left to right.
    chg_ShrinkGrow           ' Shrink old image and grow new image.
    chg_SlideLeftThenRight   'Slide old image left and then slide new image right
    chg_SlideOldLeftDisplayingNew  'Slide old image left displaying new image
    chg_CompressOldExpandNew   'Compress Old image and then Grow New image
    chg_WindowShade  'Brings in new image like a window shade.
    chg_Confused 'a nonsense interactive change that has to "Kick Bytes" to work.
    chg_GridAndFill 'Load new image in grid and fill
    chg_GrowNewOverOld 'Grow new image over old
End Enum

' Default property values.
Private Const m_def_ChangeStyle = chg_Slats
Private Const m_def_Steps = 100

' Property variables.
Private m_ChangeStyle As ChangeStyleValues
Private m_Steps As Integer
' <describe the new style here>
'Brings new image in, in Slats
'Submitted by Neil Crosby - ncrosby@swbell.net
Private Sub ChangeSlats()
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
Dim i As Integer
Dim n As Integer
    ' These values don't change.
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    'dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    src_wid = image_wid / 8
    image_hgt = picHidden.ScaleHeight
For n = 1 To 50  'm_Steps
Phase1:     For i = 1 To 8 Step 2
        ' Get the source coordinates.
        src_x = (src_wid * i) - src_wid 'image_wid - src_wid
        dst_x = (src_wid * i) - src_wid
        ' Get the destination coordinates.
        dst_wid = src_wid + 10

        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
    Next i
Next n
'=====================
For n = 1 To m_Steps
Phase2:     For i = 1 To 8 Step 2
        ' Get the source coordinates.
        src_x = (src_wid * i)
        dst_x = (src_wid * i)
        ' Get the destination coordinates.
        dst_wid = src_wid + 10
        dst_hgt = Int(image_hgt / m_Steps) * n
        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
    Next i
Next n

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub
' <describe the new style here>
'Shrinks the old image and Grows the new image
'Submitted by Neil Crosby - ncrosby@swbell.net
Private Sub ChangeShrinkGrow()
Dim picHold As New StdPicture
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
    ' These values don't change.
    src_y = 0
    dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    image_hgt = picHidden.ScaleHeight
    '===================
   Set picHold = UserControl.Picture
    On Error GoTo BlankPicture
    For step_number = m_Steps To 1 Step -1
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_hgt = image_hgt * (step_number / m_Steps)
        
        src_x = image_wid - src_wid
        src_y = image_hgt - src_hgt
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        ' Copy this part of the image.
        UserControl.Picture = LoadPicture()
        UserControl.PaintPicture picHold, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number
    '======================
        Set UserControl.Picture = LoadPicture()

BlankPicture:     For step_number = 1 To m_Steps
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_hgt = image_hgt * (step_number / m_Steps)
        
        src_x = image_wid - src_wid
        src_y = image_hgt - src_hgt
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub
' <describe the new style here>
'Slides the old image Left and then Slides the new image Right
'Submitted by Neil Crosby - ncrosby@swbell.net
Private Sub ChangeSlideLeftThenRight()
Dim picHold As New StdPicture
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
    ' These values don't change.
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    '===================
   Set picHold = UserControl.Picture
    On Error GoTo BlankPicture
    For step_number = m_Steps To 1 Step -1
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = image_wid - src_wid
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        ' Copy this part of the image.
        UserControl.Picture = LoadPicture()
        UserControl.PaintPicture picHold, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number
    '======================
        Set UserControl.Picture = LoadPicture()

BlankPicture:     For step_number = 1 To m_Steps
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = image_wid - src_wid
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub
Private Sub ChangeSlideOldLeftDisplayingNew()
' <describe the new style here>
'Slides the old image Left displaying the new image Right
'Submitted by Neil Crosby - ncrosby@swbell.net
Dim picHold As New StdPicture
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
    ' These values don't change.
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    '===================
   Set picHold = UserControl.Picture
    On Error GoTo BlankPicture 'No image in UserControl.Picture
    For step_number = m_Steps To 1 Step -1
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = image_wid - src_wid
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        ' Copy this part of the image.
        UserControl.Picture = picHidden.Picture
        UserControl.PaintPicture picHold, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number
BlankPicture:
    
End Sub
Private Sub ChangeCompressOldExpandNew()
' <describe the new style here>
'Compress old image and then Expand the new image
'Submitted by Neil Crosby - ncrosby@swbell.net

Dim picHold As New StdPicture
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
    ' These values don't change.
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    'dst_x = 0  (image_wid - src_wid)/2
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    '===================
   Set picHold = UserControl.Picture
    On Error GoTo BlankPicture
    For step_number = m_Steps To 1 Step -1
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = (image_wid - src_wid) / 2
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        dst_x = (image_wid - src_wid) / 2
        ' Copy this part of the image.
        UserControl.Picture = LoadPicture()
        UserControl.PaintPicture picHold, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number
    '======================
        Set UserControl.Picture = LoadPicture()

BlankPicture:     For step_number = 1 To m_Steps
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = (image_wid - src_wid) / 2
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        dst_x = (image_wid - src_wid) / 2
        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture

End Sub
' <Grid and Fill the new image in.>
'Submitted by Neil Crosby - ncrosby@swbell.net
Private Sub ChangeGridAndFill()
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
Dim x_pos As Single
Dim y_pos As Single
Dim Mym_Steps As Integer
    ' These values don't change.
    image_wid = picHidden.ScaleWidth
    image_hgt = picHidden.ScaleHeight
    Mym_Steps = 15
    '===================
        ' Set the Grid individual image size.
        src_wid = image_wid / Mym_Steps
        src_hgt = image_hgt / Mym_Steps
        dst_wid = src_wid
        dst_hgt = src_hgt
    
    '======================
    UserControl.Picture = LoadPicture()


Step1: For step_number = 1 To Mym_Steps Step 2
            'Set initial x positions
            src_x = (src_wid * Mym_Steps) - src_wid
            dst_x = (dst_wid * Mym_Steps) - dst_wid
        For x_pos = src_x To 0 Step -2 * src_wid
            src_x = x_pos
            dst_x = x_pos
        ' Set  y positions
            src_y = (src_hgt * step_number) - src_hgt
            dst_y = (dst_hgt * step_number) - dst_hgt
        ' Copy this part of the image.
        
            UserControl.PaintPicture picHidden, _
                dst_x, dst_y, dst_wid, dst_hgt, _
                src_x, src_y, src_wid, src_hgt
        
            DoEvents
        Next x_pos
    Next step_number
    '=========================
Step2: For step_number = 1 To Mym_Steps Step 2
            'Set initial x positions
            src_x = (src_wid * Mym_Steps) - src_wid
            dst_x = (dst_wid * Mym_Steps) - dst_wid
        For x_pos = src_x To 1 Step -1 * src_wid
            src_x = x_pos
            dst_x = x_pos
        ' Set  y positions
            src_y = (src_hgt * step_number) - src_hgt
            dst_y = (dst_hgt * step_number) - dst_hgt
        ' Copy this part of the image.
        
            UserControl.PaintPicture picHidden, _
                dst_x, dst_y, dst_wid, dst_hgt, _
                src_x, src_y, src_wid, src_hgt
        
            DoEvents
        Next x_pos
    Next step_number
    '=========================
Step3:  For step_number = Mym_Steps To 1 Step -1
            'Set initial x positions
            src_x = (src_wid * Mym_Steps) - src_wid
            dst_x = (dst_wid * Mym_Steps) - dst_wid
        For x_pos = src_x To 0 Step -1 * src_wid
            src_x = x_pos
            dst_x = x_pos
        ' Set  y positions
            src_y = (src_hgt * step_number) - src_hgt
            dst_y = (dst_hgt * step_number) - dst_hgt
        ' Copy this part of the image.
        
            UserControl.PaintPicture picHidden, _
                dst_x, dst_y, dst_wid, dst_hgt, _
                src_x, src_y, src_wid, src_hgt
        
            DoEvents
        Next x_pos
    Next step_number
    '======================
    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub

Private Sub ChangeConfused()
'Submitted by Neil Crosby
'Tried putting in sleep declaration rather than use message box,
'but lost print statements.
'==================
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
Dim midPtx As Single
Dim midPty As Single

    ' These values don't change.
    src_hgt = picHidden.ScaleHeight / 2
    dst_hgt = src_hgt
    src_wid = picHidden.ScaleWidth / 2
    dst_wid = src_wid
    midPtx = picHidden.ScaleWidth / 2
    midPty = picHidden.ScaleHeight / 2
'Copy 1Q to 4Q
src_x = 0
src_y = 0
dst_x = midPtx
dst_y = midPty

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 2Q to 3Q
src_x = midPtx
src_y = 0
dst_x = 0
dst_y = midPty

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 3Q to 2Q
src_x = 0
src_y = midPty
dst_x = midPtx
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 4Q to 1Q
src_x = midPtx
src_y = midPty
dst_x = 0
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents

UserControl.CurrentY = 10
UserControl.Print "Something"
UserControl.CurrentY = 300
UserControl.Print "is WRONG!"
UserControl.CurrentY = 600
UserControl.Print "Will Try"
UserControl.CurrentY = 900
UserControl.Print "AGAIN!!"
MsgBox "Continue?"
'========================2nd Round
'Copy 1Q to 3Q
src_x = 0
src_y = 0
dst_x = 0
dst_y = midPty

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 2Q to 1Q
src_x = midPtx
src_y = 0
dst_x = 0
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 3Q to 1Q
src_x = 0
src_y = midPty
dst_x = 0
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 4Q to 2Q
src_x = midPtx
src_y = midPty
dst_x = midPtx
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents

UserControl.CurrentY = 10
UserControl.Print "Still"
UserControl.CurrentY = 300
UserControl.Print "WRONG!"
UserControl.CurrentY = 600
UserControl.Print "Guess I"
UserControl.CurrentY = 900
UserControl.Print "will have"
UserControl.CurrentY = 1200
UserControl.Print "to KICK"
UserControl.CurrentY = 1500
UserControl.Print "some"
UserControl.CurrentY = 1800
UserControl.Print "BYTES!!"
MsgBox "Continue?"
'======================3rd  Round
'Copy 1Q to 1Q
src_x = 0
src_y = 0
dst_x = 0
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 2Q to 2Q
src_x = midPtx
src_y = 0
dst_x = midPtx
dst_y = 0

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 3Q to 3Q
src_x = 0
src_y = midPty
dst_x = 0
dst_y = midPty

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
'Copy 4Q to 4Q
src_x = midPtx
src_y = midPty
dst_x = midPtx
dst_y = midPty

        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents

UserControl.CurrentY = 10
UserControl.Print "There -"
UserControl.CurrentY = 300
UserControl.Print "NOW it"
UserControl.CurrentY = 600
UserControl.Print "Looks"
UserControl.CurrentY = 900
UserControl.Print "RIGHT."
MsgBox "Finally got it right!"
'======================
    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture

End Sub
Private Sub ChangeGrowNewOverOld()
'submitted by Neil Crosby ncrosby@swbell.net
Dim picHold As New StdPicture
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single

   Set picHold = UserControl.Picture
   
   image_wid = UserControl.ScaleWidth
    image_hgt = UserControl.ScaleHeight
    On Error GoTo BlankPicture
    For step_number = m_Steps To 1 Step -1
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_hgt = image_hgt * (step_number / m_Steps)
        
        src_x = image_wid / 2 - src_wid / 2
        src_y = image_hgt / 2 - src_hgt / 2
        
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        dst_x = src_x
        dst_y = src_y
        ' Copy this part of the image.
        UserControl.Picture = LoadPicture()
        UserControl.PaintPicture picHold, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        
        DoEvents
    Next step_number
    '======================
        Set UserControl.Picture = LoadPicture()
    
    'Grow New ==========================
BlankPicture:    image_wid = picHidden.ScaleWidth
    image_hgt = picHidden.ScaleHeight
    
    For step_number = 1 To m_Steps
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_hgt = image_hgt * (step_number / m_Steps)
        
        src_x = image_wid / 2 - src_wid / 2
        src_y = image_hgt / 2 - src_hgt / 2
        

        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = src_hgt
        dst_x = src_x
        dst_y = src_y

        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt

        DoEvents
    Next step_number

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture

End Sub
' Return the change style.
Public Property Get ChangeStyle() As ChangeStyleValues
    ChangeStyle = m_ChangeStyle
End Property

' Set the change style.
Public Property Let ChangeStyle(ByVal New_ChangeStyle As ChangeStyleValues)
    m_ChangeStyle = New_ChangeStyle
    PropertyChanged "ChangeStyle"
End Property

' Return the control's current picture.
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

' Display a new picture using the selected change style.
Public Property Set Picture(ByVal New_Picture As Picture)
    ' Copy the picture into the hidden PictureBox.
    Set picHidden.Picture = New_Picture

    ' Display the new picture in the right way.
    Select Case m_ChangeStyle
        Case chg_Slats
            ChangeSlats

        Case chg_SlideLeftToRight
            ChangeSlideLeftToRight

        Case chg_ShrinkGrow
            ChangeShrinkGrow
            
        Case chg_SlideLeftThenRight
            ChangeSlideLeftThenRight
            
        Case chg_SlideOldLeftDisplayingNew
            ChangeSlideOldLeftDisplayingNew
            
        Case chg_CompressOldExpandNew
            ChangeCompressOldExpandNew
            
        Case chg_WindowShade
            ChangeWindowShade
        
        Case chg_Confused
            ChangeConfused
            
        Case chg_GridAndFill
            ChangeGridAndFill
        Case chg_GrowNewOverOld
            ChangeGrowNewOverOld
        
    End Select

    Set UserControl.Picture = picHidden.Picture
    PropertyChanged "Picture"
End Property
' Copy the right part of the new image over the left part
' of the old image.
Private Sub ChangeSlideLeftToRight()
Dim step_number As Integer
Dim image_wid As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single

    ' These values don't change.
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth

    For step_number = 1 To m_Steps
        ' Get the source coordinates.
        src_wid = image_wid * (step_number / m_Steps)
        src_x = image_wid - src_wid

        ' Get the destination coordinates.
        dst_wid = src_wid

        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt

        DoEvents
    Next step_number

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub
' <describe the new style here>
'Brings new image in like a window shade.
'Submitted by Neil Crosby - ncrosby@swbell.net
Private Sub ChangeWindowShade()
Dim step_number As Integer
Dim image_wid As Single
Dim image_hgt As Single
Dim src_x As Single
Dim src_y As Single
Dim src_wid As Single
Dim src_hgt As Single
Dim dst_x As Single
Dim dst_y As Single
Dim dst_wid As Single
Dim dst_hgt As Single
Dim i As Integer
Dim n As Integer
    ' These values don't change.
    src_x = 0
    src_y = 0
    src_hgt = picHidden.ScaleHeight
    dst_x = 0
    dst_y = 0
    dst_hgt = src_hgt
    image_wid = picHidden.ScaleWidth
    src_wid = image_wid
    image_hgt = picHidden.ScaleHeight
For n = 1 To m_Steps
     For i = 1 To 8 Step 2
        ' Get the destination coordinates.
        dst_wid = src_wid
        dst_hgt = Int(image_hgt / m_Steps) * n
        ' Copy this part of the image.
        UserControl.PaintPicture picHidden.Picture, _
            dst_x, dst_y, dst_wid, dst_hgt, _
            src_x, src_y, src_wid, src_hgt
        DoEvents
    Next i
Next n

    ' Set the final image.
    Set UserControl.Picture = picHidden.Picture
End Sub

' Initialize the control's default properties.
Private Sub UserControl_InitProperties()
    m_ChangeStyle = m_def_ChangeStyle
    m_Steps = m_def_Steps
End Sub

' Load saved property values.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set UserControl.Picture = PropBag.ReadProperty("Picture", Nothing)
    m_ChangeStyle = PropBag.ReadProperty("ChangeStyle", m_def_ChangeStyle)
    m_Steps = PropBag.ReadProperty("Steps", m_def_Steps)
End Sub

' Save property values.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Picture", Picture, Nothing
    PropBag.WriteProperty "ChangeStyle", m_ChangeStyle, m_def_ChangeStyle
    PropBag.WriteProperty "Steps", m_Steps, m_def_Steps
End Sub

' Return the number of steps the change should use.
Public Property Get Steps() As Integer
    Steps = m_Steps
End Property

' Set the number of steps the change should use.
Public Property Let Steps(ByVal New_Steps As Integer)
    m_Steps = New_Steps
    PropertyChanged "Steps"
End Property
