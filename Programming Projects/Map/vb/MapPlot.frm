VERSION 5.00
Begin VB.Form MapPlot 
   BackColor       =   &H00000000&
   Caption         =   "Map"
   ClientHeight    =   8640
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   6.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   576
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   512
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Left            =   1080
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuRecentre 
         Caption         =   "Recentre"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "25%"
         Index           =   0
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "50%"
         Index           =   1
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "100%"
         Index           =   2
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "200%"
         Index           =   3
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "400%"
         Index           =   4
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "800%"
         Index           =   5
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "1600%"
         Index           =   6
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "3200%"
         Index           =   7
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "Background"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuNames 
         Caption         =   "Names"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Vertical"
      End
      Begin VB.Menu mnuCountries 
         Caption         =   "Countries"
         Begin VB.Menu mnuCountry 
            Caption         =   "Country1"
            Index           =   1
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country8"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country9"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country10"
            Index           =   10
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country11"
            Index           =   11
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country12"
            Index           =   12
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country13"
            Index           =   13
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country14"
            Index           =   14
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country15"
            Index           =   15
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country16"
            Index           =   16
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country17"
            Index           =   17
         End
         Begin VB.Menu mnuCountry 
            Caption         =   "Country18"
            Index           =   18
         End
      End
   End
End
Attribute VB_Name = "MapPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Type details_disc
    name As String
    lat As Integer
    long As Integer
End Type

Private non_highlight As Long
Private colours As Variant

Private search_pos As Long
Private loc_ub As Long

Private map_names As Variant

Private map_hcentre As Single
Private map_scale As Single
Private map_vcentre As Single

Private curx As Long
Private cury As Long

Private country_colour As New Collection
Private country As New Collection

Private Enum status_types
    ready_to_stop = 0
    ready_to_search = -1
End Enum

Private search_status As status_types

Private msBasePath As String

Private Sub Form_Activate()
    msBasePath = App.Path & "\"
    DoEvents
    DoEvents
    non_highlight = RGB(0, 0, 128)
    colours = Array(vbRed, vbGreen, vbYellow)
    InitialiseCountries
    InitialiseData
    Status = ready_to_search
End Sub

Private Sub InitialiseCountries()
    Dim fso As New FileSystemObject
    Dim oFile As File
    Dim dot_pos As Long
    Dim country_index As Long
    
    country_index = 1
    
    For Each oFile In fso.GetFolder(msBasePath).Files
        dot_pos = InStrRev(oFile.name, ".")
        If LCase$(Mid$(oFile.name, dot_pos + 1)) = "bin" Then
            country.Add Left$(oFile.name, dot_pos - 1)
            If country_index <= mnuCountry.Count Then
                With mnuCountry(country_index)
                    .Caption = Left$(oFile.name, dot_pos - 1)
                    .Checked = True
                    .Visible = True
                End With
            End If
            country_index = country_index + 1
        End If

    Next
End Sub

Private Sub InitialiseData()
    map_hcentre = 1.5 * 60
    map_scale = 1
    map_vcentre = 51.2 * 60

    Set country_colour = New Collection
    
    country_colour.Add RGB(0, 0, 128), "unitedkingdom"
    country_colour.Add RGB(0, 0, 128), "france"
    country_colour.Add RGB(0, 0, 128), "spain"
    country_colour.Add RGB(0, 0, 128), "belguim"
    country_colour.Add RGB(0, 0, 128), "luxembourg"
    country_colour.Add RGB(0, 0, 128), "norway1"
    country_colour.Add RGB(0, 0, 128), "norway2"
    country_colour.Add RGB(0, 0, 128), "netherlands"
    
    country_colour.Add RGB(0, 0, 128), "germany1"
    country_colour.Add RGB(0, 0, 128), "germany2"
    country_colour.Add RGB(0, 0, 128), "germany3"
    
    country_colour.Add RGB(0, 0, 128), "sweeden1"
    country_colour.Add RGB(0, 0, 128), "sweeden2"
End Sub

Private Sub cmdSearch_Click()
    If search_status = ready_to_search Then
        SwitchStatus
        PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
    Else
        SwitchStatus
    End If
End Sub

Private Property Let Status(bStatus As Boolean)
    search_status = bStatus
    If bStatus Then
        cmdSearch.Caption = "Search"
    Else
        cmdSearch.Caption = "Stop"
    End If
End Property

Private Property Get Status() As Boolean
    Status = search_status
End Property

Private Sub SwitchStatus()
    Status = Not Status
End Sub

Private Sub mnuBackground_Click()
    mnuBackground.Checked = Not mnuBackground.Checked
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
End Sub

Private Sub mnuCountry_Click(Index As Integer)
    If search_status Then
    End If
    mnuCountry(Index).Checked = Not mnuCountry(Index).Checked
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
End Sub

Private Sub mnuNames_Click()
    mnuNames.Checked = Not mnuNames.Checked
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
End Sub

Private Sub mnuVertical_Click()
    mnuVertical.Checked = Not mnuVertical.Checked
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
End Sub

Private Sub mnuRecentre_Click()
    Dim myheight As Long
    Dim mywidth As Long
    
    myheight = Me.ScaleHeight
    mywidth = Me.ScaleWidth
    
    map_hcentre = map_hcentre + (curx - mywidth / 2) / map_scale
    map_vcentre = map_vcentre + (-cury + myheight / 2) / map_scale
    
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked, mnuBackground.Checked
End Sub

Private Sub mnuZoom_Click(Index As Integer)
    Dim myheight As Long
    Dim mywidth As Long
    
    myheight = Me.ScaleHeight
    mywidth = Me.ScaleWidth
    
    map_hcentre = map_hcentre + (curx - mywidth / 2) / map_scale
    map_vcentre = map_vcentre + (-cury + myheight / 2) / map_scale
    
    map_scale = 2 ^ (Index - 2)
    
    Status = ready_to_stop
    PlotLocationsOption UCase$(txtCriteria.Text), mnuNames.Checked, mnuVertical.Checked
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        curx = X
        cury = Y
        Me.PopupMenu mnuPopup
    End If
End Sub


Private Sub PlotLocationsOption(criteria As String, with_names As Boolean, Optional vertical As Boolean, Optional background As Boolean = True)
    If with_names Then
        PlotLocationsWithNames criteria, vertical
    Else
        PlotLocations criteria, background
    End If
End Sub

Private Sub PlotLocations(criteria As String, Optional background As Boolean = True)
    Dim fso As New FileSystemObject
    Dim oFile As File
    Dim myhalfwidth As Long
    Dim myhalfheight As Long
    Dim bin_path As String
    Dim location_count As Long
    Dim temp_location As details_disc
    Dim temp_lat As Long
    Dim temp_long As Long
    Dim country_index As Long
    Dim dot_pos As Long
    
    Dim mehdc As Long
    Dim oParseTree As ParseTree
    Dim expression_index As Long
    Dim colour As Long
    
    Dim sExtentsPath As String
    
    Dim MinLat As Integer
    Dim MaxLat As Integer
    Dim MinLong As Integer
    Dim MaxLong As Integer
    
    Dim MinLatScreen As Long
    Dim MaxLatScreen As Long
    Dim MinLongScreen As Long
    Dim MaxLongScreen As Long
    
    Dim Displayable As Boolean
    Dim DisplayIt As Boolean
                
    Dim loc_lat As Integer
    Dim loc_long As Integer
    Dim loc_length As Byte
    Dim loc_name As String
    
    Me.MousePointer = vbHourglass
    
    'Plot location
    mehdc = Me.hdc
    
    Stream.Text = criteria
    Set oParseTree = New ParseTree
    
    If Evaluator.Parse(oParseTree) Then
        Me.Cls
        country_index = 1
        location_count = 0
        myhalfheight = Me.ScaleHeight \ 2
        myhalfwidth = Me.ScaleWidth \ 2
        For Each oFile In fso.GetFolder(msBasePath).Files
            dot_pos = InStrRev(oFile.name, ".")
            If (LCase$(Mid$(oFile.name, dot_pos + 1)) = "bin") Then
                sExtentsPath = msBasePath & "\" & Left$(oFile.name, dot_pos - 1) & ".ext"
                Open sExtentsPath For Binary As #1
                    Get #1, , MinLat
                    Get #1, , MaxLat
                    Get #1, , MinLong
                    Get #1, , MaxLong
                Close #1
                MinLatScreen = (-MinLat + map_vcentre) * map_scale + myhalfheight
                MaxLatScreen = (-MaxLat + map_vcentre) * map_scale + myhalfheight
                MinLongScreen = (MinLong - map_hcentre) * map_scale + myhalfwidth
                MaxLongScreen = (MaxLong - map_hcentre) * map_scale + myhalfwidth
                
                Displayable = True
                
                If MinLatScreen < 0 Then
                    Displayable = False
                ElseIf MaxLongScreen < 0 Then
                    Displayable = False
                ElseIf MaxLatScreen > (myhalfheight * 2) Then
                    Displayable = False
                ElseIf MinLongScreen > (myhalfwidth * 2) Then
                    Displayable = False
                End If
                
                If country_index <= mnuCountry.Count Then
                    DisplayIt = Displayable And mnuCountry(country_index).Checked
                Else
                    DisplayIt = Displayable
                End If
                
                If DisplayIt Then
                    Open oFile.Path For Binary As #1
                    While Not EOF(1)
                        Get #1, , loc_lat
                        Get #1, , loc_long
                        Get #1, , loc_length
                        loc_name = String$(loc_length, Chr$(0))
                        Get #1, , loc_name
                        temp_long = (loc_long - map_hcentre) * map_scale + myhalfwidth
                        temp_lat = (map_vcentre - loc_lat) * map_scale + myhalfheight
                        colour = GetPixel(ByVal mehdc, ByVal temp_long, ByVal temp_lat)
                        If colour <> -1 Then
                            'ResetText
                            For expression_index = 1 To oParseTree.Index
                                If EvalExpression(loc_name, oParseTree(expression_index)) Then
                                    SetPixelV ByVal mehdc, ByVal temp_long, ByVal temp_lat, ByVal colours((expression_index - 1))
                                    'SetPixelV ByVal mehdc, ByVal temp_long, ByVal temp_lat, vbRed
                                Else
                                    If background Then
                                        If colour = 0 Then
                                            SetPixelV ByVal mehdc, ByVal temp_long, ByVal temp_lat, ByVal non_highlight
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    Wend
                    Close #1
                End If
                country_index = country_index + 1
            End If
            DoEvents
            DoEvents
            If Status = ready_to_search Then
                Exit For
            End If
        Next
    End If
    Me.MousePointer = vbCrosshair
    Status = ready_to_search
End Sub

Private Sub PlotLocationsWithNames(criteria As String, vertical As Boolean)
    Dim fso As New FileSystemObject
    Dim oFile As File
    Dim myhalfwidth As Long
    Dim myhalfheight As Long
    Dim bin_path As String
    Dim location_count As Long
    Dim temp_location As details_disc
    Dim temp_lat As Long
    Dim temp_long As Long
    Dim country_index As Long
    Dim dot_pos As Long
    
    Dim mehdc As Long
    Dim oParseTree As ParseTree
    Dim expression_index As Long
    Dim colour As Long
    
    Dim sExtentsPath As String
    
    Dim MinLat As Integer
    Dim MaxLat As Integer
    Dim MinLong As Integer
    Dim MaxLong As Integer
    
    Dim MinLatScreen As Long
    Dim MaxLatScreen As Long
    Dim MinLongScreen As Long
    Dim MaxLongScreen As Long
    
    Dim Displayable As Boolean
    Dim DisplayIt As Boolean
    
    Dim loc_lat As Integer
    Dim loc_long As Integer
    Dim loc_length As Byte
    Dim loc_name As String
    
    Dim sDisplayName As String
    Dim iLetter As Integer
    
    Me.MousePointer = vbHourglass
    
    'Plot location
    mehdc = Me.hdc
    
    Stream.Text = criteria
    Set oParseTree = New ParseTree
    
    If Evaluator.Parse(oParseTree) Then
        Me.Cls
        country_index = 1
        location_count = 0
        myhalfheight = Me.ScaleHeight \ 2
        myhalfwidth = Me.ScaleWidth \ 2
        For Each oFile In fso.GetFolder(msBasePath).Files
            dot_pos = InStrRev(oFile.name, ".")
            If (LCase$(Mid$(oFile.name, dot_pos + 1)) = "bin") Then
                sExtentsPath = msBasePath & "\" & Left$(oFile.name, dot_pos - 1) & ".ext"
                Open sExtentsPath For Binary As #1
                    Get #1, , MinLat
                    Get #1, , MaxLat
                    Get #1, , MinLong
                    Get #1, , MaxLong
                Close #1
                MinLatScreen = (-MinLat + map_vcentre) * map_scale + myhalfheight
                MaxLatScreen = (-MaxLat + map_vcentre) * map_scale + myhalfheight
                MinLongScreen = (MinLong - map_hcentre) * map_scale + myhalfwidth
                MaxLongScreen = (MaxLong - map_hcentre) * map_scale + myhalfwidth
                
                Displayable = True
                
                If MinLatScreen < 0 Then
                    Displayable = False
                ElseIf MaxLongScreen < 0 Then
                    Displayable = False
                ElseIf MaxLatScreen > (myhalfheight * 2) Then
                    Displayable = False
                ElseIf MinLongScreen > (myhalfwidth * 2) Then
                    Displayable = False
                End If
                
                If country_index <= mnuCountry.Count Then
                    DisplayIt = Displayable And mnuCountry(country_index).Checked
                Else
                    DisplayIt = Displayable
                End If
                
                If DisplayIt Then
                    Open oFile.Path For Binary As #1
                    While Not EOF(1)
                        Get #1, , loc_lat
                        Get #1, , loc_long
                        Get #1, , loc_length
                        loc_name = String$(loc_length, Chr$(0))
                        Get #1, , loc_name
                        temp_long = (loc_long - map_hcentre) * map_scale + myhalfwidth
                        temp_lat = (map_vcentre - loc_lat) * map_scale + myhalfheight
                        colour = GetPixel(ByVal mehdc, ByVal temp_long, ByVal temp_lat)
                        If colour <> -1 Then
                            'ResetText
                            For expression_index = 1 To oParseTree.Index
                                If EvalExpression(loc_name, oParseTree(expression_index)) Then
                                    SetPixelV ByVal mehdc, ByVal temp_long, ByVal temp_lat, ByVal colours((expression_index - 1) \ 2)
                                    sDisplayName = Replace$(Mid$(loc_name, 2, Len(loc_name) - 2), "|", "-")
                                    If Not vertical Then
                                        Me.Line (temp_long, temp_lat)-Step(0, 0)
                                        Me.Print sDisplayName
                                    Else
                                        For iLetter = 1 To Len(sDisplayName)
                                            Me.Line (temp_long, temp_lat + Me.TextHeight(sDisplayName) * (iLetter - 1))-Step(0, 0)
                                            Me.Print Mid$(sDisplayName, iLetter, 1)
                                        Next
                                    End If
                                Else
                                    If colour = 0 Then
                                        SetPixelV ByVal mehdc, ByVal temp_long, ByVal temp_lat, ByVal non_highlight
                                    End If
                                End If
                            Next
                        End If
                    Wend
                    Close #1
                End If
                country_index = country_index + 1
            End If
            DoEvents
            DoEvents
            If Status = ready_to_search Then
                Exit For
            End If
        Next
    End If
    Me.MousePointer = vbCrosshair
    Status = ready_to_search
End Sub



