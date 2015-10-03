VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Show 
   BackColor       =   &H00800000&
   Caption         =   "Video Plus"
   ClientHeight    =   4035
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9390
   Icon            =   "Show.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   9390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   8880
      TabIndex        =   118
      Top             =   1440
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   8880
      TabIndex        =   117
      Top             =   1755
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   8880
      TabIndex        =   116
      Top             =   2070
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   8880
      TabIndex        =   115
      Top             =   2385
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   8880
      TabIndex        =   114
      Top             =   2700
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   8880
      TabIndex        =   113
      Top             =   3015
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   8880
      TabIndex        =   112
      Top             =   3330
      Width           =   195
   End
   Begin VB.CheckBox chkRecorded 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   8880
      TabIndex        =   111
      Top             =   3645
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   8280
      TabIndex        =   109
      Top             =   1440
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   8280
      TabIndex        =   108
      Top             =   1755
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   8280
      TabIndex        =   107
      Top             =   2070
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   8280
      TabIndex        =   106
      Top             =   2385
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   8280
      TabIndex        =   105
      Top             =   2700
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   8280
      TabIndex        =   104
      Top             =   3015
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   8280
      TabIndex        =   103
      Top             =   3330
      Width           =   195
   End
   Begin VB.CheckBox chkWeekly 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   8280
      TabIndex        =   102
      Top             =   3645
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   7740
      TabIndex        =   100
      Top             =   1440
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   7740
      TabIndex        =   99
      Top             =   1800
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   7740
      TabIndex        =   98
      Top             =   2070
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   7740
      TabIndex        =   97
      Top             =   2385
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   7740
      TabIndex        =   96
      Top             =   2700
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   7740
      TabIndex        =   95
      Top             =   3015
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   7740
      TabIndex        =   94
      Top             =   3330
      Width           =   195
   End
   Begin VB.CheckBox chkDaily 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   7740
      TabIndex        =   93
      Top             =   3645
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   7
      Left            =   7140
      TabIndex        =   83
      Top             =   3645
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   6
      Left            =   7140
      TabIndex        =   82
      Top             =   3330
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   5
      Left            =   7140
      TabIndex        =   81
      Top             =   3015
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   4
      Left            =   7140
      TabIndex        =   80
      Top             =   2700
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   3
      Left            =   7140
      TabIndex        =   79
      Top             =   2385
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   2
      Left            =   7140
      TabIndex        =   78
      Top             =   2070
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   1
      Left            =   7140
      TabIndex        =   77
      Top             =   1755
      Width           =   195
   End
   Begin VB.CheckBox chkRadio 
      Caption         =   "Check1"
      Height          =   195
      Index           =   0
      Left            =   7140
      TabIndex        =   76
      Top             =   1440
      Width           =   195
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   7
      Left            =   1560
      TabIndex        =   59
      Top             =   3645
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   6
      Left            =   1560
      TabIndex        =   51
      Top             =   3330
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   5
      Left            =   1560
      TabIndex        =   43
      Top             =   3015
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   3
      Left            =   1560
      TabIndex        =   27
      Top             =   2385
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   2
      Left            =   1560
      TabIndex        =   19
      Top             =   2070
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   1755
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtWeekday 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Index           =   4
      Left            =   1560
      TabIndex        =   35
      Top             =   2700
      Width           =   1095
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   4980
      Top             =   315
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   1
      Left            =   2700
      TabIndex        =   12
      Top             =   1755
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   2
      Left            =   2700
      TabIndex        =   20
      Top             =   2070
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   3
      Left            =   2700
      TabIndex        =   28
      Top             =   2385
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   4
      Left            =   2700
      TabIndex        =   36
      Top             =   2700
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   0
      Left            =   3120
      TabIndex        =   5
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   1
      Left            =   3120
      TabIndex        =   13
      Top             =   1755
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   2
      Left            =   3120
      TabIndex        =   21
      Top             =   2070
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   3
      Left            =   3120
      TabIndex        =   29
      Top             =   2385
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   4
      Left            =   3120
      TabIndex        =   37
      Top             =   2700
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   0
      Left            =   3780
      TabIndex        =   6
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   1
      Left            =   3780
      TabIndex        =   14
      Top             =   1755
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   2
      Left            =   3780
      TabIndex        =   22
      Top             =   2070
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   3
      Left            =   3780
      TabIndex        =   30
      Top             =   2385
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   4
      Left            =   3780
      TabIndex        =   38
      Top             =   2700
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   0
      Left            =   4560
      TabIndex        =   7
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   1
      Left            =   4560
      TabIndex        =   15
      Top             =   1755
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   2
      Left            =   4560
      TabIndex        =   23
      Top             =   2070
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   3
      Left            =   4560
      TabIndex        =   31
      Top             =   2385
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   4
      Left            =   4560
      TabIndex        =   39
      Top             =   2700
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   0
      Left            =   5460
      TabIndex        =   8
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   1
      Left            =   5460
      TabIndex        =   16
      Top             =   1755
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   2
      Left            =   5460
      TabIndex        =   24
      Top             =   2070
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   3
      Left            =   5460
      TabIndex        =   32
      Top             =   2385
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   4
      Left            =   5460
      TabIndex        =   40
      Top             =   2700
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   0
      Left            =   6300
      TabIndex        =   9
      Top             =   1440
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   1
      Left            =   6300
      TabIndex        =   17
      Top             =   1755
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   2
      Left            =   6300
      TabIndex        =   25
      Top             =   2070
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   3
      Left            =   6300
      TabIndex        =   33
      Top             =   2385
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   4
      Left            =   6300
      TabIndex        =   41
      Top             =   2700
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   1755
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   18
      Top             =   2070
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   26
      Top             =   2385
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   34
      Top             =   2700
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   5
      Left            =   2700
      TabIndex        =   44
      Top             =   3015
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   5
      Left            =   3120
      TabIndex        =   45
      Top             =   3015
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   5
      Left            =   3780
      TabIndex        =   46
      Top             =   3015
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   5
      Left            =   4560
      TabIndex        =   47
      Top             =   3015
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   5
      Left            =   5460
      TabIndex        =   48
      Top             =   3015
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   5
      Left            =   6300
      TabIndex        =   49
      Top             =   3015
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   42
      Top             =   3015
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   6
      Left            =   2700
      TabIndex        =   52
      Top             =   3330
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   6
      Left            =   3120
      TabIndex        =   53
      Top             =   3330
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   6
      Left            =   3780
      TabIndex        =   54
      Top             =   3330
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   6
      Left            =   4560
      TabIndex        =   55
      Top             =   3330
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   6
      Left            =   5460
      TabIndex        =   56
      Top             =   3330
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   6
      Left            =   6300
      TabIndex        =   57
      Top             =   3330
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   50
      Top             =   3330
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDay 
      Height          =   240
      Index           =   7
      Left            =   2700
      TabIndex        =   60
      Top             =   3645
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtMonth 
      Height          =   240
      Index           =   7
      Left            =   3120
      TabIndex        =   61
      Top             =   3645
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtChannel 
      Height          =   240
      Index           =   7
      Left            =   3780
      TabIndex        =   62
      Top             =   3645
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "00"
      Mask            =   "##"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtStartTime 
      Height          =   240
      Index           =   7
      Left            =   4560
      TabIndex        =   63
      Top             =   3645
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtEndTime 
      Height          =   240
      Index           =   7
      Left            =   5460
      TabIndex        =   64
      Top             =   3645
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0000"
      Mask            =   "####"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtDuration 
      Height          =   240
      Index           =   7
      Left            =   6300
      TabIndex        =   65
      Top             =   3645
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "#99"
      PromptChar      =   "-"
   End
   Begin MSMask.MaskEdBox txtCode 
      Height          =   240
      Index           =   7
      Left            =   480
      TabIndex        =   58
      Top             =   3645
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
      BackColor       =   12582912
      ForeColor       =   16744576
      AllowPrompt     =   -1  'True
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "99999999"
      Mask            =   "#9999999"
      PromptChar      =   "-"
   End
   Begin VB.Label Label15 
      BackColor       =   &H00800000&
      Caption         =   "Rec"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   8880
      TabIndex        =   119
      Top             =   1125
      Width           =   555
   End
   Begin VB.Label Label14 
      BackColor       =   &H00800000&
      Caption         =   "Weekly"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   8280
      TabIndex        =   110
      Top             =   1125
      Width           =   555
   End
   Begin VB.Label Label13 
      BackColor       =   &H00800000&
      Caption         =   "Daily"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   7740
      TabIndex        =   101
      Top             =   1125
      Width           =   555
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   7
      Left            =   120
      TabIndex        =   92
      Top             =   3645
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   6
      Left            =   120
      TabIndex        =   91
      Top             =   3330
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   90
      Top             =   3015
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   89
      Top             =   2700
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   88
      Top             =   2385
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   87
      Top             =   2070
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   86
      Top             =   1755
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdClear 
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   85
      Top             =   1440
      Width           =   255
      ForeColor       =   16744576
      BackColor       =   12582912
      Size            =   "450;423"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label Label12 
      BackColor       =   &H00800000&
      Caption         =   "Radio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   7140
      TabIndex        =   84
      Top             =   1125
      Width           =   555
   End
   Begin VB.Label Label6 
      BackColor       =   &H00800000&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   1560
      TabIndex        =   75
      Top             =   270
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Month"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3120
      TabIndex        =   74
      Top             =   1125
      Width           =   555
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Day"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   2700
      TabIndex        =   73
      Top             =   1125
      Width           =   375
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   585
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H00800000&
      Caption         =   "EndTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   195
      Left            =   5460
      TabIndex        =   72
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label Label10 
      BackColor       =   &H00800000&
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   6300
      TabIndex        =   71
      Top             =   1125
      Width           =   795
   End
   Begin VB.Label Label9 
      BackColor       =   &H00800000&
      Caption         =   "StartTime"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   4560
      TabIndex        =   70
      Top             =   1125
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackColor       =   &H00800000&
      Caption         =   "Channel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   3780
      TabIndex        =   69
      Top             =   1125
      Width           =   735
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      Caption         =   "Weekday"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   1560
      TabIndex        =   68
      Top             =   1125
      Width           =   915
   End
   Begin VB.Label Label4 
      BackColor       =   &H00800000&
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   240
      Left            =   480
      TabIndex        =   67
      Top             =   1125
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   285
      Left            =   120
      TabIndex        =   66
      Top             =   270
      Width           =   1335
   End
End
Attribute VB_Name = "Show"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkDaily_Click(Index As Integer)
    oProgrammes(Index).Daily = chkDaily(Index).Value
    ShowProgramme Index
End Sub

Private Sub chkWeekly_Click(Index As Integer)
    oProgrammes(Index).Weekly = chkWeekly(Index).Value
    ShowProgramme Index
End Sub

Private Sub Form_Load()
    Initialise
    Populate
    Update
End Sub

Private Sub tmrTime_Timer()
    lblDate.Caption = Format$(Now, "DD MMM YYYY")
    lblTime.Caption = Format$(Now, "HH:MM:SS")
    Update
    If CheckProgrammes Then
        Populate
    End If
End Sub

Private Sub Update()
    Dim iProgrammeIndex As Long
    
    For iProgrammeIndex = 0 To MaxSlots
        oProgrammes(iProgrammeIndex).CurrentDay = Val(Format$(Now, "DD"))
        oProgrammes(iProgrammeIndex).CurrentMonth = Val(Format$(Now, "MM"))
        oProgrammes(iProgrammeIndex).CurrentYear = Val(Format$(Now, "YYYY"))
    Next
End Sub

Private Sub txtChannel_GotFocus(Index As Integer)
    txtChannel(Index).SelStart = 0
    txtChannel(Index).SelLength = 2
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
    txtCode(Index).SelStart = 0
    txtCode(Index).SelLength = 8
End Sub

Private Sub txtEndTime_GotFocus(Index As Integer)
    txtEndTime(Index).SelStart = 0
    txtEndTime(Index).SelLength = 4
End Sub

Private Sub txtDay_GotFocus(Index As Integer)
    txtDay(Index).SelStart = 0
    txtDay(Index).SelLength = 2
End Sub

Private Sub txtMonth_GotFocus(Index As Integer)
    txtMonth(Index).SelStart = 0
    txtMonth(Index).SelLength = 2
End Sub

Private Sub txtStartTime_GotFocus(Index As Integer)
    txtStartTime(Index).SelStart = 0
    txtStartTime(Index).SelLength = 4
End Sub

Private Sub cmdClear_Click(Index As Integer)
    oProgrammes(Index).Clear
    ShowProgramme Index
End Sub

Private Sub txtCode_LostFocus(Index As Integer)
    oProgrammes(Index).PlusCode = txtCode(Index).Text
    ShowProgramme Index
    'FindDetails Index
    'TestValid Index
    'WriteFile
End Sub

Private Sub txtDay_LostFocus(Index As Integer)
    Dim sText As String
    
    oProgrammes(Index).day = txtDay(Index).Text
    ShowProgramme Index
    
'    sText = Replace$(txtDay(Index).Text, "-", "")
'    If sText <> "" Then
'        txtDay(Index).Text = Format$(sText, "00")
'    End If
'
'    FindWeekday Index
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub txtMonth_LostFocus(Index As Integer)
    Dim sText As String
    
    oProgrammes(Index).month = txtMonth(Index).Text
    ShowProgramme Index
    
'    sText = Replace$(txtMonth(Index).Text, "-", "")
'    If sText <> "" Then
'        txtMonth(Index).Text = Format$(sText, "00")
'    End If
'
'    FindWeekday Index
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub txtChannel_LostFocus(Index As Integer)
    Dim sText As String
    
    oProgrammes(Index).Channel = txtChannel(Index).Text
    ShowProgramme Index
    
'    sText = Replace$(txtChannel(Index).Text, "-", "")
'    If sText <> "" Then
'        txtChannel(Index).Text = Format$(sText, "00")
'    End If
'
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub txtStartTime_LostFocus(Index As Integer)
    Dim sText As String
    
    oProgrammes(Index).StartTime = txtStartTime(Index).Text
    ShowProgramme Index
    
'    sText = Replace$(txtStartTime(Index).Text, "-", "")
'    If sText <> "" Then
'        If Val(sText) < 25 And Len(sText) < 3 Then
'            txtStartTime(Index).Text = Format$(sText, "00") & "00"
'        Else
'            txtStartTime(Index).Text = Format$(sText, "0000")
'        End If
'    End If
'
'    FindDuration Index
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub txtEndTime_LostFocus(Index As Integer)
    Dim sText As String
    
    oProgrammes(Index).StopTime = txtEndTime(Index).Text
    ShowProgramme Index
    
'    sText = Replace$(txtEndTime(Index).Text, "-", "")
'    If sText <> "" Then
'        If Val(sText) < 25 Then
'            txtEndTime(Index).Text = Format$(sText, "00") & "00"
'        Else
'            txtEndTime(Index).Text = Format$(sText, "0000")
'        End If
'    End If
'
'    FindDuration Index
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub txtDuration_LostFocus(Index As Integer)
    oProgrammes(Index).Duration = txtDuration(Index)
    ShowProgramme Index
    
'    FindEndTime Index
'    FindCode Index
'    TestValid Index
'    WriteFile
End Sub

Private Sub chkRadio_Click(Index As Integer)
    oProgrammes(Index).Radio = chkRadio(Index).Value
    ShowProgramme Index
End Sub

Private Function CurrentDay() As Long
    CurrentDay = CLng(Format$(Now, "DD"))
End Function

Private Function CurrentMonth() As Long
    CurrentMonth = CLng(Format$(Now, "MM"))
End Function

Private Function CurrentYear() As Long
    CurrentYear = CLng(Format$(Now, "YYYY"))
End Function

Private Function CurrentHour() As Long
    CurrentHour = CLng(Format$(Now, "HH"))
End Function

Private Function ConvertTime(sTime As String) As String
    ConvertTime = Left$(sTime, 2) & ":" & Mid$(sTime, 3)
End Function

Private Sub Populate()
    Dim iProgrammeIndex As Long
    
    For iProgrammeIndex = 0 To MaxSlots
        ShowProgramme iProgrammeIndex
    Next
End Sub

Private Sub ShowProgramme(ByVal Index As Integer)
    With oProgrammes(Index)
        txtCode(Index).Text = .PlusCode
        txtDay(Index).Text = .day
        txtMonth(Index).Text = .month
        txtChannel(Index).Text = .Channel
        txtStartTime(Index).Text = .StartTime
        txtEndTime(Index).Text = .StopTime
        txtDuration(Index).Text = .Duration
        txtWeekday(Index).Text = .Weekday
        chkRadio(Index).Value = .Radio
        chkDaily(Index).Value = .Daily
        'chkMonFri(Index).Value = .MonFri
        chkWeekly(Index).Value = .Weekly
    End With
End Sub

Private Sub txtDuration_GotFocus(Index As Integer)
    txtDuration(Index).SelStart = 0
    txtDuration(Index).SelLength = 3
End Sub
