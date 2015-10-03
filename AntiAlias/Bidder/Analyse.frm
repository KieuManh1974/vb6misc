VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnalyse 
   Caption         =   "Bid Analyser"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstBidHistory 
      Height          =   3375
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   5175
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCalculateCloseDate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3840
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txtExpectedAmount2 
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtExpectedAmount1 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   2535
   End
   Begin InetCtlsObjects.Inet oInet 
      Left            =   4200
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "Analyse"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtURL 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "http://offer.ebay.co.uk/ws/eBayISAPI.dll?ViewBids&item=280027072426"
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label6 
      Caption         =   "Bid History"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Close Date"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Expected Amount (Exponential)"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Expected Amount (Linear)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmAnalyse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mdSlope As Double
Private mdIntercept As Double
Private mvExp As Variant

Private Sub cmdAnalyse_Click()
    ExpectedAmounts Analyse(oInet.OpenURL(txtURL.Text), Me)
    txtCloseDate.Text = Format$(mdCloseTime, "DD-MMM-YY HH:MM:SS")
    cmdCalculateCloseDate_Click
End Sub

Private Sub cmdCalculate_Click()
    Dim dX As Double
    
    txtExpectedAmount1.Text = Int(100 * RegLine(mdSlope, mdIntercept, CDbl(CDate(txtDate.Text))) + 0.5) / 100
    
    dX = CDbl(CDate(txtDate.Text))
    'txtExpectedAmount2.Text = mvExp(0) * mvExp(1) ^ (CDbl(CDate(txtDate.Text) - (mdCloseTime - 7)))
    txtExpectedAmount2.Text = mvExp(0) * dX + mvExp(1)
End Sub

Private Sub cmdCalculateCloseDate_Click()
    Dim dX As Double
    
    txtExpectedAmount1.Text = Int(100 * RegLine(mdSlope, mdIntercept, CDbl(CDate(txtCloseDate.Text))) + 0.5) / 100
    
    dX = CDbl(CDate(txtCloseDate.Text))
'    txtExpectedAmount2.Text = mvExp(0) * mvExp(1) ^ (CDbl(CDate(txtCloseDate.Text) - (mdCloseTime - 7)))
    txtExpectedAmount2.Text = mvExp(0) * dX + mvExp(1)
End Sub

Private Sub cmdNow_Click()
    txtDate.Text = Format$(Now, "DD-MMM-YY HH:MM:SS")
End Sub

Private Sub Form_Load()
    InitialiseParser
    txtDate.Text = Format$(Now, "DD-MMM-YY HH:MM:SS")
End Sub

Private Sub ExpectedAmounts(oData As Collection)
    Dim vTimes As Variant
    Dim vAmounts As Variant
    Dim oDatum As clsDatum
    
    vTimes = Array()
    vAmounts = Array()
    
    lstBidHistory.Clear
    
    For Each oDatum In oData
      ReDim Preserve vTimes(UBound(vTimes) + 1)
      ReDim Preserve vAmounts(UBound(vAmounts) + 1)
      
      vTimes(UBound(vTimes)) = CDbl(oDatum.BidDate)
      vAmounts(UBound(vAmounts)) = CDbl(oDatum.BidAmount)
      
      lstBidHistory.AddItem Format$(oDatum.BidDate, "DD-MMM-YY HH:MM:SS") & vbTab & oDatum.BidAmount
    Next
    mdSlope = Slope(vTimes, vAmounts)
    mdIntercept = Intercept(vTimes, vAmounts)
    
    mvExp = NumericalRegression(vTimes, vAmounts, mdCloseTime - 7)
    
    cmdCalculate_Click
    'txtExpectedAmount1.Text = RegLine(Slope(vTimes, vAmounts), Intercept(vTimes, vAmounts), CDbl(CDate(txtDate.Text)))
    'txtExpectedAmount2.Text = RegExp(EBase(vTimes, vAmounts), EMultiplier(vTimes, vAmounts), CDbl(CDate(txtDate.Text)))
End Sub

