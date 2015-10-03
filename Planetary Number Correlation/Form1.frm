VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   432
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   904
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Sub Form_Activate()
'    Const max = 55
'    Dim v(max) As Double
'    Dim d(max) As String
'
'    v(0) = 1.989E+29: d(0) = "sun mass"
'    v(1) = 1.412E+18: d(1) = "sun volume"
'    v(2) = 696000: d(2) = "sun mean radius"
'    v(3) = 1408: d(3) = "sun mean density"
'    v(4) = 274: d(4) = "sun surface gravity"
'    v(5) = 617.7: d(5) = "sun escape velocity"
'    v(6) = 0.00005: d(6) = "sun ellipticity"
'    v(7) = 0.059: d(7) = "sun moment of inertia"
'    v(8) = 3.846E+26: d(8) = "sun luminosity"
'    v(9) = 4300000000#: d(9) = "sun mass cnv rate"
'    v(10) = 0.0001937: d(10) = "sun mean energy prod."
'    v(11) = 609.12: d(11) = "sun sidereal rotation period"
'    v(12) = 7.25: d(12) = "sun obl. ecliptic"
'    v(13) = 7.349E+22: d(13) = "moon mass"
'    v(14) = 21968000000#: d(14) = "moon volume"
'    v(15) = 1737.4: d(15) = "moon radius"
'    v(16) = 3340: d(16) = "moon mean density"
'    v(17) = 1.62: d(17) = "moon surface gravity"
'    v(18) = 2.38: d(18) = "moon escape velocity"
'    v(19) = 0.394: d(19) = "moon moment of inertia"
'    v(20) = 384400#: d(20) = "moon s.m. axis"
'    v(21) = 363300#: d(21) = "moon perigee"
'    v(22) = 405500#: d(22) = "moon apogee"
'    v(23) = 27.3217: d(23) = "moon revolution period"
'    v(24) = 29.53: d(25) = "moon synodic period"
'    v(25) = 1.023: d(26) = "moon mean orbital vel."
'    v(26) = 1.076: d(26) = "moon max orbital vel."
'    v(27) = 0.964: d(27) = "moon min orbital vel."
'    v(28) = 5.145: d(28) = "moon orbit inclination"
'    v(29) = 0.0549: d(29) = "moon orbit ecc."
'    v(30) = 655.728: d(30) = "moon sidereal rotation"
'    v(31) = 6.68: d(31) = "moon equatorial incl."
'    v(32) = 3.8: d(32) = "moon recession rate"
'    v(33) = 5.9736E+24: d(33) = "earth mass"
'    v(34) = 10832100000#: d(34) = "earth volume"
'    v(35) = 6378.1: d(35) = "earth equatorial radius"
'    v(36) = 6356.8: d(36) = "earth polar radius"
'    v(37) = 6371#: d(37) = "earth mean radius"
'    v(38) = 3845: d(38) = "earth core radius"
'    v(39) = 0.00335: d(39) = "earth ellipticity"
'    v(40) = 5515: d(40) = "earth mean density"
'    v(41) = 9.78: d(41) = "earth surface gravity"
'    v(42) = 11.186: d(42) = "earth escape velocity"
'    v(43) = 0.3308: d(43) = "earth moment of intertia"
'    v(44) = 149600000#: d(44) = "earth s.m. axis"
'    v(45) = 365.256: d(45) = "earth siderial orbit per."
'    v(46) = 365.242: d(46) = "earth tropical orbit per."
'    v(47) = 147090000#: d(47) = "earth perihelion"
'    v(48) = 152100000#: d(48) = "earth aphelion"
'    v(49) = 29.78: d(49) = "earth mean orbital vel."
'    v(50) = 30.29: d(50) = "earth max orbital vel."
'    v(51) = 29.29: d(51) = "earth min orbital vel."
'    v(52) = 0.0167: d(52) = "earth orbit ecc."
'    v(53) = 23.9345: d(53) = "earth siderial rot."
'    v(54) = 23.45: d(54) = "earth obliquity to orbit"
'    v(55) = 299792458#: d(55) = "speed of light"
'
'    Open "c:\znumbers.txt" For Output As #1
'
'    Dim x As Integer
'    Dim y As Integer
'    Dim a As Double
'    Dim b As Double
'
'    For x = 0 To max
'        a = v(x)
'        Me.Line (Reduce(a) * 100, 0)-Step(0, 100)
'    Next
'
''    For x = 0 To max
''        For y = 0 To x
''            a = v(x) * v(y)
''            Me.Line (reduce(a) * 1000 - 5500, 0)-Step(0, 100)
''        Next
''    Next
''
''    For x = 0 To max
''        For y = 0 To max
''            If x <> y Then
''                b = v(x) / v(y)
''                Me.Line (reduce(b) * 1000 - 5500, 0)-Step(0, 100)
''
''            End If
''        Next
''    Next
'
'End Sub
'
''Private Function reduce(x As Double) As Double
''    Reduce = x / (10 ^ Int((Log(x) / Log(10))))
''End Function
