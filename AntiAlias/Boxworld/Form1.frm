VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   1056
      ImageHeight     =   885
      MaskColor       =   12632256
      _Version        =   327682
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const void = ""
Private Const floor = "floor"
Private Const wall = "wall"
Private Const box = "box"
Private Const redbox = "redbox"
Private Const lemon = 5
Private Const man = 6

Private Map(1 To 30, 1 To 30) As String


