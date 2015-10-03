VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Removal
    sParentPath As String
End Type

Private msRemove() As String
Private msConnect() As String
Private msExpanded() As String
Private moFSO As New FileSystemObject
Private moTreeList As New clsTreeList
Private mvDrag As Variant

Private Sub Form_Initialize()
    mvDrag = Array(Nothing, stNone, Nothing, False)
End Sub

Private Sub Form_Load()
    Dim oTree As clsTreeNode
    Dim oSubTree As clsTreeNode
    Dim vChildren As Variant
    
    Set oTree = moTreeList.NewNode
    Set goPlus = LoadPicture(App.Path & "\plus.bmp")
    Set goMinus = LoadPicture(App.Path & "\minus.bmp")
    Set goTicked = LoadPicture(App.Path & "\ticked.bmp")
    Set goUnticked = LoadPicture(App.Path & "\unticked.bmp")
    
    oTree.Position.X = 10
    oTree.Position.Y = 10
    oTree.FilePath = "C:\Projects (Other)\QuickExplorer\Module1.bas"
    
    oTree.DrawTree Me, True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim vNode As Variant
    Dim oOffset As clsPoint
    
    If Button = vbLeftButton Then
        mvDrag = moTreeList.FindNodeAtPosition(X, Y)
        ReDim Preserve mvDrag(3)
        If Not mvDrag(0) Is Nothing Then
            Set oOffset = New clsPoint
            oOffset.X = X - mvDrag(0).Position.X
            oOffset.Y = Y - mvDrag(0).Position.Y
            Set mvDrag(2) = oOffset
            mvDrag(3) = False
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not mvDrag(0) Is Nothing Then
        mvDrag(0).Position.X = X - mvDrag(2).X
        mvDrag(0).Position.Y = Y - mvDrag(2).Y
        Cls
        mvDrag(0).DrawTree Me, True
        mvDrag(3) = True
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mvDrag(3) = False Then 'not moved
            Select Case mvDrag(1)
                Case SelectionType.stExpand
                    mvDrag(0).Expanded = Not mvDrag(0).Expanded
                    mvDrag(0).DrawTree Me, True
                Case SelectionType.stCaption
                Case SelectionType.stVisibility
                Case SelectionType.stIcon
            End Select
        End If
        mvDrag = Array(Nothing, stNone, Nothing, False)
    End If
End Sub
