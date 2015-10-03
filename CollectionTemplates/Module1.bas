Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim oKey As New clsKey
    Dim oNode As New clsNode
    
    oNode.TextKey = "fatter"
    oNode.AddNew , "alpha"
    oNode.AddNew , "beta"
    oNode.AddNew , "alphanumeric"
 
    oNode.RemoveTextKey "beta"
    'Debug.Print oNode.ItemPhysical(1).TextKey
    oNode.Keys.EnumerateWordList
End Sub
