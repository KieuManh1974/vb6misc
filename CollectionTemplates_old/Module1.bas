Attribute VB_Name = "Module1"
Option Explicit

Sub main()
'    Dim oList As New clsNode
'    Dim oList2 As clsNode
'    Dim oPTR As clsNode
'
'    Set oList2 = oList.AddNew
'    Set oList2.Value = oPTR
'    oList.AddNew.Value = "ItemB"
'    oList.AddNew.Value = "ItemC"
'    oList.AddNew.Value = "ItemD"
'
'    oList2.AddNew.Value = "ItemA1"
'    oList2.AddNew.Value = "ItemA2"
'
'    Set oPTR = oList.FindLogicalItem("ItemB")

    Dim oKey As clsIKey
    Set oKey = New clsKey
    
    oKey.AddItem "abcd", "thisone"
    oKey.AddItem "hat", "thatone"
    Debug.Print oKey.Item("abcd").Item
    
End Sub
