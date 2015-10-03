VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBusses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Busses As New Collection

Public Sub AddBus(sRoute As String, sBusStop As String, sDestination As String, dTOA As Date, bTimetabled As Boolean)
    Dim oBus As clsBus
    
    For Each oBus In Busses
        If oBus.Route = sRoute And oBus.BusStop = sBusStop And oBus.Destination = sDestination Then
            If Abs(dTOA - oBus.TOA) < (4 / 1440) Then
                If Not oBus.Timetabled And bTimetabled Then
                ElseIf oBus.Timetabled And Not bTimetabled Then
                    oBus.TOA = dTOA
                    oBus.Timetabled = False
                ElseIf oBus.Timetabled And bTimetabled Then
                    oBus.TOA = dTOA
                ElseIf Not oBus.Timetabled And Not bTimetabled Then
                    oBus.TOA = dTOA
                End If
                
                Exit Sub
            End If
        End If
    Next
    
    Set oBus = New clsBus
    With oBus
        .BusStop = sBusStop
        .Destination = sDestination
        .Route = sRoute
        .TOA = dTOA
        .Timetabled = bTimetabled
    End With
    Busses.Add oBus
End Sub

Public Sub CleanUp(Optional bAll As Boolean)
    Dim oBus As clsBus
    Dim oCon As New Connection
    Dim oQuery As Recordset
    Dim sSQL As String
    Dim lThis As Long
    
    oCon.Open "Provider=Microsoft.Access.OLEDB.10.0;Persist Security Info=False;Data Source=" & App.Path & "\BusLog.mdb;User ID=Admin;Data Provider=Microsoft.Jet.OLEDB.4.0"
    For Each oBus In Busses
        lThis = lThis + 1
        If oBus.TOA < (Now() - (5 / 1440)) Or bAll Then
            oBus.BusStop = Replace(oBus.BusStop, "+", " ")
            oBus.Destination = Replace(oBus.Destination, "&nbsp;", " ")
            sSQL = ""
            sSQL = sSQL & " INSERT INTO"
            sSQL = sSQL & "     Buses"
            sSQL = sSQL & "     (Route,"
            sSQL = sSQL & "     Destination,"
            sSQL = sSQL & "     Stop,"
            sSQL = sSQL & "     TOA,"
            sSQL = sSQL & "     Timetabled)"
            sSQL = sSQL & " VALUES"
            sSQL = sSQL & "     ('" & oBus.Route & "',"
            sSQL = sSQL & "     '" & oBus.Destination & "',"
            sSQL = sSQL & "     '" & oBus.BusStop & "',"
            sSQL = sSQL & "     #" & oBus.TOA & "#,"
            sSQL = sSQL & "     " & oBus.Timetabled & ")"
            
            Set oQuery = New Recordset
            oQuery.Open sSQL, oCon, , adLockOptimistic, adCmdText
            Busses.Remove lThis
            lThis = lThis - 1
        End If
    Next
    oCon.Close
End Sub
