<ComClass(Class2.ClassId, Class2.InterfaceId, Class2.EventsId)> _
Public Class Class2

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "61DBB2C5-C8F9-4747-9029-58997EE596C6"
    Public Const InterfaceId As String = "5AAC7411-26E3-49CC-8A2D-7A229A725C7D"
    Public Const EventsId As String = "B7C593CD-7131-4538-BDD8-A8D2F5B0A642"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Name As String

    Public Enum Goat
        goat1
        goat2
        goat3
    End Enum

    Public Function GetBod()
        Dim a As Long
        a = 3
    End Function
End Class


