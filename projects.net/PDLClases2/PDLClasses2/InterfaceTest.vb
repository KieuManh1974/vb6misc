<ComClass(InterfaceTest.ClassId, InterfaceTest.InterfaceId, InterfaceTest.EventsId)> _
Public Class InterfaceTest

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "4A7FA9E6-FEDC-4AF6-A60F-ED83910C3B3E"
    Public Const InterfaceId As String = "6F15BE3A-562F-409B-BAC5-E0B848F0C250"
    Public Const EventsId As String = "ACD8FE4C-25CC-4A16-982A-E188C5780999"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Public Sub Start(ByVal x As ParseControl, ByVal errmsg As String, ByVal ParamArray pInitParam() As Object)
        Dim a As Int16
        a = 10
    End Sub
End Class


