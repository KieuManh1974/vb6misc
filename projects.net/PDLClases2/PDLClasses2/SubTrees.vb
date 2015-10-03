<ComClass(SubTrees.ClassId, SubTrees.InterfaceId, SubTrees.EventsId)> _
Public Class SubTrees

#Region "COM GUIDs"
    ' These  GUIDs provide the COM identity for this class 
    ' and its COM interfaces. If you change them, existing 
    ' clients will no longer be able to access the class.
    Public Const ClassId As String = "408581D6-648B-4B1C-B93B-9A683D808088"
    Public Const InterfaceId As String = "3ADC8815-F602-429B-8753-5572340D1300"
    Public Const EventsId As String = "7E385821-C52A-4B1F-BEA4-A97BC03C3065"
#End Region

    ' A creatable COM class must have a Public Sub New() 
    ' with no parameters, otherwise, the class will not be 
    ' registered in the COM registry and cannot be created 
    ' via CreateObject.
    Public Sub New()
        MyBase.New()
    End Sub

    Private oItems As New Collection()

    Friend Sub Add(ByVal oTree As ParseTree)
        oItems.Add(oTree)
    End Sub

    Default Public ReadOnly Property Item(ByVal iIndex As Long) As ParseTree
        Get
            Return oItems.Item(iIndex)
        End Get
    End Property

    Public ReadOnly Property Count() As Long
        Get
            Return oItems.Count
        End Get
    End Property

    'Public Property Get NewEnum() As IUnknown
    '    Set NewEnum = oItems.[_NewEnum]
    'End Property
End Class


