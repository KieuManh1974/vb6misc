Public Module Definition
    Public Enum ParseControl
        pcNormal = 0
        pcIgnore = 1
        pcOmit = 2
    End Enum

    Public Structure ParseTree
        Public SubTree As SubTrees()
        Public Text As String
        Public Index As Integer
        Public Control As Long
        Public TextStart As Long
        Public TextEnd As Long

        Public ErrorMessage As String
    End Structure

    Public Structure SubTrees
        Private oItems As Collection

        Public Sub Add(ByVal oTree As ParseTree)
            oItems.Add(oTree)
        End Sub

        Public ReadOnly Property Item(ByVal iIndex As Long) As ParseTree
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
    End Structure
End Module
