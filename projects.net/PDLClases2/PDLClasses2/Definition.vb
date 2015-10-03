Public Module Definition
    Public Enum ParseControl
        pcNormal = 0
        pcIgnore = 1
        pcOmit = 2
    End Enum

    Public Structure ParseTree
        Public SubTree As SubTrees
        Public Text As String
        Public Index As Integer
        Public Control As Long
        Public TextStart As Long
        Public TextEnd As Long

        Public ErrorMessage As String
    End Structure
End Module
