Public Enum TypeEnum
    TYPE_ROLE = 1
    TYPE_PERMISSION = 2
End Enum

Public Class Item
    Public Property type As TypeEnum
    Public Property name As String
    Public Property description As String
    Public Property ruleName As String
    Public Property data As Object
    Public Property createdAt As DateTime
    Public Property updatedAt As DateTime

    Public Sub New()
        createdAt = Now()
        updatedAt = Now()
    End Sub
End Class
