<Serializable>
Public MustInherit Class Rule
    Public Property name As String
    Public Property createdAt As DateTime
    Public Property updatedAt As DateTime
    Public Property data As Object
    Public Sub New()
        createdAt = Now()
        updatedAt = Now()
    End Sub
    Public MustOverride Function executeRule(userid As Object, Optional item As Item = Nothing, Optional params As Hashtable = Nothing) As Boolean

End Class
