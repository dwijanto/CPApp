Public Class Role
    Inherits Item
    Public Overloads Property type As TypeEnum = TypeEnum.TYPE_ROLE
    Public Sub New()
        MyBase.New()
        MyBase.type = TypeEnum.TYPE_ROLE
    End Sub
End Class
