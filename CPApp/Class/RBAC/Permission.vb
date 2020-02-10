Public Class Permission
    Inherits Item
    Public Overloads Property type As TypeEnum = TypeEnum.TYPE_PERMISSION
    Public Sub New()
        MyBase.New()
        MyBase.type = TypeEnum.TYPE_PERMISSION
    End Sub
End Class