Public Interface IController
    ReadOnly Property GetTable As DataTable
    Function loaddata() As Boolean
    Function save() As Boolean
End Interface