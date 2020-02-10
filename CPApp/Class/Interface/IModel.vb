Public Interface IModel
    ReadOnly Property tablename As String
    ReadOnly Property sortField As String
    Function LoadData(ByRef DS As DataSet) As Boolean    
    Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean
End Interface
