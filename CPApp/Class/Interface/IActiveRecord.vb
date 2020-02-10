Interface IActiveRecord
    Function findByCondition(Optional condition As Object = Nothing, Optional other As Object = Nothing) As Object
    Function deleteAll(Optional condition As Object = Nothing) As Integer
    Function delete(Optional ByVal condition As Object = Nothing)
    Function Find()
    Function findOne(Optional ByVal condition As Object = Nothing)
    Function findAll(Optional ByVal condition As Object = Nothing, Optional other As Object = Nothing)
    Function insert(Optional ByVal condition As Object = Nothing, Optional other As Object = Nothing)

End Interface