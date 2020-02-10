Interface IIdentity
    Function findIdentity(id As Object) As Object
    Function findIdentityByAccessToken(token As Object, Optional type As Object = Nothing) As Object
    Function getAuthKey() As String
    Function getId() As Object
    Function validateAuthKey(authkey As String) As Boolean
End Interface
