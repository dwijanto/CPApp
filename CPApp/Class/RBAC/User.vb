Public Class User
    Public Shared Property IdentityClass As Object
    Private Shared _id As Object = Nothing
    Private Shared _isguest As Boolean = True

    Public Shared Property identity As Object = Nothing

    Public Shared Function getIsGuest() As Boolean
        Return IsNothing(identity)
    End Function

    Public Shared Function can(permissionname As String, Optional params As Hashtable = Nothing)
        If IsNothing(_id) Then
            Throw New InvalidValueException("User must perform Login function.")
        End If
        Return getAuthManager.checkAccess(_id, permissionname, params)
    End Function

    Protected Shared Function getAuthManager() As BaseManager
        Return New DbManager
    End Function

    Public Shared Function getId() As Object
        Return _id
    End Function

    Public Shared Function getIdentity() As Object
        Return identity
    End Function

    Public Shared Sub setIdentity(ByVal value As Object)
        If TypeOf value Is IIdentity Then
            identity = value
        Else
            identity = value
            Throw New InvalidValueException("The Identity object must implement IdentityInterface.")
        End If
    End Sub

    Public Shared Function login(Identity As Object) As Boolean
        _id = Identity.getId
        Return Not getIsGuest()
    End Function

    Public Function loginByAcesToken(token As String, Optional type As Object = Nothing) As Object
        Dim obj = IdentityClass
        _identity = obj.findIdentityByAccessToken(token, type)
        If Not IsNothing(_identity) AndAlso login(_identity) Then
            Return _identity
        Else
            Return Nothing
        End If
    End Function

    Public Shared Function logout() As Boolean
        _identity = Nothing
        Return getIsGuest()
    End Function


End Class


Class InvalidValueException
    Inherits ApplicationException

    Public Sub New(errorMessage As String)
        MyBase.New(errorMessage)
    End Sub

End Class