Imports System.DirectoryServices
Imports System.DirectoryServices.AccountManagement
Public Class ADPrincipalContext
    Private _domain As String
    Private _user As String

    Private Manager As String
    Public Shared ADPrincipalContexts As List(Of ADPrincipalContext)

    Public Property Userid As String 'sAMAccountName
    Public Property UserName As String
    Public Property Email As String 'mail
    Public Property Company As String 'Company
    Public Property Title As String 'Title
    Public Property DisplayName As String 'displyaname
    Public Property Surename As String 'sn
    Public Property GivenName As String 'givenname
    Public Property Department As String 'department
    Public Property EmployeeNumber As String 'employeenumber
    Public Property TelephoneNumber As String 'telephonenumber
    Public Property Location As String 'l
    Public Property Country As String 'co
    Public Property ErrorMessage As String = String.Empty

    Public Sub New()

    End Sub

    Public Function GetInfo(ByVal domainusername As String) As Boolean
        Dim myret As Boolean = False
        ADPrincipalContexts = New List(Of ADPrincipalContext)
        Dim myvar = domainusername.Split("\")
        _domain = myvar(0)
        _user = myvar(1)
        Dim ctx As PrincipalContext = New PrincipalContext(ContextType.Domain, _domain)

        Dim usp = UserPrincipal.FindByIdentity(ctx, _user)
        If Not IsNothing(usp) Then
            Dim userinfo As DirectoryEntry = DirectCast(usp.GetUnderlyingObject(), DirectoryEntry)
            Dim myInfo As ADPrincipalContext = New ADPrincipalContext

            myInfo.Userid = domainusername
            myInfo.DisplayName = userinfo.Properties("displayname")(0).ToString
            myInfo.UserName = myInfo.DisplayName
            myInfo.Email = userinfo.Properties("mail")(0).ToString
            If userinfo.Properties("company").Count > 0 Then
                myInfo.Company = userinfo.Properties("company")(0).ToString
            End If
            If userinfo.Properties("title").Count > 0 Then
                myInfo.Title = userinfo.Properties("title")(0).ToString
            End If

            myInfo.Surename = userinfo.Properties("sn")(0).ToString
            myInfo.GivenName = userinfo.Properties("givenname")(0).ToString
            If userinfo.Properties("employeenumber").Count > 0 Then
                myInfo.EmployeeNumber = userinfo.Properties("employeenumber")(0).ToString
            End If
            If userinfo.Properties("telephonenumber").Count > 0 Then               
                myInfo.TelephoneNumber = userinfo.Properties("telephonenumber")(0).ToString
            End If

            If userinfo.Properties("department").Count > 0 Then
                myInfo.Department = userinfo.Properties("department")(0).ToString
            End If


            myInfo.Country = userinfo.Properties("co")(0).ToString
            If userinfo.Properties("l").Count > 0 Then
                myInfo.Location = userinfo.Properties("l")(0).ToString
            End If
            ADPrincipalContexts.Add(myInfo)
            'If userinfo.Properties("manager").Count > 0 Then
            '    Dim mgrDN As String = userinfo.Properties("manager")(0).ToString
            '    Dim mgp As UserPrincipal = GetManager(ctx, mgrDN)
            '    If Not IsNothing(mgp) Then
            '        Dim mgrInfo As DirectoryEntry = DirectCast(mgp.GetUnderlyingObject(), DirectoryEntry)

            '        myInfo = New ADPrincipalContext
            '        Dim myArray() = mgrDN.Split(",")
            '        Dim myDomain() = myArray(4).Split("=")
            '        myInfo.Userid = String.Format("{0}\{1}", myDomain(1).ToUpper, mgrInfo.Properties("sAMAccountName")(0).ToString)

            '        myInfo.DisplayName = mgrInfo.Properties("displayname")(0).ToString
            '        myInfo.UserName = myInfo.DisplayName
            '        myInfo.Email = mgrInfo.Properties("mail")(0).ToString
            '        myInfo.Company = mgrInfo.Properties("company")(0).ToString
            '        myInfo.Title = mgrInfo.Properties("title")(0).ToString
            '        myInfo.Surename = mgrInfo.Properties("sn")(0).ToString
            '        myInfo.GivenName = mgrInfo.Properties("givenname")(0).ToString
            '        If mgrInfo.Properties("employeenumber").Count > 0 Then
            '            myInfo.EmployeeNumber = mgrInfo.Properties("employeenumber")(0).ToString
            '        End If
            '        If mgrInfo.Properties("telephonenumber").Count > 0 Then
            '            myInfo.TelephoneNumber = mgrInfo.Properties("telephonenumber")(0).ToString
            '        End If
            '        myInfo.Department = mgrInfo.Properties("department")(0).ToString

            '        If mgrInfo.Properties("co").Count > 0 Then
            '            myInfo.Country = mgrInfo.Properties("co")(0).ToString
            '        End If


            '        If mgrInfo.Properties("l").Count > 0 Then
            '            myInfo.Location = mgrInfo.Properties("l")(0).ToString
            '        End If

            '        mgrDN = mgrInfo.Properties("manager")(0).ToString
            '        ADPrincipalContexts.Add(myInfo)
            '    End If
            'End If
            
            myret = True
        Else
            ErrorMessage = "Unknown User. Please contact Admin."
        End If
        Return myret
    End Function

    Private Function GetManager(ctx As PrincipalContext, mgrDN As String) As UserPrincipal
        Return UserPrincipal.FindByIdentity(ctx, mgrDN)
    End Function

End Class
