Public Class InvalidParamException
    Inherits ApplicationException
    Sub New(errorMessage)
        MyBase.New(errorMessage)
    End Sub
End Class

Public Class InvalidConfigException
    Inherits ApplicationException

    Sub New(errorMessage)
        MyBase.New(errorMessage)
    End Sub
End Class

Public MustInherit Class BaseManager
    Inherits Rule
    Implements IManager

    Public defaultRoles As List(Of Role)

    Public MustOverride Function getItem(name As String) As Item
    Public MustOverride Function getItems(type As TypeEnum) As List(Of Item)
    Public MustOverride Function addItem(item As Item) As Boolean
    Public MustOverride Function addRule(rule As Rule) As Boolean
    Public MustOverride Function removeItem(item As Item) As Boolean
    Public MustOverride Function removeRule(rule As Rule) As Boolean
    Public MustOverride Function updateItem(name As String, item As Item) As Boolean
    Public MustOverride Function updateRule(name As String, rule As Rule) As Boolean

    Public MustOverride Function addChild(parent As Item, child As Item) As Boolean Implements IManager.addChild
    Public MustOverride Function assign(role As Role, userid As Object) As Assignment Implements IManager.assign
    Public MustOverride Function checkAccess(userid As Object, permissionname As String, Optional params As Hashtable = Nothing) As Boolean Implements IManager.checkAccess
    Public MustOverride Function getAssignment(rolename As String, userid As Object) As Assignment Implements IManager.getAssignment
    Public MustOverride Function getAssignments(userid As Object) As List(Of Assignment) Implements IManager.getAssignments
    Public MustOverride Function getChildren(name As String) As List(Of Item) Implements IManager.getChildren
    Public MustOverride Function getPermissionByRole(name As String) As List(Of Permission) Implements IManager.getPermissionByRole
    Public MustOverride Function getPermissionByUser(userid As Object) As List(Of Permission) Implements IManager.getPermissionByUser
    Public MustOverride Function getRolesByUser(userid As Object) As List(Of Role) Implements IManager.getRolesByUser
    Public MustOverride Function getRule(name As String) As Rule Implements IManager.getRule
    Public MustOverride Function getRules() As List(Of Rule) Implements IManager.getRules
    Public MustOverride Function hasChild(parent As Item, child As Item) As Boolean Implements IManager.hasChild
    Public MustOverride Function removeAllAssignments() As Boolean Implements IManager.removeAllAssignments
    Public MustOverride Function removeAllPermission() As Boolean Implements IManager.removeAllPermission
    Public MustOverride Function removeAllRoles() As Boolean Implements IManager.removeAllRoles
    Public MustOverride Function removeAllRules() As Boolean Implements IManager.removeAllRules
    Public MustOverride Function removeChildren(parent As Item) As Boolean Implements IManager.removeChildren
    Public MustOverride Function removeChild(parent As Item, child As Item) As Boolean Implements IManager.removeChild
    Public MustOverride Function revoke(role As Role, userid As Object) As Boolean Implements IManager.revoke
    Public MustOverride Function revokeAll() As Boolean Implements IManager.revokeAll
    Public MustOverride Function revokeAll(userid As Object) As Boolean Implements IManager.revokeAll


    Public Function createPermission(name As String) As Permission Implements IManager.createPermission
        Dim permission = New Permission
        permission.name = name
        Return permission
    End Function

    Public Function createRole(name As String) As Role Implements IManager.createRole
        Dim role = New Role
        role.name = name
        Return role
    End Function

    Public Function add(obj As Object) As Boolean Implements IManager.add
        If TypeOf obj Is Item Then
            Return Me.addItem(obj)
        ElseIf TypeOf obj Is Rule Then
            Return Me.addRule(obj)
        Else
            Throw New InvalidParamException("Adding unsupported object type.")
        End If
    End Function

    Public Function getPermission(name As String) As Permission Implements IManager.getPermission
        Dim item = Me.getItem(name)
        Return (IIf(TypeOf item Is Item AndAlso item.type = TypeEnum.TYPE_PERMISSION, item, Nothing))
    End Function

    Public Function getPermissions() As List(Of Item) Implements IManager.getPermissions
        Return Me.getItems(TypeEnum.TYPE_PERMISSION)
    End Function

    Public Function getRole(name As String) As Role Implements IManager.getRole
        Dim item = Me.getItem(name)
        Return (IIf(TypeOf item Is Item AndAlso item.type = TypeEnum.TYPE_ROLE, item, Nothing))
    End Function

    Public Function getRoles() As List(Of Item) Implements IManager.getRoles
        Return Me.getItems(TypeEnum.TYPE_ROLE)
    End Function

    Public Function remove(obj As Object) As Boolean Implements IManager.remove
        If TypeOf obj Is Item Then
            Return Me.removeItem(obj)
        ElseIf TypeOf obj Is Rule Then
            Return Me.removeRule(obj)
        Else : Throw New InvalidParamException("Removing unsupported object type.")
        End If
    End Function

    Public Function update(name As String, obj As Object) As Boolean Implements IManager.update
        If TypeOf obj Is Item Then
            Return Me.updateItem(name, obj)
        ElseIf TypeOf obj Is Rule Then
            Return Me.updateRule(name, obj)
        Else
            Throw New InvalidParamException("Updating unsupported object type.")
        End If
    End Function

    Public Overrides Function executeRule(userid As Object, Optional item As Item = Nothing, Optional params As Hashtable = Nothing) As Boolean
        If IsNothing(item.ruleName) Then
            Return True
        End If
        If IsNothing(params) Then
            Return True
        End If
        If item.ruleName = "" Then
            Return True
        End If

        Dim rule = Me.getRule(item.ruleName)
        If TypeOf rule Is Rule Then
            Return rule.executeRule(userid, item, params)
        Else
            Throw New InvalidConfigException(String.Format("Rule not found: {0}", item.ruleName))
        End If
    End Function


End Class
