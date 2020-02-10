Interface IManager
    Function checkAccess(userid As Object, permissionname As String, Optional params As Hashtable = Nothing) As Boolean
    Function createRole(name As String) As Role
    Function createPermission(name As String) As Permission
    Function add(obj As Object) As Boolean
    Function remove(obj As Object) As Boolean
    Function update(name As String, obj As Object) As Boolean
    Function getRole(name As String) As Role
    Function getRoles() As List(Of Item)
    Function getRolesByUser(userid As Object) As List(Of Role)
    Function getPermission(name As String) As Permission
    Function getPermissions() As List(Of Item)
    Function getPermissionByRole(name As String) As List(Of Permission)
    Function getPermissionByUser(userid As Object) As List(Of Permission)
    Function getRule(name As String) As Rule
    Function getRules() As List(Of Rule)
    Function addChild(parent As Item, child As Item) As Boolean
    Function removeChild(parent As Item, child As Item) As Boolean
    Function removeChildren(parent As Item) As Boolean
    Function hasChild(parent As Item, child As Item) As Boolean
    Function getChildren(name As String) As List(Of Item)
    Function assign(role As Role, userid As Object) As Assignment
    Function revoke(role As Role, userid As Object) As Boolean
    Function revokeAll(userid As Object) As Boolean
    Function getAssignment(rolename As String, userid As Object) As Assignment
    Function getAssignments(userid As Object) As List(Of Assignment)
    Function revokeAll() As Boolean
    Function removeAllPermission() As Boolean
    Function removeAllRoles() As Boolean
    Function removeAllRules() As Boolean
    Function removeAllAssignments() As Boolean




End Interface