Imports System.Text
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization
Imports System.IO
Imports System.IO.BinaryWriter
'Imports Npgsql

Public Class DbManager
    Inherits BaseManager
    Implements IDisposable

    'Tables
    'auth_item
    'auth_item_child
    'auth_assignment
    'auth_rule 

    'Private As IDbDataParameter As PostgreSQLDBAdapter
    Private Schemaname As String = "cp"

    Public Property itemTable = String.Format("{0}.auth_item", Schemaname)
    Public Property itemChildTable = String.Format("{0}.auth_item_child", Schemaname)
    Public Property assignmentTable = String.Format("{0}.auth_assignment", Schemaname)
    Public Property ruleTable = String.Format("{0}.auth_rule", Schemaname)

    Protected items As New List(Of Item)
    Protected rules As New List(Of Rule)
    'Protected parents As New Hashtable

    Public Sub New()
        MyBase.New()
        'As IDbDataParameter = PostgreSQLDBAdapter.getInstance
    End Sub

    Public Overrides Function addChild(parent As Item, child As Item) As Boolean
        Dim myret As Boolean = False

        If parent.name = child.name Then
            Throw New InvalidParamException(String.Format("Cannot add {0} as a child of itself.", parent.name))
        End If
        If TypeOf parent Is Permission AndAlso TypeOf child Is Role Then
            Throw New InvalidParamException(String.Format("Cannot add a role as a child of a permission"))
        End If
        Dim sqlstr = String.Format("insert into {0}(parent,child) values(:parent,:child);", itemChildTable)
        
        Dim params(1) As IDbDataParameter

        params(0) = DataAccess.factory.CreateParameter("parent", parent.name)
        params(1) = DataAccess.factory.CreateParameter("child", child.name)

        Dim message As String = String.Empty
        Dim result As Object = Nothing
        'If Not As IDbDataParameter.ExecuteScalar(sqlstr, params, result, message) Then
        '    myret = False
        '    Throw New InvalidParamException(message)
        'End If
        DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)
        myret = True
        Return myret
    End Function

    Public Overrides Function addItem(item As Item) As Boolean
        Dim myret As Boolean = False
        Dim mytime As DateTime = Now
        If IsNothing(item.createdAt) Then
            item.createdAt = mytime
        End If
        If IsNothing(item.updatedAt) Then
            item.updatedAt = mytime
        End If

        'insert to db

        Dim sqlstr = String.Format("insert into {0}(name,type,description,rule_name,data,created_at,updated_at) values(:name,:type,:description,:rulename," &
                                   ":data,:createdat,:updatedat);", itemTable)
        If IsNothing(item.data) Then
            item.data = DBNull.Value
        End If

        Dim params(6) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("name", item.name)
        params(1) = DataAccess.factory.CreateParameter("type", Int(item.type))
        params(2) = DataAccess.factory.CreateParameter("description", item.description)
        params(3) = DataAccess.factory.CreateParameter("rulename", item.ruleName)
        params(4) = DataAccess.factory.CreateParameter("data", item.data)
        params(5) = DataAccess.factory.CreateParameter("createdat", item.createdAt)
        params(6) = DataAccess.factory.CreateParameter("updatedat", item.updatedAt)

        Dim message As String = String.Empty
        Dim result As Object = Nothing
        'If Not As IDbDataParameter.ExecuteScalar(sqlstr, params, result, message) Then
        '    Throw New InvalidParamException(message)
        'Else
        '    Return True
        'End If
        DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)
        myret = True
        Return myret
    End Function

    Public Overrides Function addRule(rule As Rule) As Boolean
        Dim myret As Boolean = False
        Dim mytime As DateTime = Now
        If IsNothing(rule.createdAt) Then
            rule.createdAt = mytime
        End If
        If IsNothing(rule.updatedAt) Then
            rule.updatedAt = mytime
        End If

        'insert to db

        Dim sqlstr = String.Format("insert into {0}(name,data,created_at,updated_at) values(:name,:data,:createdat,:updatedat);", ruleTable)


        Dim params(3) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("name", rule.name)
        params(1) = DataAccess.factory.CreateParameter("createdat", rule.createdAt)
        params(2) = DataAccess.factory.CreateParameter("updatedat", rule.updatedAt)
        params(3) = DataAccess.factory.CreateParameter("data", rule.data)
        Dim message As String = String.Empty
        Dim result As Object = Nothing
        'If Not As IDbDataParameter.ExecuteScalar(sqlstr, params, result, message) Then
        '    Throw New InvalidParamException(message)
        'End If
        DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)
        myret = True
        Return result
    End Function

    Public Overrides Function assign(role As Role, userid As Object) As Assignment
        Dim assignment = New Assignment With {.userid = userid,
                                              .rolename = role.name,
                                              .cretedAt = Now()}

        'insert to db

        Dim sqlstr = String.Format("insert into {0}(user_id,item_name,created_at) values(:userid,:itemname,:createdat);", assignmentTable)
        'Dim params As New List(Of IDataParameter)
        'params.Add(dbadapter1.getParam("userid", assignment.userid, DbType.String))
        'params.Add(dbadapter1.getParam("itemname", assignment.rolename, DbType.String))
        'params.Add(dbadapter1.getParam("createdat", assignment.cretedAt, DbType.DateTime))
        Dim params(2) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("userid", assignment.userid)
        params(1) = DataAccess.factory.CreateParameter("itemname", assignment.rolename)
        params(2) = DataAccess.factory.CreateParameter("createdat", assignment.cretedAt)
        Dim message As String = String.Empty
        Dim result As Object = Nothing
        'If Not dbadapter1.ExecuteScalar(sqlstr, params, result, message) Then
        '    Throw New InvalidParamException(message)
        'End If
        DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)

        Return assignment

    End Function

    Public Overrides Function checkAccess(userid As Object, permissionname As String, Optional params As Hashtable = Nothing) As Boolean
        Dim assigments = Me.getAssignments(userid)
        If assigments.Count > 0 Then
            Return Me.checkAccessRecursive(userid, permissionname, params, assigments)
        Else
            Return False
        End If

    End Function

    Protected Function checkAccessRecursive(userid As Object, itemname As String, params As Hashtable, assigments As List(Of Assignment))
        Dim item = Me.getItem(itemname)
        If IsNothing(item) Then
            Return False
        End If
        'Check RuleName Parameter to be compared inside Rule 
        If Not IsNothing(params) Then
            If Not Me.executeRule(userid, item, params) Then
                Return False
            End If
        End If

        'check assignment
        For Each ass As Assignment In assigments
            If ass.rolename = itemname Then
                Return True
            End If
        Next


        Dim sqlstr = String.Format("select parent from {0} where child = :itemname", itemChildTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("itemname", itemname, DbType.String))

        Dim dbparams(0) As IDbDataParameter
        dbparams(0) = DataAccess.factory.CreateParameter("itemname", itemname)

        'Dim ds As New DataSet
        'If dbadapter1.GetDataset(sqlstr, ds, dbparams) Then
        '    For Each dr As DataRow In ds.Tables(0).Rows
        '        'myparam.Add("role", dr.Item("rolename"))
        '        If Me.checkAccessRecursive(userid, dr.Item("parent"), params, assigments) Then
        '            Return True
        '        End If
        '    Next
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, dbparams)
        If Not IsNothing(ds) Then
            For Each dr As DataRow In ds.Tables(0).Rows               
                If Me.checkAccessRecursive(userid, dr.Item("parent"), params, assigments) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function


    Public Overrides Function getAssignment(rolename As String, userid As Object) As Assignment
        If IsNothing(rolename) Then
            Return Nothing
        End If


        Dim sqlstr = String.Format("select * from {0} where userid = :userid and item_name = :rolename", assignmentTable)

        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("userid", userid, DbType.String))
        'dbparams.Add(dbadapter1.getParam("rolename", rolename, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("userid", userid)
        params(1) = DataAccess.factory.CreateParameter("userid", rolename)

        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Return New Assignment With {.userid = ds.Tables(0).Rows(0).Item("user_id"),
                                    .rolename = ds.Tables(0).Rows(0).Item("rolename"),
                                    .cretedAt = ds.Tables(0).Rows(0).Item("created_at")}

    End Function

    Public Overrides Function getAssignments(userid As Object) As List(Of Assignment)
        If (IsNothing(userid)) Then
            Return Nothing
        End If
        Dim sqlstr As String = String.Empty
        If IsNumeric(userid) Then
            sqlstr = String.Format("select * from {0} where user_id::integer = :userid", assignmentTable)
        Else
            sqlstr = String.Format("select * from {0} where user_id = :userid", assignmentTable)
        End If

        Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("userid", userid, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("userid", userid)


        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If

        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If IsNothing(ds) Then
            Return Nothing
        End If

        Dim assignments = New List(Of Assignment)
        For Each dr As DataRow In ds.Tables(0).Rows
            assignments.Add(New Assignment With {.userid = dr.Item("user_id"),
                                    .rolename = dr.Item("item_name"),
                                    .cretedAt = dr.Item("created_at")})
        Next
        Return assignments

    End Function

    Public Overrides Function getChildren(name As String) As List(Of Item)

        Dim sqlstr = String.Format("select i.* from {0} c left join {1} i on i.name = c.child where c.parent = :parent ", itemChildTable, itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("parent", name, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("parent", name)

        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim children = New List(Of Item)
        For Each dr As DataRow In ds.Tables(0).Rows
            children.Add(populateItem(dr))
        Next
        Return children
    End Function

    Public Overrides Function getItem(name As String) As Item
        If IsNothing(name) Then
            Return Nothing
        End If

        Dim sqlstr = String.Format("select * from {0} where name = :name limit 1", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("name", name, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("name", name)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        End If
        Return populateItem(ds.Tables(0).Rows(0))
    End Function

    Protected Function populateItem(ByVal dr As DataRow)
        Dim obj As Object = Nothing
        Select Case dr.Item("type")
            Case TypeEnum.TYPE_PERMISSION
                obj = New Permission
            Case TypeEnum.TYPE_ROLE
                obj = New Role
        End Select

        With obj
            .name = dr.Item("name")
            .description = "" & dr.Item("description")
            .rulename = "" & dr.Item("rule_name")
            .data = dr.Item("data")
            .createdAt = dr.Item("created_at")
            .updatedat = dr.Item("updated_at")
        End With
        Return obj
    End Function

    Public Overrides Function getItems(type As TypeEnum) As List(Of Item)

        Dim sqlstr = String.Format("select * from {0} where type = :type order by name", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("type", Int(type), DbType.Int16))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("type", Int(type))
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim items = New List(Of Item)
        For Each dr As DataRow In ds.Tables(0).Rows
            items.Add(populateItem(dr))
        Next
        Return items
    End Function

    Public Function getAllItems() As List(Of Item)

        Dim sqlstr = String.Format("select * from {0} order by type,name", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("type", Int(type), DbType.Int16))
       
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim items = New List(Of Item)
        For Each dr As DataRow In ds.Tables(0).Rows
            items.Add(populateItem(dr))
        Next
        Return items
    End Function

    Public Function test(ByRef result As List(Of String))
        getChildrenRecursive("author", getChildrenList(), result)
        Return True
    End Function


    Protected Function getChildrenList() As BindingSource
        Dim parents = New BindingSource
        Dim sqlstr = String.Format("select parent::text,child::text from {0}", itemChildTable)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        ds.Tables(0).TableName = "Parent"
        parents.DataSource = ds.Tables(0)
        Return parents

    End Function

    Protected Sub getChildrenRecursive(ByVal name As String, childrenlist As BindingSource, ByRef result As List(Of String))
        Dim dt = DirectCast(childrenlist.DataSource, DataTable)
        'use dataview
        Dim mydata = dt.AsDataView
        mydata.RowFilter = String.Format("[parent] = '{0}'", name)
        For Each dr As DataRowView In mydata
            'If Not result.ContainsKey(dr.Item("child")) Then
            result.Add(dr.Item("child"))
            getChildrenRecursive(dr.Item("child"), childrenlist, result)
            'End If
        Next
    End Sub
    Public Overrides Function getPermissionByRole(name As String) As List(Of Permission)
        Dim childrenlist = Me.getChildrenList
        Dim result As New List(Of String)
        getChildrenRecursive(name, childrenlist, result)
        If IsNothing(result) Then
            Return Nothing
        End If
        'get record permission within result.toarray
        Dim sb As New StringBuilder
        For Each arr In result
            If sb.Length > 0 Then
                sb.Append(",")
            End If
            sb.Append(String.Format("'{0}'", arr))
        Next

        Dim sqlstr = String.Format("select * from {0} where name in ({1}) and type = 2;", itemTable, sb.ToString)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        'convert to list of permission
        Dim mylist As New List(Of Permission)
        For Each dr In ds.Tables(0).Rows
            mylist.Add(populateItem(dr))
        Next
        Return mylist
    End Function

    Public Overrides Function getPermissionByUser(userid As Object) As List(Of Permission)
        If IsNothing(userid) Then
            Return Nothing
        End If

        Dim directPermission = getDirectPermissionbyUser(userid)
        Dim inheritedPermission = getInheritedPermissionByUser(userid)

        Dim result = directPermission
        For Each p In inheritedPermission
            result.Add(p)
        Next
        Return result
    End Function

    Protected Function getDirectPermissionbyUser(userid As Object) As List(Of Permission)
        Dim sqlstr = String.Format("select b.* from {0} a left join {1} b on b.name = a.item_name  where a.user_id ='{2}' and type = 2;", assignmentTable, itemTable, userid)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        'convert to list of permission
        Dim mylist As New List(Of Permission)
        For Each dr In ds.Tables(0).Rows
            mylist.Add(populateItem(dr))
        Next
        Return mylist

    End Function

    Private Function getInheritedPermissionByUser(userid As Object) As Object
        Dim sqlstr = String.Format("select item_name from {0} where user_id = '{1}'", assignmentTable, userid)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim mylist As New List(Of Object)
        Dim myHashTable As New Hashtable
        For Each role In ds.Tables(0).Rows
            For Each p As Permission In getPermissionByRole(role.item("item_name"))
                If Not myHashTable.ContainsKey(p.name) Then
                    myHashTable.Add(p.name, True)
                    mylist.Add(p)
                End If
            Next
        Next

        Return mylist
    End Function

    Public Overrides Function getRolesByUser(userid As Object) As List(Of Role)
        If IsNothing(userid) Then
            Return Nothing
        End If
        Dim sqlstr = String.Format("select b.* from {0} a left join {1} b on b.name = a.item_name  where a.user_id ='{2}' and type = 1;", assignmentTable, itemTable, userid)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        'convert to list of role
        Dim mylist As New List(Of Role)
        For Each dr In ds.Tables(0).Rows
            mylist.Add(populateItem(dr))
        Next
        Return mylist
    End Function

    Public Overrides Function getRule(name As String) As Rule

        Dim sqlstr = String.Format("select data from {0} where name = :name", ruleTable)
        'Dim sqlstr = String.Format("select * from {0} where rule_name = :name limit 1", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("name", name, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("name", name)
        Dim ds As New DataSet
        Dim result As Object = Nothing
        'If Not dbadapter1.ExecuteScalar(sqlstr, params, result) Then
        '    Return Nothing
        'End If
        result = DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)
        If IsNothing(result) Then
            Return Nothing
        End If
        Return Deserialize(name, result)

    End Function

    Public Overrides Function getRules() As List(Of Rule)
        Dim sqlstr = String.Format("select * from {0}", ruleTable)
        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim result As New List(Of Rule)
        For Each dr In ds.Tables(0).Rows
            result.Add(Deserialize(dr.item("name"), dr.item("data")))
        Next
        Return result
    End Function
    'Public Function getRuleClass(ByVal rulename As String, ByVal obj As Object) As Rule
    '    Serialize(obj)
    '    Return Deserialize(rulename)
    'End Function

    Public Function Serialize(ByVal obj As Object) As Byte()
        Dim bytes() As Byte
        Try
            If File.Exists(String.Format("{0}.dat", obj.name)) Then
                File.Delete(String.Format("{0}.dat", obj.name))
            End If
            'Using mystream As Stream = File.Create(String.Format("{0}.dat", obj.name))
            '    Dim serializer As New BinaryFormatter
            '    serializer.Serialize(mystream, obj)
            '    Serialize = Nothing           
            'End Using
            Using mystream = New FileStream(String.Format("{0}.dat", obj.name), FileMode.Create, FileAccess.Write)
                Dim serializer As New BinaryFormatter
                serializer.Serialize(mystream, obj)
                Serialize = Nothing
            End Using

        Catch ex As Exception
            Debug.Print("error")
        End Try

        Using fs As FileStream = New FileStream(String.Format("{0}.dat", obj.name), FileMode.Open, FileAccess.Read)
            Dim br = New BinaryReader(New BufferedStream(fs))
            bytes = br.ReadBytes(fs.Length)
            br.Close()
            br.Dispose()
            br = Nothing
        End Using
        Using mystream = New FileStream("beforesave.dat", FileMode.Create, FileAccess.Write)
            Dim br = New BinaryWriter(New BufferedStream(mystream))
            br.Write(bytes)
            br.Flush()
            'br.Close()
        End Using
        Return bytes
    End Function

    Protected Function Deserialize(ByVal classname As String, ByVal data As Object) As Rule
        Using mystream = New FileStream(String.Format("{0}.dat", classname), FileMode.Create, FileAccess.Write)
            Dim br = New BinaryWriter(New BufferedStream(mystream))
            br.Write(data)
            br.Flush()
        End Using
        Using fs As New FileStream(String.Format("{0}.dat", classname), FileMode.Open)
            Dim formatter As New BinaryFormatter
            Return DirectCast(formatter.Deserialize(fs), Rule)
        End Using
    End Function


    Public Overrides Function hasChild(parent As Item, child As Item) As Boolean
        If IsNothing(parent) Or IsNothing(child) Then
            Return Nothing
        End If
        Dim sqlstr = String.Format("select * from {0} where parent = :parent and child = :child limit 1", itemChildTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("parent", parent.name, DbType.String))
        'dbparams.Add(dbadapter1.getParam("child", child.name, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("parent", parent.name)
        params(1) = DataAccess.factory.CreateParameter("child", child.name)
        Dim ra As Object = Nothing
        Dim message As String = String.Empty
        'If dbadapter1.ExecuteScalar(sqlstr, params, ra, message) Then
        '    Return Not IsNothing(ra)
        'End If
        ra = DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params)
        Return Not IsNothing(ra)
        'Return False
    End Function

    Public Overrides Function removeAllAssignments() As Boolean
        Dim sqlstr = String.Format("delete from {0}", assignmentTable)
        'Return dbadapter1.ExecuteNonQuery(sqlstr)
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
    End Function

    Public Overrides Function removeAllPermission() As Boolean
        Return removeAllItems(TypeEnum.TYPE_PERMISSION)
    End Function

    Private Function removeAllItems(typeEnum As TypeEnum) As Boolean

        Dim sqlstr = String.Format("select name from {0} where type = :type", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("type", DirectCast(typeEnum, Integer), DbType.Int16))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("type", DirectCast(typeEnum, Integer))

        'Dim ds As New DataSet
        'If Not dbadapter1.GetDataset(sqlstr, ds, params) Then
        '    Return Nothing
        'End If
        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, params)
        If IsNothing(ds) Then
            Return Nothing
        End If
        Dim sb As New StringBuilder
        If ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        End If
        For Each dr In ds.Tables(0).Rows
            If sb.Length > 0 Then
                sb.Append(",")
            End If
            sb.Append(String.Format("'{0}'", dr.item("name")))
        Next

        'delete table auth_item_child where : if typeenum permission , child else parent in sb.tostring
        Select Case typeEnum
            Case typeEnum.TYPE_PERMISSION
                sqlstr = String.Format("delete from {0} where child in ({1}) ", itemChildTable, sb.ToString)
            Case typeEnum.TYPE_ROLE
                sqlstr = String.Format("delete from {0} where parent in ({1}) ", itemChildTable, sb.ToString)
        End Select
        'dbadapter1.ExecuteNonQuery(sqlstr)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)

        'delete table auth_assignment where item_name in sb.tostring
        sqlstr = String.Format("delete from {0} where item_name in ({1})", assignmentTable, sb.ToString)
        'dbadapter1.ExecuteNonQuery(sqlstr)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)

        'delete table auth_item where name int sb.tostring
        sqlstr = String.Format("delete from {0} where name in ({1})", itemTable, sb.ToString)
        'dbadapter1.ExecuteNonQuery(sqlstr)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        Return True
    End Function

    Public Overrides Function removeAllRoles() As Boolean
        Return removeAllItems(TypeEnum.TYPE_ROLE)
    End Function

    Public Overrides Function removeAllRules() As Boolean
        Dim sqlstr = String.Format("update {0} set rule_name = null", itemTable)
        'dbadapter1.ExecuteNonQuery(sqlstr)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)

        sqlstr = String.Format("delete from {0}", ruleTable)
        'dbadapter1.ExecuteNonQuery(sqlstr)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        Return True
    End Function

    Public Overrides Function removeChild(parent As Item, child As Item) As Boolean

        Dim sqlstr = String.Format("delete from {0} where parent = :parent and child = :child", itemChildTable)
        'Dim dbparam As New List(Of IDataParameter)
        'dbparam.Add(dbadapter1.getParam("parent", parent.name, DbType.String))
        'dbparam.Add(dbadapter1.getParam("child", child.name, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("parent", parent.name)
        params(1) = DataAccess.factory.CreateParameter("child", child.name)
        Dim ra As Object = Nothing
        'If dbadapter1.ExecuteNonQuery(sqlstr, params, recordAffected:=ra) Then
        '    Return IsNothing(ra)
        'End If
        ra = DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        Return IsNothing(ra)
    End Function

    Public Overrides Function removeChildren(parent As Item) As Boolean

        Dim sqlstr = String.Format("delete from {0} where parent = :parent", itemChildTable)
        Dim ra As Integer = 0
        'Dim dbparam As New List(Of IDataParameter)
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("parent", parent.name)
        'If dbadapter1.ExecuteNonQuery(sqlstr, params:=params, recordAffected:=ra) Then
        '    Return ra > 0
        'End If
        ra = DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        Return ra > 0
        Return False
    End Function

    Public Overrides Function removeItem(item As Item) As Boolean


        Dim sqlstr = String.Format("delete from {0} where child = :itemname or parent = : itemname ", itemChildTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("parent", item.name, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("parent", item.name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)

        'delete table auth_assignment where item_name in sb.tostring
        sqlstr = String.Format("delete from {0} where item_name = :itemname", assignmentTable)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)

        'delete table auth_item where name int sb.tostring
        sqlstr = String.Format("delete from {0} where name = :itemname", itemTable)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        Return True
    End Function

    Public Overrides Function removeRule(rule As Rule) As Boolean
        Dim sqlstr = String.Format("update {0} set rule_name = null where rule_name = :rulename", itemTable)

        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("rulename", rule.name, DbType.String))
        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("rulename", rule.name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)

        sqlstr = String.Format("delete from {0} where name = :rulename", ruleTable)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        Return True
    End Function

    Public Overrides Function revoke(role As Role, userid As Object) As Boolean
        If IsNothing(userid) Then
            Return False
        End If
        Dim sqlstr = String.Format("delete from {0} where user_id::integer = :userid and item_name = :rolename", assignmentTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("userid", userid, DbType.String))
        'dbparams.Add(dbadapter1.getParam("rolename", role.name, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("userid", userid)
        params(1) = DataAccess.factory.CreateParameter("rolename", role.name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        Return True
    End Function

    Public Overloads Overrides Function revokeAll() As Boolean
        removeAllAssignments()
        'Dim sb As New StringBuilder
        'sb.Append(String.Format("delete from {0};", itemChildTable))
        'sb.Append(String.Format("delete from {0};"), itemTable)
        'sb.Append(String.Format("delete from {0};"), ruleTable)
        'dbadapter1.ExecuteNonQuery(sb.ToString)
        Dim sqlstr = String.Format("delete from {0}", itemChildTable)
        'Return dbadapter1.ExecuteNonQuery(sqlstr)
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        sqlstr = String.Format("delete from {0}", itemTable)
        'Return dbadapter1.ExecuteNonQuery(sqlstr)
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        sqlstr = String.Format("delete from {0}", ruleTable)
        'Return dbadapter1.ExecuteNonQuery(sqlstr)
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        'Return True
    End Function

    Public Overloads Overrides Function revokeAll(userid As Object) As Boolean
        If IsNothing(userid) Then
            Return False
        End If


        Dim sqlstr = String.Format("delete from {0} where user_id = :userid", assignmentTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("userid", userid, DbType.String))

        Dim params(0) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("userid", userid)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        Return DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, Nothing)
        Return True
    End Function

    Public Overrides Function updateItem(name As String, item As Item) As Boolean

        Dim sqlstr = String.Format("update {0} set parent = :itemname where parent = :name", itemChildTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("itemname", item.name, DbType.String))
        'dbparams.Add(dbadapter1.getParam("name", name, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("itemname", item.name)
        params(1) = DataAccess.factory.CreateParameter("name", name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        sqlstr = String.Format("update {0} set child = :itemname where child = :name", itemChildTable)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        sqlstr = String.Format("update {0} set item_name = :itemname where item_name = :name", assignmentTable)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)
        item.updatedAt = Now

        
        sqlstr = String.Format("update {0} set name = :itemname,description = :description,rule_name = :rulename,data=:data, updated_at =:updatedat where name=:name", itemTable)

        'dbparams.Add(dbadapter1.getParam("description", item.description, DbType.String))
        'dbparams.Add(dbadapter1.getParam("rulename", item.ruleName, DbType.String))
        'dbparams.Add(dbadapter1.getParam("data", item.data, DbType.Binary))
        'dbparams.Add(dbadapter1.getParam("updatedat", item.updatedAt, DbType.String))

        Dim params2(5) As IDbDataParameter
        params2(0) = DataAccess.factory.CreateParameter("itemname", item.name)
        params2(1) = DataAccess.factory.CreateParameter("name", name)
        params2(2) = DataAccess.factory.CreateParameter("description", item.description)
        params2(3) = DataAccess.factory.CreateParameter("rulename", item.ruleName)
        params2(4) = DataAccess.factory.CreateParameter("data", item.data)
        params2(5) = DataAccess.factory.CreateParameter("updatedat", item.updatedAt)
        'dbadapter1.ExecuteNonQuery(sqlstr, params2)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params2)
        Return True

    End Function

    Public Overrides Function updateRule(name As String, rule As Rule) As Boolean

        Dim sqlstr = String.Format("update {0} set rule_name = :rulename where rule_name = :name", itemTable)
        'Dim dbparams As New List(Of IDataParameter)
        'dbparams.Add(dbadapter1.getParam("rulename", rule.name, DbType.String))
        'dbparams.Add(dbadapter1.getParam("name", name, DbType.String))
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("rulename", rule.name)
        params(1) = DataAccess.factory.CreateParameter("name", name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params)

        rule.updatedAt = Now

        sqlstr = String.Format("update {0} set name = :rulename,updated_at =:updatedat where name=:name", ruleTable)
        'dbparams.Add(dbadapter1.getParam("rulename", rule.name, DbType.String))
        'dbparams.Add(dbadapter1.getParam("updatedat", rule.updatedAt, DbType.String))

        Dim params2(2) As IDbDataParameter
        params2(0) = DataAccess.factory.CreateParameter("rulename", rule.name)
        params2(1) = DataAccess.factory.CreateParameter("updatedat", rule.updatedAt)
        params2(2) = DataAccess.factory.CreateParameter("name", name)
        'dbadapter1.ExecuteNonQuery(sqlstr, params2)
        DataAccess.ExecuteNonQuery(sqlstr, CommandType.Text, params2)
        Return True
    End Function

    Public Function isAssignedto(userid As Object, role As String) As Boolean
        Dim sqlstr = String.Format("select count(0) from {0} where item_name = :role and user_id = :userid", assignmentTable)
        Dim params(1) As IDbDataParameter
        params(0) = DataAccess.factory.CreateParameter("role", role)
        params(1) = DataAccess.factory.CreateParameter("userid", userid.ToString)

        Return DataAccess.ExecuteScalar(sqlstr, CommandType.Text, params) > 0
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

Public Class Auth_Item_Class
    Public Property parent As String
    Public Property child As String
End Class