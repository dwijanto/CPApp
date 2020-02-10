Public Class UserController
    Inherits ActiveRecord
    Implements IController
    Implements IToolbarAction
    Implements IIdentity



    Public Property userid As String
    Public Property username As String
    Public Property id As Object
    Public Property isAdmin As Boolean
    Public Property email As String
    Public Property isActive As Boolean
    Public Property password_hash As String
    Public Property deptid As Integer
    Public Property ApprovalId As Integer
   
    Public Model As New UserModel    
    Dim DS As DataSet
    Public BS As BindingSource


    Public Sub New()
        MyBase.New()
        tableName = "cp.user" 'tablename
        primarykey = "id" 'primarykey
    End Sub

    Public ReadOnly Property GetTable As DataTable Implements IController.GetTable
        Get
            Return DS.Tables(Model.TableName).copy()
        End Get
    End Property
   

    Public ReadOnly Property GetBindingSource As BindingSource
        Get
            Dim BS As New BindingSource
            BS.DataSource = GetTable
            BS.Sort = Model.SortField
            Return BS
        End Get
    End Property

    Public Function loaddata() As Boolean Implements IController.loaddata
        Dim myret As Boolean = False
        Model = New UserModel
        DS = New DataSet
        If Model.LoadData(DS) Then
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = pk
            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1
            BS = New BindingSource


            BS.DataSource = DS.Tables(0)
           
            myret = True
        End If
        Return myret
    End Function

    Public Function save() As Boolean Implements IController.save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean Implements IToolbarAction.Save
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

    Public Property ApplyFilter As String Implements IToolbarAction.ApplyFilter
        Get
            Return BS.Filter
        End Get
        Set(ByVal value As String)
            BS.Filter = String.Format(Model.FilterField, value)
        End Set
    End Property
    Public Function GetCurrentRecord() As DataRowView Implements IToolbarAction.GetCurrentRecord
        Return BS.Current
    End Function

    Public Function GetNewRecord() As DataRowView Implements IToolbarAction.GetNewRecord
        Return BS.AddNew
    End Function

    Public Sub RemoveAt(value As Integer) Implements IToolbarAction.RemoveAt
        BS.RemoveAt(value)
    End Sub

    Public Function findIdentity(id As Object) As Object Implements IIdentity.findIdentity
        Dim ds As DataSet = Me.findOne(id)
        If Not IsNothing(ds) Then
            Return populatedata(ds.Tables(0).Rows(0))
        Else
            Return Nothing
        End If
    End Function

    Private Function populatedata(dr As DataRow) As IIdentity
        Dim Identity = New UserController With {.id = dr.Item("id"),
                                                .username = "" & dr.Item("username"),
                                                .isAdmin = dr.Item("isadmin"),
                                                .isActive = dr.Item("isactive"),
                                                .userid = "" & dr.Item("userid"),
                                                .email = "" & dr.Item("email"),
                                                .password_hash = ""}
        Return Identity
    End Function
    Public Function findIdentityByAccessToken(token As Object, Optional type As Object = Nothing) As Object Implements IIdentity.findIdentityByAccessToken
        Dim myCondition As New Hashtable
        myCondition.Add("accestoken", token)
        Return findOne(myCondition)
    End Function

    Public Function getAuthKey() As String Implements IIdentity.getAuthKey
        Throw New NotImplementedException
    End Function

    Public Function getId() As Object Implements IIdentity.getId
        Return _id
    End Function

    Public Function validateAuthKey(authkey As String) As Boolean Implements IIdentity.validateAuthKey
        Throw New NotImplementedException
    End Function
    Public Function findByUserName(ByVal username As String)
        Dim myCondition As New Hashtable
        myCondition.Add("lower(username)", username.ToLower)
        Return findOne(myCondition)
    End Function

    Public Function findByUserid(ByVal userid As String)
        Dim myCondition As New Hashtable
        myCondition.Add("lower(userid)", userid.ToLower)
        Return findOne(myCondition)
    End Function




End Class
