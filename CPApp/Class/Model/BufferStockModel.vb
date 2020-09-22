Imports System.Text
Public Class BufferStockModel
    Implements IModel


    Public Property customername As String
    Public Property customercode As String
    Public Property vendorname As String
    Public Property vendorcode As String
    Public Property partnumber As String
    Public Property description As String
    Public Property projectname As String
    Public Property leadtime As String
    Public Property t2 As String
    Public Property bufferqty As String
    Public Property unit As String
    Public Property unitprice As String


    Public ReadOnly Property FilterField
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cp.bufferstock"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property


    Private Function GetSqlstr(ByVal criteria) As String
        Dim sb As New StringBuilder
        sb.Append(String.Format("select * from {0} u {1} ;select * from cp.customer;select * from cp.vendor;", TableName, criteria))
        Return sb.ToString
    End Function

    Public Function GetSQLSTRReport(ByVal criteria) As String
        Dim sb As New StringBuilder
        sb.Append(String.Format("select u.*,c.customername,v.vendorname,u.bufferqty * u.unitprice as total from {0} u" &
                                " left join cp.customer c on c.customercode = u.customercode " &
                                " left join cp.vendor v on v.vendorcode = u.vendorcode {1} ;", TableName, criteria))
        Return sb.ToString
    End Function

    'Public Function LoadData1(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
    '    Return False
    'End Function

    'Public Function GetBPartnerBS() As BindingSource
    '    Dim ds As New DataSet
    '    Dim ExpensesTypeBS As New BindingSource
    '    Dim sqlstr = "select bpa.id as id,bp.bpname as bpartnername,coalesce(bpa.line1,'') || coalesce(bpa.line2,'') || coalesce(bpa.line3,'') as bpartneraddress,bp.bpcode,bp.bpcode || ' - ' || bp.bpname  || ' (' || bpa.addressid || ')' as bpartnerfullname ,bpa.region,bpa.country  " &
    '                 " from cp.bpartner bp left join cp.bpaddress bpa on bpa.bpid = bp.id and bpa.addresstype = 'S' where not bpa.id isnull order by bpcode,bpartneraddress"
    '    ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
    '    ds.Tables(0).TableName = TableName
    '    ExpensesTypeBS.DataSource = ds.Tables(0)
    '    Return ExpensesTypeBS
    'End Function

    Public Function LoadData(ByRef DS As DataSet, ByVal criteria As String) As Boolean
        Dim sqlstr = GetSqlstr("")
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName
        Return True
    End Function

    Public Function save(ByVal obj As Object, ByVal mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim myret As Boolean = False
        Dim factory = DataAccess.factory
        Dim mytransaction As IDbTransaction
        Using conn As IDbConnection = factory.CreateConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            Dim dataadapter = factory.CreateAdapter
            Dim sqlstr As String = String.Empty

            'sqlstr = "cp.sp_insertexpensestype"
            'dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            'dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "deptid", DataRowVersion.Current))
            'dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "expensesacc", DataRowVersion.Current))
            'dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "expensesname", DataRowVersion.Current))
            'dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            'dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            'sqlstr = "cp.sp_updateexpensestype"
            'dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            'dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            'dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "deptid", DataRowVersion.Current))
            'dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "expensesacc", DataRowVersion.Current))
            'dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "expensesname", DataRowVersion.Current))
            'dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            'sqlstr = "cp.sp_deleteexpensestype"
            'dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            'dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            'dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            'dataadapter.InsertCommand.Transaction = mytransaction
            'dataadapter.UpdateCommand.Transaction = mytransaction
            'dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(TableName))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Public Function LoadData1(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Return False
    End Function
End Class
