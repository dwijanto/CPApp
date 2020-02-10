Imports System.Text

Public Class ProductRequestModel
    Implements IModel
    Private _criteria As String = String.Empty

    Public Property Criteria As String
        Get
            Return _criteria
        End Get
        Set(value As String)
            _criteria = value
        End Set
    End Property

    Public ReadOnly Property FilterField
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "prhd"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property

    Private Function GetSqlstr(ByVal criteria) As String
        Dim sb As New StringBuilder
        'sb.Append(String.Format("select *,bp.name as bpartnername,bp.address as bpartneraddress,bp.bpartner,bp.bpartner || ' - ' || bp.name as bpartnerfullname  from cp.prhd hd " &
        '                        "left join cp.bpartner bp on bp.id = hd.bpartnerid {0};", criteria))
        sb.Append(String.Format("select *,bp.bpname as bpartnername,coalesce(bpa.line1,'') || coalesce(bpa.line2,'') || coalesce(bpa.line3,'') as bpartneraddress,bp.bpcode,bp.bpcode || ' - ' || bp.bpname || ' (' || bpa.addressid || ')' as bpartnerfullname,bpa.region,bpa.country ,us.email as applicantemail,up.username as approvalname,cp.getapprovaldate(hd.id) as approvaldate " &
                                " from cp.prhd hd left join cp.bpaddress bpa on bpa.id = hd.bpartnerid left join cp.bpartner bp on bp.id = bpa.bpid left join cp.user us on us.userid = hd.createdby left join cp.user up on up.userid = hd.deptapproval {0};", criteria))
        sb.Append(String.Format("select dt.*,c.commercialcode,c.localdescription,e.expensesname,coalesce(dt.price * dt.qty,0) as total from cp.prdt dt" &
                                " inner join cp.prhd hd on hd.id = dt.prhdid" &
                                " left join cp.cmmf c on c.cmmf = dt.cmmf " &
                                " left join cp.expensestype e on e.id = dt.expensestypeid " &
                                " {0} order by dt.id;", criteria))
        sb.Append(String.Format("select pa.*,cp.getstatusname(pa.status) as statusname from cp.praction pa" &
                                " inner join cp.prhd hd on hd.id = pa.prhdid {0};", criteria))
        Return sb.ToString
    End Function


    Public Function LoadData(ByRef DS As DataSet) As Boolean Implements IModel.LoadData        
        Dim sqlstr = GetSqlstr(criteria)
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName
        DS.Tables(1).TableName = "PRDT"
        DS.Tables(2).TableName = "PRAct"

        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1

        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(1).Columns("id")
        DS.Tables(1).PrimaryKey = pk1
        DS.Tables(1).Columns("id").AutoIncrement = True
        DS.Tables(1).Columns("id").AutoIncrementSeed = -1
        DS.Tables(1).Columns("id").AutoIncrementStep = -1

        Dim pk2(0) As DataColumn
        pk2(0) = DS.Tables(2).Columns("id")
        DS.Tables(2).PrimaryKey = pk2
        DS.Tables(2).Columns("id").AutoIncrement = True
        DS.Tables(2).Columns("id").AutoIncrementSeed = -1
        DS.Tables(2).Columns("id").AutoIncrementStep = -1

        'Create Relation HD-DT
        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn

        hcol = DS.Tables(0).Columns("id") 'id in table header
        dcol = DS.Tables(1).Columns("prhdid") 'headerid in table detail
        rel = New DataRelation("hdrel", hcol, dcol)
        DS.Relations.Add(rel)

        hcol = DS.Tables(0).Columns("id") 'id in table header
        dcol = DS.Tables(2).Columns("prhdid") 'headerid in table detail
        rel = New DataRelation("hdrel-action", hcol, dcol)
        DS.Relations.Add(rel)
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

            sqlstr = "cp.sp_insertprhd"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "applicantname", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "applicantdate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "sendto", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "reason", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "deliverydate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "address", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "instruction", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "deptid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "deptapproval", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "mdapproval", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "status", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "createdby", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "createddate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "bpartnerid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "refnumber", ParameterDirection.Output))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updateprhd"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "applicantname", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "applicantdate", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "sendto", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "reason", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "deliverydate", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "address", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "instruction", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "deptid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "deptapproval", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "mdapproval", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "status", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "createdby", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "createddate", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "bpartnerid", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_deleteprhd"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(TableName))

            sqlstr = "cp.sp_insertprdt"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "prhdid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "price", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "confirmedqty", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "remarks", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "expensestypeid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "createdby", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "createddate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updateprdt"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "prhdid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "qty", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "price", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "confirmedqty", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "remarks", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "expensestypeid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "createdby", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "createddate", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_deleteprdt"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(1))

            sqlstr = "cp.sp_insertpraction"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "prhdid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "status", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "modifiedby", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "latestupdate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "remarks", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(2))


            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function FindLoadData(ByRef DS As DataSet) As Boolean
        Dim myret As Boolean = True
        Try
            Dim sb As New StringBuilder
            sb.Append(String.Format("select cp.getstatusname(status) as statusname,hd.* from cp.prhd hd {0};", Criteria))
            DS = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)
            DS.Tables(0).TableName = TableName
        Catch ex As Exception
            myret = False
        End Try
        Return myret
    End Function

    Function MyTasksLoadData(ByRef DS As DataSet, MyTaskcriteria As String, Historycriteria As String) As Boolean
        Dim myret As Boolean = True
        Try
            Dim sb As New StringBuilder
            sb.Append(String.Format("select cp.getstatusname(status) as statusname,hd.* from cp.prhd hd {0};", MyTaskcriteria))
            sb.Append(String.Format("select cp.getstatusname(status) as statusname,hd.* from cp.prhd hd {0};", Historycriteria))
            DS = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myret = False
        End Try
        Return myret
    End Function

End Class
