Imports System.Text
Public Class ExposureModel
    Implements IModel


    Public Property stPeriod As Date
    Public Property ndPeriod As Date

    Public ReadOnly Property FilterField
        Get
            Return ""
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cp.expensestype"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "id"
        End Get
    End Property


    Private Function GetSqlstr(ByVal criteria) As String

        'FGProjectFamily
        'FGCMMF
        'FGBOM
        'FGBomTX
        'FGBOMUsage
        'FGBOMUsageTx
        Dim sb As New StringBuilder
        sb.Append(String.Format("select * from cp.fgprojectfamily;"))
        sb.Append(String.Format("select * from cp.fgcmmf;"))
        sb.Append(String.Format("select * from cp.vendor;"))
        sb.Append(String.Format("select * from cp.fgbom;"))
        sb.Append(String.Format("select * from cp.fgbomusage;"))
        sb.Append(String.Format("select * from cp.fgbomtx;"))
        sb.Append(String.Format("select * from cp.fgbomusagetx;"))
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'Division' ;"))
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'Main Reason' ;"))
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = 'Component Category' ;"))
        sb.Append(String.Format("select vendorcode,vendorname from vendor;")) 'SAP Vendor
        Return sb.ToString
    End Function

    Public Function LoadData1(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Return False
    End Function

    Public Function GetExpensesTypeBS(ByVal criteria As String) As BindingSource
        Dim ds As New DataSet
        Dim ExpensesTypeBS As New BindingSource
        Dim sqlstr = GetSqlstr(criteria)
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        ds.Tables(0).TableName = TableName
        ExpensesTypeBS.DataSource = ds.Tables(0)
        Return ExpensesTypeBS
    End Function

    Public Function GetPeriod() As BindingSource
        Dim ds As New DataSet
        Dim ExpensesTypeBS As New BindingSource
        Dim sqlstr = "select distinct txdate,to_char(txdate,'DD-Mon-yyyy') as txdatestring  from cp.fgbomtx order by txdate desc"
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        ds.Tables(0).TableName = TableName
        ExpensesTypeBS.DataSource = ds.Tables(0)
        Return ExpensesTypeBS
    End Function
    Public Function GetPeriodAll() As BindingSource
        Dim ds As New DataSet
        Dim ExpensesTypeBS As New BindingSource
        Dim sqlstr = "select '2000-01-01'::date as txdate,'All' as txdatestring union all (select distinct txdate,to_char(txdate,'DD-Mon-yyyy') as txdatestring  from cp.fgbomtx order by txdate desc)"
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        ds.Tables(0).TableName = TableName
        ExpensesTypeBS.DataSource = ds.Tables(0)
        Return ExpensesTypeBS
    End Function

    Public Function LoadData(ByRef DS As DataSet, ByVal criteria As String) As Boolean
        Dim sqlstr = GetSqlstr("")
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        DS.Tables(0).TableName = TableName
        Return True
    End Function

    Public Function GetDataSet() As DataSet
        Dim ds As New DataSet
        Dim sqlstr = GetSqlstr("")
        DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        Return ds
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

            sqlstr = "cp.sp_insertvendor"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "vendorcode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "shortnamecp", DataRowVersion.Current))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updatevendor"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "vendorcode", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "vendorcode", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "shortnamecp", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables("Vendor"))

            sqlstr = "cp.sp_insertcmmf"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "modelnumber", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "description", DataRowVersion.Current))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure
            dataadapter.InsertCommand.Transaction = mytransaction
            mye.ra = factory.Update(mye.dataset.Tables("FGCMMF"))

            sqlstr = "cp.sp_insertfgprojectfamily"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "vendorcode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "projectfamily", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "colt", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updatefgprojectfamily"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))            
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "colt", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction


            mye.ra = factory.Update(mye.dataset.Tables("FGProjectFamily"))

            sqlstr = "cp.sp_insertfgbom"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "fgprojectfamilyid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "createddate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "componentcategoryid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "partnumber", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "componentdescription", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "modelversion", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "leadtime", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "moq", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialvendorname", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialvendorlocation", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "mainreasonid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "unit", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "division", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updatefgbom"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "fgprojectfamilyid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "createddate", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "componentcategoryid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "partnumber", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "componentdescription", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "modelversion", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "leadtime", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "moq", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialvendorname", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "materialvendorlocation", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "mainreasonid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "unit", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "division", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            mye.ra = factory.Update(mye.dataset.Tables("FGBOM"))

            sqlstr = "cp.sp_insertfgbomTX"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "fgbomid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "txdate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "unitprice", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "stock", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_deletefgbomtx"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id"))
            dataadapter.DeleteCommand.Parameters(0).Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure


            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction
            mye.ra = factory.Update(mye.dataset.Tables("FGBOMTX"))

            sqlstr = "cp.sp_insertfgbomusage"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "fgbomid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables("FGBOMUsage"))

            sqlstr = "cp.sp_insertfgbomusagetx"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "fgbomusageid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "txdate", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "quantity", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure
            sqlstr = "cp.sp_deletefgbomusagetx"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id"))
            dataadapter.DeleteCommand.Parameters(0).Direction = ParameterDirection.Input
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.DeleteCommand.Transaction = mytransaction
            dataadapter.InsertCommand.Transaction = mytransaction
            mye.ra = factory.Update(mye.dataset.Tables("FGBOMUsageTX"))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function GetSQLSTRReport(Criteria As String) As String
        Dim sqlstr As String
        sqlstr = String.Format("( with dbd as (select * from cp.fgdbdemand union all select * from cp.cpdbdemand)" &
                 " select pf.*,v.vendorname,cpv.shortnamecp,b.*,cc.tvalue as groupcomponentcategory,cc.cvalue as componentcategory,mr.cvalue as mainreason,dv.tvalue as groupdivision,dv.cvalue as divisionname,btx.*,bu.*,butx.*,dbd.yearweek,dbd.qty,w.monthly,(btx.unitprice * coalesce(dbd.qty,0) * coalesce(butx.quantity,0))*-1 as exposure from cp.fgprojectfamily pf" &
                 " left join vendor v on v.vendorcode = pf.vendorcode" &
                 " left join cp.vendor cpv on cpv.vendorcode = pf.vendorcode" &
                 " left join cp.fgbom b on b.fgprojectfamilyid = pf.id" &
                 " left join cp.paramdt cc on cc.paramdtid = b.componentcategoryid" &
                 " left join cp.paramdt mr on mr.paramdtid = b.mainreasonid" &
                 " left join cp.paramdt dv on dv.paramdtid = b.division" &
                 " left join cp.fgbomtx btx on btx.fgbomid = b.id" &
                 " left join cp.fgbomusage bu on bu.fgbomid = b.id" &
                 " left join cp.fgbomusagetx butx on butx.fgbomusageid = bu.id" &
                 " left join dbd on dbd.vendorcode = pf.vendorcode and dbd.cmmf = bu.cmmf" &
                 " left join weektomonth w on w.yearweek = dbd.yearweek where (not b.partnumber isnull) and (not btx.txdate isnull) {0})" &
                 " union all" &
                 " (with q2 as (select pf.*,v.vendorname,cpv.shortnamecp,b.*,cc.tvalue as groupcomponentcategory,cc.cvalue as componentcategory,mr.cvalue as mainreason,dv.tvalue as groupdivision,dv.cvalue as divisionname,btx.*,null::bigint,null::bigint,null::bigint,null::bigint,null::bigint,null::date,null::numeric,null::integer,null::integer,'2000-01-01'::date,(btx.unitprice * btx.stock) as exposure from cp.fgprojectfamily pf" &
                 " left join vendor v on v.vendorcode = pf.vendorcode" &
                 " left join cp.vendor cpv on cpv.vendorcode = pf.vendorcode" &
                 " left join cp.fgbom b on b.fgprojectfamilyid = pf.id" &
                 " left join cp.paramdt cc on cc.paramdtid = b.componentcategoryid" &
                 " left join cp.paramdt mr on mr.paramdtid = b.mainreasonid" &
                 " left join cp.paramdt dv on dv.paramdtid = b.division" &
                 " left join cp.fgbomtx btx on btx.fgbomid = b.id where (not b.partnumber isnull) {0}) select * from q2 where not exposure isnull)", Criteria)
        Return sqlstr
        'null::bigint,null::bigint,null::bigint,null::bigint,null::bigint,null::date,null::numeric,null::integer,null::integer,'2000-01-01'::date,
        'null,null,null,null,null,null,null,null,null,'2000-01-01',
    End Function


End Class
