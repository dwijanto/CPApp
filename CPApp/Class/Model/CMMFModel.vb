Imports System.Text

Public Class CMMFModel
    Implements IModel



    Public ReadOnly Property FilterField
        Get
            Return "[cmmfstring] like '%{0}%' or [localdescription] like '%{0}%' or [commercialcode] like '%{0}%'"
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cp.cmmf"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "cmmf"
        End Get
    End Property


    Private Function GetSqlstr(ByVal criteria) As String
        Dim sb As New StringBuilder
        sb.Append(String.Format("select u.*,u.cmmf::character varying as cmmfstring,coalesce(cp.pricehkd,0) as price,coalesce(cp.priceusd,0) as priceusd from {0} {1} u left join cp.cmmfprice cp on cp.cmmf = u.cmmf ", TableName, criteria))
        Return sb.ToString
    End Function

    Public Function LoadData1(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Return False
    End Function


    Public Function GetCMMFBS() As BindingSource
        Dim ds As New DataSet
        Dim CMMFBS As New BindingSource
        Dim sqlstr = GetSqlstr("")
        ds = DataAccess.GetDataSet(sqlstr, CommandType.Text, Nothing)
        ds.Tables(0).TableName = TableName
        CMMFBS.DataSource = ds.Tables(0)
        Return CMMFBS
    End Function

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

            sqlstr = "cp.sp_insertcmmf"

            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "commercialcode", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "localdescription", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "price", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "pricehusd", DataRowVersion.Current))

            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updatecmmf"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "commercialcode", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "localdescription", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "price", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "priceusd", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_deletecmmf"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "cmmf", DataRowVersion.Original))
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(TableName))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function


End Class
