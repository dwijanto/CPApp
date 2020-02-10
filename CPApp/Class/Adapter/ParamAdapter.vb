Imports Npgsql
Imports System.Text

Public Class ParamAdapter
    Public Shared myInstance As ParamAdapter
    Dim factory = DataAccess.factory
    Private Const SchemaName As String = "cp"
    Public DS As DataSet
    Public BS As BindingSource
    Public BS2 As BindingSource
    Public BS3 As BindingSource
    Private ErrMessage As String = String.Empty

    Public ReadOnly Property getErrorMessage As String
        Get
            Return ErrMessage
        End Get
    End Property

    Public Shared Function getInstance() As ParamAdapter
        If myInstance Is Nothing Then
            myInstance = New ParamAdapter
        End If
        Return myInstance
    End Function

    Public Function GetParentid(ByVal paramName As String) As String
        Dim sqlstr = String.Format("select paramhdid from cp.paramhd where paramname =:paramname", paramName)
        Dim myresult As String = String.Empty
        Dim myparam(0) As System.Data.IDbDataParameter
        myparam(0) = factory.CreateParameter("paramname", paramName)
        myresult = DataAccess.ExecuteScalar(sqlstr, CommandType.Text, myparam)
        Return myresult
    End Function
    Function getLimit(ByVal paramName) As Decimal
        Dim sqlstr = String.Format("select dt.nvalue,ph.* from cp.paramdt dt left join cp.paramhd ph on ph.paramhdid = dt.paramhdid where dt.paramname =:paramname and ph.paramname = 'DirectorLimit'", paramName)
        Dim myresult As String = String.Empty
        Dim myparam(0) As System.Data.IDbDataParameter
        myparam(0) = factory.CreateParameter("paramname", paramName)
        myresult = DataAccess.ExecuteScalar(sqlstr, CommandType.Text, myparam)
        Return myresult
    End Function
    Public Function GetApprovalName(ByVal paramdtId As Integer) As ApprovalInfo
        Dim myApprovalInfo As ApprovalInfo = Nothing
        Try
            Dim sqlstr = String.Format("select u.* from cp.paramdt pd " &
                                   " left join cp.user u on u.userid = pd.cvalue where pd.paramdtid =:id;", paramdtId)
            Dim myresult As String = String.Empty
            Dim myparam(0) As System.Data.IDbDataParameter
            myparam(0) = factory.CreateParameter("id", paramdtId)
            'myresult = DataAccess.ExecuteScalar(sqlstr, CommandType.Text, myparam)
            DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, myparam)
            If DS.Tables(0).Rows.Count > 0 Then
                myApprovalInfo = New ApprovalInfo With {.ID = DS.Tables(0).Rows(0).Item("userid"),
                                                    .Name = DS.Tables(0).Rows(0).Item("username"),
                                                    .Email = DS.Tables(0).Rows(0).Item("email")
                                                    }
            End If
            
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return myApprovalInfo
    End Function

    Public Function GetEmail(ByVal paramname As String) As String

        Dim sb As New StringBuilder
        Try
            Dim sqlstr = String.Format("select dt.cvalue from cp.paramdt dt" &
                                       " left join cp.paramhd ph on ph.paramhdid = dt.paramhdid " &
                                       " where ph.paramname = 'SupplyChainEmail' and dt.paramname =:paramname ", paramname)
            Dim myresult As String = String.Empty
            Dim myparam(0) As System.Data.IDbDataParameter
            myparam(0) = factory.CreateParameter("paramname", paramname)

            DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, myparam)

            If DS.Tables(0).Rows.Count > 0 Then
                For Each drv As DataRow In DS.Tables(0).Rows
                    If sb.Length > 0 Then
                        sb.Append(";")
                    End If
                    sb.Append(drv.Item("cvalue"))
                Next
            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        Return sb.ToString
    End Function

    Public Function LoadData()
        Dim sb As New StringBuilder
        Dim myret As Boolean = False
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :approvalparam order by paramdtid;"))
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :exrate order by cvalue;"))
        sb.Append(String.Format("select pd.* from cp.paramdt pd left join cp.paramhd ph on ph.paramhdid = pd.paramhdid where ph.paramname = :supplychainemail order by paramname desc;"))
        Dim sqlstr = sb.ToString
        DS = New DataSet
        BS = New BindingSource
        BS2 = New BindingSource
        BS3 = New BindingSource

        Dim myparam(2) As System.Data.IDbDataParameter
        myparam(0) = factory.CreateParameter("approvalparam", "Approval")
        myparam(1) = factory.CreateParameter("exrate", "ExRate")
        myparam(2) = factory.CreateParameter("supplychainemail", "SupplyChainEmail")
        Try
            DS = DataAccess.GetDataSet(sqlstr, CommandType.Text, myparam)

            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("paramdtid")
            DS.Tables(0).PrimaryKey = pk
            BS.DataSource = DS.Tables(0)
            DS.Tables(0).Columns("paramdtid").AutoIncrement = True
            DS.Tables(0).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(0).Columns("paramdtid").AutoIncrementStep = -1

            Dim pk2(0) As DataColumn
            pk2(0) = DS.Tables(1).Columns("paramdtid")
            DS.Tables(1).PrimaryKey = pk2
            BS2.DataSource = DS.Tables(1)
            DS.Tables(1).Columns("paramdtid").AutoIncrement = True
            DS.Tables(1).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(1).Columns("paramdtid").AutoIncrementStep = -1
            myret = True

            Dim pk3(0) As DataColumn
            pk3(0) = DS.Tables(2).Columns("paramdtid")
            DS.Tables(2).PrimaryKey = pk3
            BS3.DataSource = DS.Tables(2)
            DS.Tables(2).Columns("paramdtid").AutoIncrement = True
            DS.Tables(2).Columns("paramdtid").AutoIncrementSeed = -1
            DS.Tables(2).Columns("paramdtid").AutoIncrementStep = -1
            myret = True
        Catch ex As Exception
            ErrMessage = ex.Message
        End Try
        Return myret
    End Function

    Public Function save() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()
        BS2.EndEdit()

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

    Public Function Save(mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        Dim factory = DataAccess.factory
        Dim mytransaction As IDbTransaction
        Using conn As IDbConnection = factory.CreateConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            Dim dataadapter = factory.CreateAdapter
            Dim sqlstr As String

            sqlstr = "cp.sp_deleteparameter"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "paramdtid"))
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_insertparameter"

            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "paramhdid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "paramname", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "cvalue", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "dvalue", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ivalue", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "nvalue", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "ts", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "paramdtid", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_updateparameter"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "paramdtid", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "paramhdid", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "paramname", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "cvalue", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Date, 0, "dvalue", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int32, 0, "ivalue", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Decimal, 0, "nvalue", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.DateTime, 0, "ts", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = factory.Update(mye.dataset.Tables(0))
            mye.ra = factory.Update(mye.dataset.Tables(1))
            mye.ra = factory.Update(mye.dataset.Tables(2))
            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function




End Class
Public Class ApprovalInfo
    Public Property ID
    Public Property Name
    Public Property Email
End Class

