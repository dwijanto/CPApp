Imports System.Text

'-- Table: sales._user

'-- DROP TABLE sales._user;

'CREATE TABLE sales._user
'(
'  id bigserial NOT NULL,
'  userid character varying,
'  username character varying,
'  email character varying,
'  isadmin boolean NOT NULL DEFAULT false,
'  isactive boolean NOT NULL DEFAULT true,
'  CONSTRAINT userpk PRIMARY KEY (id)
')
'WITH (
'  OIDS=FALSE
');
'ALTER TABLE sales._user
'  OWNER TO postgres;
'GRANT ALL ON TABLE sales._user TO postgres;
'GRANT ALL ON TABLE sales._user TO public;

'-- Index: sales.useridx1

'-- DROP INDEX sales.useridx1;

'CREATE UNIQUE INDEX useridx1
'  ON sales._user
'  USING btree
'  (userid COLLATE pg_catalog."default");


Public Class UserModel
    Implements IModel

    Public ReadOnly Property FilterField
        Get
            Return "[userid] like '%{0}%' or [username] like '%{0}%' or [email] like '%{0}%'"
        End Get
    End Property

    Public ReadOnly Property TableName As String Implements IModel.tablename
        Get
            Return "cp.user"
        End Get
    End Property

    Public ReadOnly Property SortField As String Implements IModel.sortField
        Get
            Return "username"
        End Get
    End Property

    Public Function LoadData(ByRef DS As DataSet) As Boolean Implements IModel.LoadData
        Dim SB As New StringBuilder
        SB.Append(String.Format("select u.* from {0} u  order by {1};", TableName, SortField))
        Dim sqlstr = SB.ToString
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
            'Update
            Dim sqlstr = "cp.sp_update_user"
            dataadapter.UpdateCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", DataRowVersion.Original))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "userid", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "username", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "email", DataRowVersion.Current))
            dataadapter.UpdateCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "isactive", DataRowVersion.Current))
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_insert_user"
            dataadapter.InsertCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "userid", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "username", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.String, 0, "email", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Boolean, 0, "isactive", DataRowVersion.Current))
            dataadapter.InsertCommand.Parameters.Add(factory.CreateParameter("", DbType.Int64, 0, "id", ParameterDirection.InputOutput))
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "cp.sp_delete_user"
            dataadapter.DeleteCommand = factory.CreateCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add(factory.CreateParameter("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id"))
            dataadapter.DeleteCommand.Parameters(0).Direction = ParameterDirection.Input
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

    Public Function ADDUPDUserManager(ByVal myDATA As List(Of ADPrincipalContext))
        Dim myret As Boolean
        Dim mgrId As Long
        Try
            Dim factory = DataAccess.factory
            Dim params(8) As System.Data.IDbDataParameter
            If myDATA.Count > 1 Then
                'User with ID Manager
                params(0) = factory.CreateParameter("iuserid", myDATA(1).Userid)
                params(1) = factory.CreateParameter("iusername", myDATA(1).DisplayName)
                params(2) = factory.CreateParameter("icompany", myDATA(1).Company)
                params(3) = factory.CreateParameter("iemail", myDATA(1).Email)
                params(4) = factory.CreateParameter("ititle", myDATA(1).Title)
                params(5) = factory.CreateParameter("iemployeenumber", myDATA(1).EmployeeNumber)
                params(6) = factory.CreateParameter("idepartment", myDATA(1).Department)
                params(7) = factory.CreateParameter("icountry", myDATA(1).Country)
                params(8) = factory.CreateParameter("ilocation", myDATA(1).Location)
                mgrId = DataAccess.ExecuteScalar("cp.sp_addupduser", CommandType.StoredProcedure, params)
            End If
            'User Part
            params(0) = factory.CreateParameter("iuserid", myDATA(0).Userid)
            params(1) = factory.CreateParameter("iusername", myDATA(0).DisplayName)
            params(2) = factory.CreateParameter("icompany", myDATA(0).Company)
            params(3) = factory.CreateParameter("iemail", myDATA(0).Email)
            params(4) = factory.CreateParameter("ititle", myDATA(0).Title)
            params(5) = factory.CreateParameter("iemployeenumber", myDATA(0).EmployeeNumber)
            params(6) = factory.CreateParameter("idepartment", myDATA(0).Department)
            params(7) = factory.CreateParameter("icountry", myDATA(0).Country)
            params(8) = factory.CreateParameter("ilocation", myDATA(0).Location)
            DataAccess.ExecuteScalar("cp.sp_addupduser", CommandType.StoredProcedure, params)
            myret = True
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            myret = False
        End Try
        Return myret
    End Function
End Class
