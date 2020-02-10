Imports Npgsql
Imports System.IO

Public Class PostgreSQLFactory
    Inherits DbFactory

    Private myadapter As NpgsqlDataAdapter
    Private CopyIn1 As NpgsqlCopyIn

    Dim conninfo As New NpgsqlConnectionStringBuilder

    Public Overrides ReadOnly Property GetUserName As String
        Get
            conninfo.ConnectionString = ConnectionString
            Return conninfo.UserName
        End Get
    End Property

    Public Overrides ReadOnly Property GetPassword As String
        Get
            Return "admin"
        End Get
    End Property
    Public Overrides ReadOnly Property GetHostName As String
        Get
            conninfo.ConnectionString = ConnectionString
            Return conninfo.Host
        End Get
    End Property
    Public Overrides ReadOnly Property GetDataBaseName As String
        Get
            conninfo.ConnectionString = ConnectionString
            Return conninfo.Database
        End Get
    End Property

    Public Overrides ReadOnly Property ConnectionString As String
        Get
            'Return String.Format("host=hon13-0002L;port=5432;database=LogisticDb20150120;CommandTimeout=10000;TimeOut=1024;Userid=admin;Password=admin;")
            'Return String.Format("host=localhost;port=5432;database=LogisticDb20150120;CommandTimeout=10000;TimeOut=1024;Userid=admin;Password=admin;")
            Return String.Format(My.Settings.Conn1)
        End Get
    End Property

    Public Overloads Overrides Function CreateAdapter() As IDbDataAdapter
        myadapter = New NpgsqlDataAdapter
        AddHandler myadapter.RowUpdated, AddressOf onRowUpdate
        Return myadapter
    End Function

    Public Overrides Function Update(tablename As Object) As Integer
        ' Return myadapter.Update(DirectCast(tablename, DataTable))
        Return myadapter.Update(tablename)
    End Function

    Public Overrides Function CreateAdapter(commandText As String) As IDbDataAdapter
        Return New NpgsqlDataAdapter(New NpgsqlCommand(commandText))
    End Function

    Public Overrides Function CreateCommand(commandText As String, connection As IDbConnection) As IDbCommand
        Return New NpgsqlCommand(commandText, connection)
    End Function

    Public Overrides Function CreateConnection() As IDbConnection
        Dim myret As IDbConnection
        myret = New NpgsqlConnection(ConnectionString)
        Return myret
    End Function

    Public Overloads Overrides Function CreateParameter() As IDbDataParameter
        Return New NpgsqlParameter
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, value As Object) As IDbDataParameter
        Return New NpgsqlParameter(name, value)
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        Return parm

    End Function
    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourceColumn As String) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourceColumn
        Return parm
    End Function
    Public Overrides Function GetParameterValue(parameter As Object) As Object
        Debug.Assert(parameter <> Nothing)
        If (parameter Is Nothing) Then Return Nothing
        Return CType(parameter, NpgsqlParameter).Value
    End Function

    Private Sub onRowUpdate(sender As Object, e As NpgsqlRowUpdatedEventArgs)
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub



    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourcecolumn As String, sourceversion As DataRowVersion) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourcecolumn
        parm.SourceVersion = sourceversion
        Return parm
    End Function

    Public Overloads Overrides Function CreateParameter(name As String, type As DbType, size As Integer, sourcecolumn As String, paramInputOutput As ParameterDirection) As IDbDataParameter
        Dim parm As NpgsqlParameter = New NpgsqlParameter
        parm.ParameterName = name
        parm.DbType = type
        parm.Size = size
        parm.SourceColumn = sourcecolumn
        parm.Direction = paramInputOutput
        Return parm
    End Function

    Public Overloads Overrides Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString.Replace("\", "\\")))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(ConnectionString)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & vbCrLf & ex.Detail & vbCrLf & ex.Where
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ex.Where
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function


End Class

