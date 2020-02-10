Imports System.Reflection
Imports System.IO

Public Class DataAccess
    Public Shared factory As DbFactory = DbFactory.GetInstancePostgreSQLFactory
    'Public Shared factory As DbFactory = DbFactory.GetInstanceSQLFactory

    Public Shared ReadOnly Property GetHostName
        Get
            Return factory.GetHostName
        End Get
    End Property

    Public Shared ReadOnly Property GetDataBaseName
        Get
            Return factory.GetDataBaseName
        End Get
    End Property

    Public Shared ReadOnly Property GetUserName
        Get
            Return factory.GetUserName
        End Get
    End Property

    Public Shared ReadOnly Property GetPassword
        Get
            Return factory.GetPassword
        End Get
    End Property

    Public Shared Function isAdmin(ByVal userid As String) As Boolean
        Dim myret As Boolean = False
        Dim myParams(0) As IDbDataParameter
        myParams(0) = factory.CreateParameter("userid", userid)
        Return ExecuteScalar("cp.sp_isadmin", CommandType.StoredProcedure, myParams)
    End Function


    Public Shared Function LogLogin(ByVal UserInfo As UserInfo)
        Dim myret As Boolean = False
        'Dim applicationname As String = "Product Request Apps"
        'Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        'Dim computername As String = My.Computer.Name
        Dim myParams(3) As IDbDataParameter
        myParams(0) = factory.CreateParameter("iapplicationname", UserInfo.ApplicationName)
        myParams(1) = factory.CreateParameter("iuserid", UserInfo.Userid)
        myParams(2) = factory.CreateParameter("iusername", UserInfo.Username)
        myParams(3) = factory.CreateParameter("icomputername", UserInfo.computerName)
        Return ExecuteScalar("sp_insertlogonhistory", CommandType.StoredProcedure, myParams)
    End Function

    Public Shared Function ExecuteNonQuery(ByVal procedureName As String, ByVal cmdType As CommandType, ByVal ParamArray parameters() As IDbDataParameter) As Integer
        Debug.Assert(procedureName <> Nothing)
        Dim myret As Object = Nothing
        Using connection As IDbConnection = factory.CreateConnection()
            connection.Open()
            Dim command As IDbCommand = factory.CreateCommand(procedureName, connection)
            command.Connection = connection
            command.CommandType = cmdType
            If (Not IsNothing(parameters)) Then
                Dim p As IDbDataParameter
                For Each p In parameters
                    command.Parameters.Add(p)
                Next
            End If
            myret = command.ExecuteNonQuery
        End Using
        Return myret
    End Function

    Public Shared Function ExecuteScalar(ByVal procedureName As String, ByVal cmdType As CommandType, ByVal ParamArray parameters() As IDbDataParameter) As Object
        Debug.Assert(procedureName <> Nothing)
        Dim myret As Object = Nothing
        Using connection As IDbConnection = factory.CreateConnection()
            connection.Open()
            Dim command As IDbCommand = factory.CreateCommand(procedureName, connection)
            command.Connection = connection
            command.CommandType = cmdType
            If (Not IsNothing(parameters)) Then
                Dim p As IDbDataParameter
                For Each p In parameters
                    command.Parameters.Add(p)
                Next
            End If
            myret = command.ExecuteScalar
        End Using
        Return myret
    End Function

    Public Shared Function GetDataSet(ByVal procedureName As String, ByVal cmdType As CommandType, ByVal ParamArray parameters() As IDbDataParameter) As DataSet
        Dim DS As DataSet = New DataSet
        Debug.Assert(procedureName <> Nothing)

        Using connection As IDbConnection = factory.CreateConnection()
            Dim DataAdapter As IDbDataAdapter = factory.CreateAdapter
            connection.Open()
            Dim command As IDbCommand = factory.CreateCommand(procedureName, connection)
            command.Connection = connection
            command.CommandType = cmdType
            DataAdapter.SelectCommand = command

            If (Not IsNothing(parameters)) Then
                Dim p As IDbDataParameter
                For Each p In parameters
                    command.Parameters.Add(p)
                Next
            End If
            DataAdapter.Fill(DS)
            Return DS
        End Using
    End Function

    Public Shared Function GetDataTable(ByVal procedureName As String, ByVal cmdType As CommandType, ByVal ParamArray parameters() As IDbDataParameter) As DataTable
        Dim DS As DataSet = New DataSet
        Debug.Assert(procedureName <> Nothing)

        Using connection As IDbConnection = factory.CreateConnection()
            Dim DataAdapter As IDbDataAdapter = factory.CreateAdapter
            connection.Open()
            Dim command As IDbCommand = factory.CreateCommand(procedureName, connection)
            command.Connection = connection
            command.CommandType = cmdType
            DataAdapter.SelectCommand = command

            If (Not IsNothing(parameters)) Then
                Dim p As IDbDataParameter
                For Each p In parameters
                    command.Parameters.Add(p)
                Next
            End If
            DataAdapter.Fill(DS)
            Return DS.Tables(0)
        End Using
    End Function

    Public Shared Function ExecuteReader(Of T)(ByVal procedureName As String, ByVal cmdType As CommandType, ByVal handler As ReadEventHandler(Of T), ByVal ParamArray parameters() As IDbDataParameter) As T
        Debug.Assert(handler <> Nothing)
        Using connection As IDbConnection = factory.CreateConnection()
            connection.Open()
            Dim command As IDbCommand = factory.CreateCommand(procedureName, connection)
            command.Connection = connection
            command.CommandType = cmdType

            If (parameters Is Nothing = False) Then
                Dim p As IDbDataParameter
                For Each p In parameters
                    command.Parameters.Add(p)
                Next
            End If

            Dim reader As IDataReader = command.ExecuteReader()
            Return handler.Invoke(reader)
        End Using

    End Function

    'populate model, return List of model
    Public Shared Function OnReadAnyList(Of T As New)(ByVal reader As IDataReader) As List(Of T)
        If (reader Is Nothing) Then Return New List(Of T)()
        Dim list As List(Of T) = New List(Of T)()
        While (reader.Read())
            list.Add(OnReadAny(Of T)(reader))
        End While

        Return list
    End Function

    ' read the public properties and use these to read field names
    Public Shared Function OnReadAny(Of T As New)(ByVal reader As IDataReader) As T

        Dim genericType As Type = GetType(T)

        Dim properties() As PropertyInfo = genericType.GetProperties(BindingFlags.Instance Or BindingFlags.Public)

        Dim prop As PropertyInfo
        Dim model As T = New T()

        For Each prop In properties

            Try
                Dim columnName As String = GetColumnName(prop)

                If (reader(columnName).Equals(System.DBNull.Value) = False) Then
                    Dim value As Object = reader(columnName)
                    prop.SetValue(model, value, Nothing)
                End If

            Catch ex As Exception
                Debug.WriteLine("Couldn't write " + prop.Name)
            End Try
        Next
        Return model
    End Function

    Private Shared Function GetColumnName(ByVal prop As PropertyInfo) As String

        Debug.Assert(prop Is Nothing = False)
        If (prop Is Nothing) Then Return ""

        Dim attributes() As Object = prop.GetCustomAttributes(True)
        Dim attr As Object
        For Each attr In attributes
            'If (TypeOf attr Is MYVB.BusinessObjects.ColumnNameAttribute) Then
            '    Debug.WriteLine("Uses ColumnNameAttribute")
            '    Return CType(attr, MYVB.BusinessObjects.ColumnNameAttribute).ColumnName
            'End If
        Next

        Return prop.Name

    End Function

    Public Shared Function SafeRead(Of T)(ByVal field As T, ByVal reader As IDataReader, ByVal name As String) As T
        If (reader(name).Equals(System.DBNull.Value) = False) Then
            Dim result As Object = reader(name)
            Return CType(Convert.ChangeType(result, GetType(T)), T)
        Else
            Return field
        End If
    End Function

    Public Shared Function Copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        Return factory.Copy(sqlstr, InputString, result)
    End Function
   
End Class
