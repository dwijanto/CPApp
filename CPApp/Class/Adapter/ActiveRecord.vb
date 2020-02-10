Imports System.Text
Public MustInherit Class ActiveRecord
    Implements IActiveRecord

    Public tableName As String
    Public primarykey As String

    Public Sub New()

    End Sub


    Public Function deleteAll(Optional condition As Object = Nothing) As Integer Implements IActiveRecord.deleteAll
        Throw New NotImplementedException
    End Function

    Public Function Find() As Object Implements IActiveRecord.Find
        Dim sqlstr = String.Format("select * from {0} ", tableName)
        Return sqlstr
    End Function

    Public Function delete(Optional condition As Object = Nothing) As Object Implements IActiveRecord.delete
        Throw New NotImplementedException
    End Function

    Protected Function findByCondition(Optional condition As Object = Nothing, Optional other As Object = Nothing) As Object Implements IActiveRecord.findByCondition
        Dim sb As New StringBuilder
        sb.Append(Find())
        If Not IsNothing(condition) Then
            Dim isb As New StringBuilder
            sb.Append(" where ")
            If TypeOf condition Is Hashtable Then
                For Each obj As DictionaryEntry In condition
                    If isb.Length > 0 Then
                        isb.Append(" and ")
                    End If
                    isb.Append(String.Format("{0}::character varying = '{1}'", obj.Key, obj.Value))
                Next
                sb.Append(isb.ToString)
            ElseIf TypeOf condition Is Array Then

                For Each obj In condition
                    If isb.Length > 0 Then
                        isb.Append(",")
                    End If
                    isb.Append(String.Format("'{0}'", obj))
                Next
                sb.Append(String.Format(" {0}::character varying in ({1})", primarykey, isb.ToString))
            Else
                sb.Append(String.Format(" {0}::character varying = '{1}'", primarykey, condition))
            End If
        End If

        If Not IsNothing(other) Then
            sb.Append(String.Format(" {0}", other))
        End If

        Dim ds As DataSet
        ds = DataAccess.GetDataSet(sb.ToString, CommandType.Text, Nothing)
        Return ds
    End Function

    Public Function findOne(Optional condition As Object = Nothing) As Object Implements IActiveRecord.findOne
        Return findByCondition(condition, "limit 1")
    End Function

    Public Function insert(Optional condition As Object = Nothing, Optional other As Object = Nothing) As Object Implements IActiveRecord.insert
        Throw New NotImplementedException
    End Function

    Public Function findAll(Optional condition As Object = Nothing, Optional other As Object = Nothing) As Object Implements IActiveRecord.findAll
        Return findByCondition(condition, other)
    End Function
End Class
