Public MustInherit Class BaseImport
    Public Function validStr(ByVal p As String) As String
        If p = "" Then
            Return "Null"
        End If
        Return String.Format("{0}", p.Replace(vbTab, "").Replace(Chr(13), "").Replace(Chr(10), "").Replace("\", "\\"))
    End Function

    Public Function validNumeric(ByVal p As String) As String
        If p = "" Then
            Return "Null"
        End If
        Return String.Format("{0}", p.Replace(",", ""))
    End Function

    Public Function validLong(ByVal p As String) As String
        If p = "" Then
            Return "Null"
        End If
        Return String.Format("{0}", p.Replace(",", ""))
    End Function

    Public Function validint(ByVal p As String) As String
        If p = "" Then
            Return "Null"
        End If
        p = p.Replace(",", "")
        If Not IsNumeric(p) Then
            Return "Null"
        End If
        Return String.Format("{0}", p.Replace(",", ""))
    End Function

    Public Function validDate(ByVal p As String) As String

        If p = "" Then
            Return "Null"
        End If
        If IsDate(p) Then
            Return String.Format("'{0:yyyy-MM-dd}'", CDate(p))
        Else
            Return String.Format("'{0:yyyy-MM-dd}'", p)
        End If

        
    End Function

    Public Function validSAPDate(ByVal p As String) As String
        If p = "00.00.0000" Then
            Return "Null"
        End If
        Dim myvalue() = p.Split(".")
        Return String.Format("'{2}-{1}-{0}'", myvalue(0), myvalue(1), myvalue(2))
    End Function
End Class
