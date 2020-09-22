Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports System.IO
Public Class ImportCPDBDemand
    Inherits BaseImport

    Dim Parent As Object
    Dim FileNameFullPath As String
    'Public ErrorMessage As String
    Public Property errMsgSB As New StringBuilder

    Dim DS As DataSet

    Dim DBDemandSB As New StringBuilder


    Dim ImportStatusSB As New StringBuilder
    Dim myuser As UserInfo = UserInfo.getInstance

    Public Sub New(parent As Object)
        Me.Parent = parent
        FileNameFullPath = parent.filename
        DS = parent.ds
    End Sub

    Public Function run() As Boolean
        Dim mylist As New List(Of String())
        Dim myret As Boolean
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(FileNameFullPath)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(";")
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                Parent.ProgressReport(1, "Read Data")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    mylist.Add(myrecord)
                Loop
                'Check Template File
                If mylist(0)(0) <> "Year Week" Then
                    errMsgSB.Append("Sorry wrong file template.")
                    Return False
                End If

                Parent.ProgressReport(1, "Build Record..")
                For i = 1 To mylist.Count - 1

                    Dim mymodel As DBDemandModel = New DBDemandModel With {.cmmf = mylist(i)(4),
                                                                           .vendorcode = mylist(i)(1),
                                                                           .qty = mylist(i)(11),
                                                                           .yearweek = mylist(i)(14)}
                    If mymodel.qty <> "" Then
                        If mymodel.qty <> 0 Then
                            DBDemandSB.Append(mymodel.cmmf & vbTab &
                                         mymodel.vendorcode & vbTab &
                                         mymodel.qty & vbTab &
                                         mymodel.yearweek & vbCrLf)
                        End If
                       
                    End If
                Next


                'Copy

                myret = CopyTx()

            End With
        End Using
        Return myret
    End Function

    Private Function CopyTx() As Boolean
        Parent.ProgressReport(1, "Copy records.")

        Dim myret As Boolean = False
        Try

            If DBDemandSB.Length > 0 Then
                Dim Sqlstr As String = String.Empty

                Dim message As String = String.Empty

                ImportStatusSB.Append(String.Format("-- Copy: {0}", vbCrLf))
                Sqlstr = "begin;set statement_timeout to 0;end;delete from cp.cpdbdemand;select setval('cp.cpdbdemand_id_seq',1,false);copy cp.cpdbdemand(cmmf,vendorcode,qty,yearweek)  from stdin with null as 'Null';"
                message = DataAccess.Copy(Sqlstr, DBDemandSB.ToString, myret)

                If Not myret Then
                    'save to text file first
                    'Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
                    '    mystream.WriteLine(NewSB.ToString)
                    'End Using
                    'Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
                    'Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
                    errMsgSB.Append(message)
                Else
                    ImportStatusSB.Append(String.Format("-- Copy: Done {0}", vbCrLf))
                End If

            End If
            myret = True
        Catch ex As Exception
            myret = False
            errMsgSB.Append(ex.Message)
            ImportStatusSB.Append(String.Format("-- Found Error {0}", errMsgSB.ToString, vbCrLf))
        End Try
        Return myret
    End Function
End Class
