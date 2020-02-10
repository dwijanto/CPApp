Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports System.IO
Public Class ImportBufferStock
    Inherits BaseImport

    Dim Parent As Object
    Dim FileNameFullPath As String
    Public ErrorMessage As String
    Dim DS As DataSet

    Dim NewSB As New StringBuilder
    Dim UpdateSB As New StringBuilder
    Dim CustomerSB As New StringBuilder
    Dim VendorSB As New StringBuilder
    Dim ImportStatusSB As New StringBuilder
    Dim myuser As UserInfo = UserInfo.getInstance

    Public Sub New(parent As Object)
        Me.Parent = parent
        FileNameFullPath = parent.filename
        DS = parent.ds
    End Sub

    Public Function run() As Boolean
        'Open ExcelFile
        'Check each worksheet
        'Check Rule :
        '   Header  : Cell(12,3) = "Customer Name"
        '   Contents: cell(row,12) is numeric
        Dim myret As Boolean = False
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr

        Try
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.Visible = False
            oXl.DisplayAlerts = False
            oWb = oXl.Workbooks.Open(FileNameFullPath)
            'Check Each Worksheet
            For Each ws As Excel.Worksheet In oWb.Worksheets
                If Not IsNothing(ws.Cells(12, 3).value) Then
                    If ws.Cells(12, 3).value.ToString = "Customer Name" Then
                        ImportStatusSB.Append(String.Format("{0}{1}", ws.Name, vbCrLf))
                        AddUpdData(ws)
                    Else
                        ImportStatusSB.Append(String.Format("{0} Wrong Template{1}", ws.Name, vbCrLf))
                    End If
                End If
            Next
            'Copy Data
            ImportStatusSB.Append(String.Format("Copy & Update File{0}", vbCrLf))

            If CopyTx() Then
                Parent.ProgressReport(1, "Import Done.")
            Else
                Return myret
            End If
            myret = True
        Catch ex As Exception
            myret = False
            ErrorMessage = ex.Message
            ImportStatusSB.Append(String.Format("--Found Error {0} {1}", ErrorMessage, vbCrLf))
        Finally
            oXl.Quit()
            ExportToExcelFile.releaseComObject(oSheet)
            ExportToExcelFile.releaseComObject(oWb)
            ExportToExcelFile.releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                ExportToExcelFile.EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

            Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\ImportStatus.txt")
                mystream.WriteLine(ImportStatusSB.ToString)
            End Using
            Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\ImportStatus.txt")
        End Try

        Return myret
    End Function

    Private Sub AddUpdData(ws As Excel.Worksheet)
        'Check Content for column 3,6,7 
        Dim i As Integer = 13 'Data Content starting row
        Dim Check As Boolean = True
        Try
            While Check
                Dim myModel = New BufferStockModel With {.customername = validStr(ws.Cells(i, 3).value),
                                                         .customercode = validNumeric(ws.Cells(i, 4).value),
                                                         .vendorname = validStr(ws.Cells(i, 5).value),
                                                         .vendorcode = validNumeric(ws.Cells(i, 6).value),
                                                         .partnumber = validStr(ws.Cells(i, 7).value),
                                                         .description = validStr(ws.Cells(i, 8).value),
                                                         .projectname = validStr(ws.Cells(i, 9).value),
                                                         .leadtime = validStr(ws.Cells(i, 10).value),
                                                         .t2 = validStr(ws.Cells(i, 11).value),
                                                         .bufferqty = validNumeric(ws.Cells(i, 12).value),
                                                         .unit = validStr(ws.Cells(i, 13).value),
                                                         .unitprice = validNumeric(ws.Cells(i, 14).value)}

                ' If (Not IsNothing(ws.Cells(i, 7).value)) And (Not IsNothing(ws.Cells(i, 2).value)) And (Not IsNothing(ws.Cells(i, 3).value)) Then
                If myModel.partnumber <> "Null" And myModel.customercode <> "Null" And myModel.vendorcode <> "Null" Then
                    ' If Not IsNothing(ws.Cells(i, 12).value) Then
                    'If Not IsNumeric(ws.Cells(i, 12).value) Then
                    If myModel.bufferqty = "Null" Then
                        Check = False
                        ImportStatusSB.Append(String.Format("-- Row {0} column 12 value {1}{2}", i, myModel.bufferqty, vbCrLf))
                    Else
                        'Find Existing if not avail then create else update
                        Dim customercode As Long = CLng(myModel.customercode) 'ws.Cells(i, 4).value
                        Dim vendorcode As Long = CLng(myModel.vendorcode) 'ws.Cells(i, 6).value
                        Dim partno As String = validStr(myModel.partnumber) 'ws.Cells(i, 7).value)
                        Dim customername As String = validStr(myModel.customername) 'ws.Cells(i, 3).value)
                        Dim vendorname As String = validStr(myModel.vendorname) 'ws.Cells(i, 5).value)
                        Dim mykey(2) As Object
                        Dim mykey1(0) As Object
                        Dim mykey2(0) As Object
                        mykey(0) = customercode
                        mykey(1) = vendorcode
                        mykey(2) = partno
                        mykey1(0) = customercode
                        mykey2(0) = vendorcode

                        Dim result As DataRow
                        result = DS.Tables(0).Rows.Find(mykey)
                        If IsNothing(result) Then
                            'create
                            Dim dr = DS.Tables(0).NewRow
                            dr.Item("customercode") = customercode
                            dr.Item("vendorcode") = vendorcode
                            dr.Item("partno") = partno
                            DS.Tables(0).Rows.Add(dr)

                            'Check CustomerCode

                            result = DS.Tables(1).Rows.Find(mykey1)
                            If IsNothing(result) Then
                                'Create Customer
                                Dim custdr = DS.Tables(1).NewRow
                                custdr.Item("customercode") = customercode
                                custdr.Item("customername") = customername
                                DS.Tables(1).Rows.Add(custdr)
                                CustomerSB.Append(customercode & vbTab & customername & vbCrLf)
                            Else
                                'Update customer
                            End If
                            'Check Vendorcode
                            result = DS.Tables(2).Rows.Find(mykey2)
                            If IsNothing(result) Then
                                'Create Vendor
                                Dim vendordr = DS.Tables(2).NewRow
                                vendordr.Item("vendorcode") = vendorcode
                                vendordr.Item("vendorname") = vendorname
                                DS.Tables(2).Rows.Add(vendordr)
                                VendorSB.Append(vendorcode & vbTab & vendorname & vbCrLf)
                            Else
                                'Update vendor
                            End If

                            'customercode bigint,  vendorcode bigint,  partno character varying,
                            'description character varying,projectname character varying,leadtime integer,
                            't2vendor character varying,bufferqty integer,unit character varying,unitprice numeric,
                            'NewSB.Append(validStr(ws.Cells(i, 4).value) & vbTab &
                            '            validStr(ws.Cells(i, 6).value) & vbTab &
                            '             validStr(ws.Cells(i, 7).value) & vbTab &
                            '             validStr(ws.Cells(i, 8).value) & vbTab &
                            '             validStr(ws.Cells(i, 9).value) & vbTab &
                            '             validStr(ws.Cells(i, 10).value) & vbTab &
                            '             validStr(ws.Cells(i, 11).value) & vbTab &
                            '             validNumeric(ws.Cells(i, 12).value) & vbTab &
                            '             validStr(ws.Cells(i, 13).value) & vbTab &
                            '             validStr(ws.Cells(i, 14).value) & vbCrLf)
                            NewSB.Append(validStr(myModel.customercode) & vbTab &
                                        validStr(myModel.vendorcode) & vbTab &
                                         validStr(myModel.partnumber) & vbTab &
                                         validStr(myModel.description) & vbTab &
                                         validStr(myModel.projectname) & vbTab &
                                         validStr(myModel.leadtime) & vbTab &
                                         validStr(myModel.t2) & vbTab &
                                         validNumeric(myModel.bufferqty) & vbTab &
                                         validStr(myModel.unit) & vbTab &
                                         validStr(myModel.unitprice) & vbTab &
                                         validStr(myuser.Userid) & vbCrLf)
                        Else
                            'update() LeadTime, Buffer, UnitPrice
                            Dim flagUpdate As Boolean = False
                            If result.Item("leadtime").ToString <> myModel.leadtime Then
                                result.Item("leadtime") = myModel.leadtime
                                flagUpdate = True
                            End If
                            If result.Item("bufferqty").ToString <> myModel.bufferqty Then
                                result.Item("bufferqty") = myModel.bufferqty
                                flagUpdate = True
                            End If
                            If result.Item("unitprice").ToString <> myModel.unitprice Then
                                result.Item("unitprice") = myModel.unitprice
                                flagUpdate = True
                            End If
                            If flagUpdate Then
                                If UpdateSB.Length > 0 Then
                                    UpdateSB.Append(",")
                                End If
                                UpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying]", myModel.customercode, myModel.vendorcode, myModel.partnumber, myModel.leadtime, myModel.bufferqty, myModel.unitprice))
                            End If
                        End If
                        'End If
                    End If
                Else
                    Check = False
                    ImportStatusSB.Append(String.Format("-- Row {0} Customer Code:'{1}' Vendor Code:'{2}' PartNo:'{3}'{4}", i, ws.Cells(i, 4).value, ws.Cells(i, 6).value, ws.Cells(i, 7).value, vbCrLf))
                End If
                i = i + 1
            End While
        Catch ex As Exception
            ImportStatusSB.Append(String.Format("-- Row {0} Customer Code:'{1}' Vendor Code:'{2}' PartNo:'{3}'{4}", i, ws.Cells(i, 4).value, ws.Cells(i, 6).value, ws.Cells(i, 7).value, vbCrLf))
        End Try
        
       

    End Sub


    Private Function CopyTx() As Boolean
        Parent.ProgressReport(1, "Copy records.")
       
        Dim myret As Boolean = False
        Try
            If CustomerSB.Length > 0 Then
                Dim Sqlstr As String = String.Empty
                Dim message As String = String.Empty
                If CustomerSB.Length > 0 Then

                    Sqlstr = "copy cp.customer(customercode,customername)  from stdin with null as 'Null';"
                    message = DataAccess.Copy(Sqlstr, CustomerSB.ToString, myret)
                    If Not myret Then
                        Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\errorcustomer.txt")
                            mystream.WriteLine(NewSB.ToString)
                        End Using
                        Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\errorcustomer.txt")
                        Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
                    Else

                    End If
                End If
            End If
            If VendorSB.Length > 0 Then
                Dim Sqlstr As String = String.Empty

                Dim message As String = String.Empty
                If VendorSB.Length > 0 Then

                    Sqlstr = "copy cp.vendor(vendorcode,vendorname)  from stdin with null as 'Null';"
                    message = DataAccess.Copy(Sqlstr, VendorSB.ToString, myret)
                    If Not myret Then
                        Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\errorvendor.txt")
                            mystream.WriteLine(NewSB.ToString)
                        End Using
                        Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\errorvendor.txt")
                        Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
                    End If
                End If
            End If
            If NewSB.Length > 0 Then
                Dim Sqlstr As String = String.Empty

                Dim message As String = String.Empty
                If NewSB.Length > 0 Then
                    ImportStatusSB.Append(String.Format("-- Copy: {0}", vbCrLf))
                    Sqlstr = "begin;set statement_timeout to 0;end;copy cp.bufferstock(customercode,vendorcode,partno,description,projectname,leadtime,t2vendor,bufferqty,unit,unitprice,modifiedby)  from stdin with null as 'Null';"
                    message = DataAccess.Copy(Sqlstr, NewSB.ToString, myret)

                    If Not myret Then
                        'save to text file first
                        Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
                            mystream.WriteLine(NewSB.ToString)
                        End Using
                        Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
                        Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
                    Else
                        ImportStatusSB.Append(String.Format("-- Copy: Done {0}", vbCrLf))
                    End If
                End If
            End If

            If UpdateSB.Length > 0 Then
                Parent.ProgressReport(1, "Update BufferStock")
                ' myModel.customercode, myModel.vendorcode, myModel.partnumber, myModel.leadtime, myModel.bufferqty, myModel.unitprice
                Dim sqlstr = String.Format("update cp.bufferstock set leadtime= foo.leadtime::numeric,bufferqty = foo.bufferqty::integer,partno= foo.partno ,modifiedby='{0}'" &
                            " from (select * from array_to_set6(Array[{1}]) as tb (customercode character varying,vendorcode character varying, partno character varying, leadtime character varying, bufferqty character varying, unitprice character varying))foo where cp.bufferstock.customercode = foo.customercode::bigint and cp.bufferstock.vendorcode = foo.vendorcode::bigint and cp.bufferstock.partno = foo.partno;", myuser.Userid, UpdateSB.ToString)
                Dim ra As Long
                Dim errmsg As String = String.Empty
                ImportStatusSB.Append(String.Format("-- Update: {0}", vbCrLf))
                If Not DataAccess.ExecuteNonQuery(sqlstr, ra, Nothing) Then
                    Parent.ProgressReport(1, "Update BufferStock" & "::" & errmsg)
                End If
                ImportStatusSB.Append(String.Format("-- Update: Done {0}", vbCrLf))
            End If
            myret = True
        Catch ex As Exception
            myret = False
            ErrorMessage = ex.Message
            ImportStatusSB.Append(String.Format("-- Found Error {0}", ErrorMessage, vbCrLf))
        End Try        
        Return myret
    End Function




End Class
