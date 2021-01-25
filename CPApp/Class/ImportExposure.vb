Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports System.IO
Public Class ImportExposure
    Inherits BaseImport

    Dim Parent As Object
    Dim FileNameFullPath As String
    Dim errMsgSB As New StringBuilder
    Dim myModel As ExposureModel = New ExposureModel

    Public Property ErrorMessage As String
        Get
            Return errMsgSB.ToString
        End Get
        Set(value As String)
            errMsgSB.Append(value)
        End Set
    End Property

    Dim DS As DataSet

    Dim NewSB As New StringBuilder
    Dim UpdateSB As New StringBuilder
    Dim CustomerSB As New StringBuilder
    Dim VendorSB As New StringBuilder
    Dim ImportStatusSB As New StringBuilder
    Dim myuser As UserInfo = UserInfo.getInstance
    Dim mylist As List(Of String)



    Public Sub New(parent As Object)
        Me.Parent = parent
        FileNameFullPath = parent.filename
        initdata()
    End Sub

    Public Function run() As Boolean

        mylist = New List(Of String)
        Dim myret As Boolean
        Parent.ProgressReport(1, "Preparing data..")

        If Not ConvertToTxtFile() Then

            'Return False
        End If
        If mylist.Count >= 1 Then
            Dim myflag As Boolean = False
            For Each myitem In mylist
                If Not BuildData(myitem) Then
                    myflag = True
                End If
                Thread.Sleep(100)
                Kill(myitem)               
            Next
            If myflag Then
                Return False
            End If
            Parent.ProgressReport(1, "Saving record..")
            If save() Then
                Parent.ProgressReport(1, "Import Done.")
                myret = True


            Else
                Parent.ProgressReport(1, "Found Error.")
            End If
        Else
            errMsgSB.Append("File is not correct.")
        End If
        

        'Copy Data
        ImportStatusSB.Append(String.Format("Copy & Update File{0}", vbCrLf))

        'If CopyTx() Then
        '    Parent.ProgressReport(1, "Import Done.")
        'Else
        '    Return myret
        'End If myret = True


        Return myret
    End Function

    'Private Sub AddUpdData(ws As Excel.Worksheet)
    '    'Check Content for column 3,6,7 
    '    Dim i As Integer = 13 'Data Content starting row
    '    Dim Check As Boolean = True
    '    Try
    '        While Check
    '            Dim myModel = New BufferStockModel With {.customername = validStr(ws.Cells(i, 3).value),
    '                                                     .customercode = validNumeric(ws.Cells(i, 4).value),
    '                                                     .vendorname = validStr(ws.Cells(i, 5).value),
    '                                                     .vendorcode = validNumeric(ws.Cells(i, 6).value),
    '                                                     .partnumber = validStr(ws.Cells(i, 7).value),
    '                                                     .description = validStr(ws.Cells(i, 8).value),
    '                                                     .projectname = validStr(ws.Cells(i, 9).value),
    '                                                     .leadtime = validStr(ws.Cells(i, 10).value),
    '                                                     .t2 = validStr(ws.Cells(i, 11).value),
    '                                                     .bufferqty = validNumeric(ws.Cells(i, 12).value),
    '                                                     .unit = validStr(ws.Cells(i, 13).value),
    '                                                     .unitprice = validNumeric(ws.Cells(i, 14).value)}

    '            ' If (Not IsNothing(ws.Cells(i, 7).value)) And (Not IsNothing(ws.Cells(i, 2).value)) And (Not IsNothing(ws.Cells(i, 3).value)) Then
    '            If myModel.partnumber <> "Null" And myModel.customercode <> "Null" And myModel.vendorcode <> "Null" Then
    '                ' If Not IsNothing(ws.Cells(i, 12).value) Then
    '                'If Not IsNumeric(ws.Cells(i, 12).value) Then
    '                If myModel.bufferqty = "Null" Then
    '                    Check = False
    '                    ImportStatusSB.Append(String.Format("-- Row {0} column 12 value {1}{2}", i, myModel.bufferqty, vbCrLf))
    '                Else
    '                    'Find Existing if not avail then create else update
    '                    Dim customercode As Long = CLng(myModel.customercode) 'ws.Cells(i, 4).value
    '                    Dim vendorcode As Long = CLng(myModel.vendorcode) 'ws.Cells(i, 6).value
    '                    Dim partno As String = validStr(myModel.partnumber) 'ws.Cells(i, 7).value)
    '                    Dim customername As String = validStr(myModel.customername) 'ws.Cells(i, 3).value)
    '                    Dim vendorname As String = validStr(myModel.vendorname) 'ws.Cells(i, 5).value)
    '                    Dim mykey(2) As Object
    '                    Dim mykey1(0) As Object
    '                    Dim mykey2(0) As Object
    '                    mykey(0) = customercode
    '                    mykey(1) = vendorcode
    '                    mykey(2) = partno
    '                    mykey1(0) = customercode
    '                    mykey2(0) = vendorcode

    '                    Dim result As DataRow
    '                    result = DS.Tables(0).Rows.Find(mykey)
    '                    If IsNothing(result) Then
    '                        'create
    '                        Dim dr = DS.Tables(0).NewRow
    '                        dr.Item("customercode") = customercode
    '                        dr.Item("vendorcode") = vendorcode
    '                        dr.Item("partno") = partno
    '                        DS.Tables(0).Rows.Add(dr)

    '                        'Check CustomerCode

    '                        result = DS.Tables(1).Rows.Find(mykey1)
    '                        If IsNothing(result) Then
    '                            'Create Customer
    '                            Dim custdr = DS.Tables(1).NewRow
    '                            custdr.Item("customercode") = customercode
    '                            custdr.Item("customername") = customername
    '                            DS.Tables(1).Rows.Add(custdr)
    '                            CustomerSB.Append(customercode & vbTab & customername & vbCrLf)
    '                        Else
    '                            'Update customer
    '                        End If
    '                        'Check Vendorcode
    '                        result = DS.Tables(2).Rows.Find(mykey2)
    '                        If IsNothing(result) Then
    '                            'Create Vendor
    '                            Dim vendordr = DS.Tables(2).NewRow
    '                            vendordr.Item("vendorcode") = vendorcode
    '                            vendordr.Item("vendorname") = vendorname
    '                            DS.Tables(2).Rows.Add(vendordr)
    '                            VendorSB.Append(vendorcode & vbTab & vendorname & vbCrLf)
    '                        Else
    '                            'Update vendor
    '                        End If

    '                        'customercode bigint,  vendorcode bigint,  partno character varying,
    '                        'description character varying,projectname character varying,leadtime integer,
    '                        't2vendor character varying,bufferqty integer,unit character varying,unitprice numeric,
    '                        'NewSB.Append(validStr(ws.Cells(i, 4).value) & vbTab &
    '                        '            validStr(ws.Cells(i, 6).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 7).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 8).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 9).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 10).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 11).value) & vbTab &
    '                        '             validNumeric(ws.Cells(i, 12).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 13).value) & vbTab &
    '                        '             validStr(ws.Cells(i, 14).value) & vbCrLf)
    '                        NewSB.Append(validStr(myModel.customercode) & vbTab &
    '                                    validStr(myModel.vendorcode) & vbTab &
    '                                     validStr(myModel.partnumber) & vbTab &
    '                                     validStr(myModel.description) & vbTab &
    '                                     validStr(myModel.projectname) & vbTab &
    '                                     validStr(myModel.leadtime) & vbTab &
    '                                     validStr(myModel.t2) & vbTab &
    '                                     validNumeric(myModel.bufferqty) & vbTab &
    '                                     validStr(myModel.unit) & vbTab &
    '                                     validStr(myModel.unitprice) & vbTab &
    '                                     validStr(myuser.Userid) & vbCrLf)
    '                    Else
    '                        'update() LeadTime, Buffer, UnitPrice
    '                        Dim flagUpdate As Boolean = False
    '                        If result.Item("leadtime").ToString <> myModel.leadtime Then
    '                            result.Item("leadtime") = myModel.leadtime
    '                            flagUpdate = True
    '                        End If
    '                        If result.Item("bufferqty").ToString <> myModel.bufferqty Then
    '                            result.Item("bufferqty") = myModel.bufferqty
    '                            flagUpdate = True
    '                        End If
    '                        If result.Item("unitprice").ToString <> myModel.unitprice Then
    '                            result.Item("unitprice") = myModel.unitprice
    '                            flagUpdate = True
    '                        End If
    '                        If flagUpdate Then
    '                            If UpdateSB.Length > 0 Then
    '                                UpdateSB.Append(",")
    '                            End If
    '                            UpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying]", myModel.customercode, myModel.vendorcode, myModel.partnumber, myModel.leadtime, myModel.bufferqty, myModel.unitprice))
    '                        End If
    '                    End If
    '                    'End If
    '                End If
    '            Else
    '                Check = False
    '                ImportStatusSB.Append(String.Format("-- Row {0} Customer Code:'{1}' Vendor Code:'{2}' PartNo:'{3}'{4}", i, ws.Cells(i, 4).value, ws.Cells(i, 6).value, ws.Cells(i, 7).value, vbCrLf))
    '            End If
    '            i = i + 1
    '        End While
    '    Catch ex As Exception
    '        ImportStatusSB.Append(String.Format("-- Row {0} Customer Code:'{1}' Vendor Code:'{2}' PartNo:'{3}'{4}", i, ws.Cells(i, 4).value, ws.Cells(i, 6).value, ws.Cells(i, 7).value, vbCrLf))
    '    End Try



    'End Sub


    'Private Function CopyTx() As Boolean
    '    Parent.ProgressReport(1, "Copy records.")

    '    Dim myret As Boolean = False
    '    Try
    '        If CustomerSB.Length > 0 Then
    '            Dim Sqlstr As String = String.Empty
    '            Dim message As String = String.Empty
    '            If CustomerSB.Length > 0 Then

    '                Sqlstr = "copy cp.customer(customercode,customername)  from stdin with null as 'Null';"
    '                message = DataAccess.Copy(Sqlstr, CustomerSB.ToString, myret)
    '                If Not myret Then
    '                    Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\errorcustomer.txt")
    '                        mystream.WriteLine(NewSB.ToString)
    '                    End Using
    '                    Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\errorcustomer.txt")
    '                    Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
    '                Else

    '                End If
    '            End If
    '        End If
    '        If VendorSB.Length > 0 Then
    '            Dim Sqlstr As String = String.Empty

    '            Dim message As String = String.Empty
    '            If VendorSB.Length > 0 Then

    '                Sqlstr = "copy cp.vendor(vendorcode,vendorname)  from stdin with null as 'Null';"
    '                message = DataAccess.Copy(Sqlstr, VendorSB.ToString, myret)
    '                If Not myret Then
    '                    Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\errorvendor.txt")
    '                        mystream.WriteLine(NewSB.ToString)
    '                    End Using
    '                    Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\errorvendor.txt")
    '                    Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
    '                End If
    '            End If
    '        End If
    '        If NewSB.Length > 0 Then
    '            Dim Sqlstr As String = String.Empty

    '            Dim message As String = String.Empty
    '            If NewSB.Length > 0 Then
    '                ImportStatusSB.Append(String.Format("-- Copy: {0}", vbCrLf))
    '                Sqlstr = "begin;set statement_timeout to 0;end;copy cp.bufferstock(customercode,vendorcode,partno,description,projectname,leadtime,t2vendor,bufferqty,unit,unitprice,modifiedby)  from stdin with null as 'Null';"
    '                message = DataAccess.Copy(Sqlstr, NewSB.ToString, myret)

    '                If Not myret Then
    '                    'save to text file first
    '                    Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
    '                        mystream.WriteLine(NewSB.ToString)
    '                    End Using
    '                    Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\error.txt")
    '                    Parent.ProgressReport(1, String.Format("Error Found. {0}", message))
    '                Else
    '                    ImportStatusSB.Append(String.Format("-- Copy: Done {0}", vbCrLf))
    '                End If
    '            End If
    '        End If

    '        If UpdateSB.Length > 0 Then
    '            Parent.ProgressReport(1, "Update BufferStock")
    '            ' myModel.customercode, myModel.vendorcode, myModel.partnumber, myModel.leadtime, myModel.bufferqty, myModel.unitprice
    '            Dim sqlstr = String.Format("update cp.bufferstock set leadtime= foo.leadtime::numeric,bufferqty = foo.bufferqty::integer,partno= foo.partno ,modifiedby='{0}'" &
    '                        " from (select * from array_to_set6(Array[{1}]) as tb (customercode character varying,vendorcode character varying, partno character varying, leadtime character varying, bufferqty character varying, unitprice character varying))foo where cp.bufferstock.customercode = foo.customercode::bigint and cp.bufferstock.vendorcode = foo.vendorcode::bigint and cp.bufferstock.partno = foo.partno;", myuser.Userid, UpdateSB.ToString)
    '            Dim ra As Long
    '            Dim errmsg As String = String.Empty
    '            ImportStatusSB.Append(String.Format("-- Update: {0}", vbCrLf))
    '            If Not DataAccess.ExecuteNonQuery(sqlstr, ra, Nothing) Then
    '                Parent.ProgressReport(1, "Update BufferStock" & "::" & errmsg)
    '            End If
    '            ImportStatusSB.Append(String.Format("-- Update: Done {0}", vbCrLf))
    '        End If
    '        myret = True
    '    Catch ex As Exception
    '        myret = False
    '        ErrorMessage = ex.Message
    '        ImportStatusSB.Append(String.Format("-- Found Error {0}", ErrorMessage, vbCrLf))
    '    End Try
    '    Return myret
    'End Function

    Private Function ConvertToTxtFile() As Boolean
        'Open ExcelFile
        'Check each worksheet
        'Check Rule :
        '   Header  : Cell(6,2) = "Add item updated date"
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
            For i = 2 To oWb.Worksheets.Count
                oWb.Worksheets(i).select()
                Dim ws As Excel.Worksheet = oWb.Worksheets(i)

                If Not IsNothing(ws.Cells(6, 2).value) Then
                    If ws.Cells(6, 2).value.ToString = "Add item updated date" Then
                        'Save ws to txt
                        Dim FileNameWrk = String.Format("{0}\{1}-{2}.txt", IO.Path.GetDirectoryName(FileNameFullPath), ws.Range("G2").Value, ws.Name)

                        oWb.SaveAs(FileNameWrk, FileFormat:=Excel.XlFileFormat.xlUnicodeText, CreateBackup:=False)
                        mylist.Add(FileNameWrk)
                        'Else
                        '    ErrorMessage = String.Format("{0} Wrong Template{1}", ws.Name, vbCrLf)
                    End If
                End If
            Next
            
            myret = True
        Catch ex As Exception
            myret = False
            ErrorMessage = String.Format("Found Error {0} {1}", ex.Message, vbCrLf)
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

            'Using mystream As New StreamWriter(Path.GetDirectoryName(FileNameFullPath) & "\ImportStatus.txt")
            '    mystream.WriteLine(ImportStatusSB.ToString)
            'End Using
            'Process.Start(Path.GetDirectoryName(FileNameFullPath) & "\ImportStatus.txt")
        End Try
        Return myret
    End Function

    Private Function BuildData(myitem As String) As Boolean
        Dim mylist As New List(Of String())
        'Dim myret As Boolean
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(myitem)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                Parent.ProgressReport(1, "Read Data")

                Do Until .EndOfData
                    myrecord = .ReadFields
                    mylist.Add(myrecord)
                Loop
                'Check Template File
                If mylist(6)(0) <> "Example" Then
                    errMsgSB.Append("Sorry wrong file template.")
                    Return False
                End If

                'Check VendorCode


                Parent.ProgressReport(1, "Build Record..")
                'FGCMMF
                'FGProjectFamily
                'FGBOM
                'FGBomTX
                'FGBOMUsage
                'FGBOMUsageTx

                Dim WSDate = (mylist(2)(0).Split(" "))(3)

                If WSDate <> Parent.period Then
                    errMsgSB.Append("Sorry wrong file date.")
                    Return False
                End If

                Dim Vendorcode As String = mylist(1)(6)
                Dim shortnamecp As String = mylist(1)(3)
                Dim colt As String = mylist(1)(9)
                Dim VendorNameSAP As String = String.Empty
                'Find Vendor SAP
                Dim pk0(0) As Object
                pk0(0) = Vendorcode
                Dim myresult = DS.Tables("VendorSAP").Rows.Find(pk0)
                If IsNothing(myresult) Then
                    'Show error message and quit.
                    errMsgSB.Append(String.Format("This Vendor Code '{0}' is not registered in our system. Please verify the Vendor Code in this file.", Vendorcode))
                    Return False
                End If
                VendorNameSAP = myresult.Item("vendorname").ToString.TrimEnd

                'Find VendorCP
                Dim pk(0) As Object
                pk(0) = Vendorcode
                myresult = DS.Tables("Vendor").Rows.Find(pk)
                If IsNothing(myresult) Then                    
                    Dim dr = DS.Tables("Vendor").NewRow
                    dr.Item("vendorcode") = Vendorcode
                    dr.Item("shortnamecp") = shortnamecp
                    dr.Item("vendorname") = VendorNameSAP
                    DS.Tables("Vendor").Rows.Add(dr)
                Else
                    If IsDBNull(myresult.Item("shortnamecp")) Then
                        myresult.Item("shortnamecp") = shortnamecp
                    Else
                        If myresult.Item("shortnamecp") <> shortnamecp Then
                            myresult.Item("shortnamecp") = shortnamecp
                        End If
                    End If
                    
                End If

                'Find FG Project Family
                Dim fgProjectFamilyid As Long = 0
                Dim wstab = IO.Path.GetFileNameWithoutExtension(myitem).Split("-")(1)
                Dim pk1(1) As Object
                pk1(0) = Vendorcode
                pk1(1) = wstab
                myresult = DS.Tables("FGProjectFamily").Rows.Find(pk1)
                If IsNothing(myresult) Then
                    Dim dr = DS.Tables("FGProjectFamily").NewRow
                    dr.Item("vendorcode") = Vendorcode
                    dr.Item("projectfamily") = wstab
                    dr.Item("colt") = colt
                    fgProjectFamilyid = dr.Item("id")
                    DS.Tables("FGProjectFamily").Rows.Add(dr)
                Else
                    fgProjectFamilyid = myresult.Item("id")
                End If


                For i = 7 To mylist.Count - 1  'BOM Iteration
                    If mylist(i)(1) = "" Then
                        Exit For
                    End If

                    'Find FGBOM FGProjectFamilyid & PartNumber
                    Dim divisionid As String = 0
                    Dim componentcategoryid As String = 0
                    Dim mainreasonid As String = 0

                    Dim division As String = mylist(i)(2)
                    Dim pkdiv(0) As Object
                    pkdiv(0) = division
                    Dim result = DS.Tables("Division").Rows.Find(pkdiv)
                    If Not IsNothing(result) Then
                        divisionid = result.Item("paramdtid")
                    Else
                        errMsgSB.Append(String.Format("Division ""{0}"" is not correct. Tab {1} Row {2}.{3}", mylist(i)(2), wstab, i + 1, vbCrLf))
                        'Exit For
                    End If

                    Dim mainreason As String = mylist(i)(12)
                    Dim pkmr(0) As Object
                    pkmr(0) = mainreason.ToString.Trim
                    result = DS.Tables("MainReason").Rows.Find(pkmr)
                    If Not IsNothing(result) Then
                        mainreasonid = result.Item("paramdtid")
                    Else
                        errMsgSB.Append(String.Format("Main Reason ""{0}"" is not correct. Tab {1} Row {2}.{3}", mylist(i)(12), wstab, i + 1, vbCrLf))
                        'Exit For
                    End If

                    Dim componentcategory As String = mylist(i)(3)
                    Dim pkcp(0) As Object
                    pkcp(0) = componentcategory
                    result = DS.Tables("ComponentCategory").Rows.Find(pkcp)
                    If Not IsNothing(result) Then
                        componentcategoryid = result.Item("paramdtid")
                    Else
                        errMsgSB.Append(String.Format("Component Category ""{0}"" is not correct. Tab {1} Row {2}.{3}", mylist(i)(3), wstab, i + 1, vbCrLf))
                        'Exit For
                    End If

                    Dim moq As Integer = 0
                    If IsNumeric(mylist(i)(8)) Then
                        moq = mylist(i)(8)
                    End If
                    Dim FGBOM As New FGBOM With {.partnumber = mylist(i)(4),
                                                 .fgprojectfamilyid = fgProjectFamilyid,
                                                 .createddate = mylist(i)(1),
                                                 .division = divisionid,
                                                 .componentcategoryid = componentcategoryid,
                                                 .componentdescription = mylist(i)(5),
                                                 .modelversion = mylist(i)(6),
                                                 .leadtime = mylist(i)(7),
                                                 .moq = moq,
                                                 .materialvendorname = mylist(i)(10),
                                                 .materialvendorlocation = mylist(i)(11),
                                                 .mainreasonid = mainreasonid,
                                                 .unit = mylist(i)(14)}



                    Dim fpk1(1) As Object
                    fpk1(0) = FGBOM.fgprojectfamilyid
                    fpk1(1) = FGBOM.partnumber

                    myresult = DS.Tables("FGBOM").Rows.Find(fpk1)
                    If IsNothing(myresult) Then
                        Dim dr1 As DataRow = DS.Tables("FGBOM").NewRow
                        dr1.Item("fgProjectFamilyid") = FGBOM.fgprojectfamilyid
                        dr1.Item("createddate") = FGBOM.createddate
                        dr1.Item("division") = FGBOM.division
                        dr1.Item("componentcategoryid") = FGBOM.componentcategoryid
                        dr1.Item("partnumber") = FGBOM.partnumber
                        dr1.Item("componentdescription") = FGBOM.componentdescription
                        dr1.Item("modelversion") = FGBOM.modelversion
                        dr1.Item("leadtime") = FGBOM.leadtime
                        dr1.Item("moq") = FGBOM.moq
                        dr1.Item("materialvendorname") = FGBOM.materialvendorname
                        dr1.Item("materialvendorlocation") = FGBOM.materialvendorlocation
                        dr1.Item("mainreasonid") = FGBOM.mainreasonid
                        dr1.Item("unit") = FGBOM.unit

                        DS.Tables("FGBOM").Rows.Add(dr1)
                        FGBOM.id = dr1.Item("id")
                    Else
                        FGBOM.id = myresult.Item("id")
                        myresult.Item("fgProjectFamilyid") = FGBOM.fgprojectfamilyid
                        myresult.Item("createddate") = FGBOM.createddate
                        myresult.Item("division") = FGBOM.division
                        myresult.Item("componentcategoryid") = FGBOM.componentcategoryid
                        myresult.Item("partnumber") = FGBOM.partnumber
                        myresult.Item("componentdescription") = FGBOM.componentdescription
                        myresult.Item("modelversion") = FGBOM.modelversion
                        myresult.Item("leadtime") = FGBOM.leadtime
                        myresult.Item("moq") = FGBOM.moq
                        myresult.Item("materialvendorname") = FGBOM.materialvendorname
                        myresult.Item("materialvendorlocation") = FGBOM.materialvendorlocation
                        myresult.Item("mainreasonid") = FGBOM.mainreasonid
                        myresult.Item("unit") = FGBOM.unit
                        Dim bs As New BindingSource
                        bs.DataSource = DS.Tables("FGBOMTX")
                        Dim myfilter = String.Format("fgbomid = {0} and txdate = '{1:yyyy-MM-dd}'", FGBOM.id, CDate(WSDate))
                        bs.Filter = myfilter
                        For Each drv In bs.List
                            bs.Remove(drv)
                        Next
                    End If

                    'FGBOMTX - Clear First based on Submit Date and FGProjectFamilyId (Vendorcode,ProjectFamily)
                    'The initial DataSet always blank
                    Dim dr = DS.Tables("FGBOMTX").NewRow
                    dr.Item("fgbomid") = FGBOM.id
                    dr.Item("txdate") = WSDate
                    If IsNumeric(mylist(i)(9)) Then
                        dr.Item("unitprice") = mylist(i)(9)
                    Else
                        dr.Item("unitprice") = 0
                    End If
                    If IsNumeric(mylist(i)(13)) Then
                        dr.Item("stock") = mylist(i)(13)
                    Else
                        dr.Item("stock") = 0
                    End If
                    DS.Tables("FGBOMTX").Rows.Add(dr)

                    For j = 15 To mylist(i).Length - 1 'CMMF Iteration
                        If mylist(2)(j) = "" Then
                            Exit For
                        End If

                        'Find CMMF
                        Dim cmmf As String = mylist(2)(j)
                        Dim modelnumber As String = mylist(1)(j)
                        Dim description As String = mylist(0)(j)

                        Dim fpk(0) As Object
                        fpk(0) = cmmf
                        myresult = DS.Tables("FGCMMF").Rows.Find(fpk)
                        If IsNothing(myresult) Then
                            dr = DS.Tables("FGCMMF").NewRow
                            dr.Item("cmmf") = cmmf
                            dr.Item("modelnumber") = modelnumber
                            dr.Item("description") = description
                            DS.Tables("FGCMMF").Rows.Add(dr)
                        End If

                        'Find FGBOMUsage
                        Dim FGBOMUsage = New FGBOMUsage With {.fgbomid = FGBOM.id,
                                                              .cmmf = cmmf}
                        Dim fpk2(1) As Object
                        fpk2(0) = FGBOMUsage.fgbomid
                        fpk2(1) = FGBOMUsage.cmmf

                        myresult = DS.Tables("FGBOMUsage").Rows.Find(fpk2)
                        If IsNothing(myresult) Then
                            dr = DS.Tables("FGBOMUsage").NewRow
                            dr.Item("fgbomid") = FGBOMUsage.fgbomid
                            dr.Item("cmmf") = FGBOMUsage.cmmf
                            DS.Tables("FGBOMUsage").Rows.Add(dr)
                            FGBOMUsage.id = dr.Item("id")
                        Else
                            FGBOMUsage.id = myresult.Item("id")
                            'Delete Existing FGBOMUsageTX
                            Dim bs As New BindingSource
                            bs.DataSource = DS.Tables("FGBOMUsageTX")
                            Dim myfilter = String.Format("fgbomusageid = {0} and txdate = '{1:yyyy-MM-dd}'", FGBOMUsage.id, CDate(WSDate))
                            bs.Filter = myfilter
                            For Each drv In bs.List
                                bs.Remove(drv)
                            Next
                        End If



                        'FGBOMUsageTx
                        Dim quantity As String = mylist(i)(j)
                        If quantity <> "" Then


                            dr = DS.Tables("FGBOMUsageTx").NewRow
                            dr.Item("fgbomusageid") = FGBOMUsage.id
                            dr.Item("txdate") = WSDate
                            dr.Item("quantity") = mylist(i)(j)
                            DS.Tables("FGBOMUsageTx").Rows.Add(dr)
                        End If
                    Next
                Next
            End With
        End Using
        Return errMsgSB.Length = 0
    End Function

    Private Sub initdata()
        '-FGProjectFamily
        '-FGCMMF
        '-Vendor
        '-FGBOM
        'FGBOMUsage
        'FGBomTX
        'FGBOMUsageTx
        DS = myModel.GetDataSet
        DS.Tables(0).TableName = "FGProjectFamily"
        DS.Tables(1).TableName = "FGCMMF"
        DS.Tables(2).TableName = "Vendor"
        DS.Tables(3).TableName = "FGBOM"
        DS.Tables(4).TableName = "FGBOMUsage"
        DS.Tables(5).TableName = "FGBOMTX"
        DS.Tables(6).TableName = "FGBOMUsageTX"
        DS.Tables(7).TableName = "Division"
        DS.Tables(8).TableName = "MainReason"
        DS.Tables(9).TableName = "ComponentCategory"
        DS.Tables(10).TableName = "VendorSAP"

        'FGProjectFamily
        Dim pk(1) As DataColumn
        pk(0) = DS.Tables(0).Columns("vendorcode")
        pk(1) = DS.Tables(0).Columns("projectfamily")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1

        'FGCMMF
        Dim pk1(0) As DataColumn
        pk1(0) = DS.Tables(1).Columns("cmmf")
        DS.Tables(1).PrimaryKey = pk1

        'Vendor
        Dim pk2(0) As DataColumn
        pk2(0) = DS.Tables(2).Columns("vendorcode")  'Vendor
        DS.Tables(2).PrimaryKey = pk2

        'FGBOM
        Dim pk3(1) As DataColumn
        pk3(0) = DS.Tables(3).Columns("fgprojectfamilyid")
        pk3(1) = DS.Tables(3).Columns("partnumber")
        DS.Tables(3).PrimaryKey = pk3
        DS.Tables(3).Columns("id").AutoIncrement = True
        DS.Tables(3).Columns("id").AutoIncrementSeed = -1
        DS.Tables(3).Columns("id").AutoIncrementStep = -1


        'FGBOMUsage
        Dim pk4(1) As DataColumn
        pk4(0) = DS.Tables(4).Columns("fgbomid")  '
        pk4(1) = DS.Tables(4).Columns("cmmf")
        DS.Tables(4).PrimaryKey = pk4
        DS.Tables(4).Columns("id").AutoIncrement = True
        DS.Tables(4).Columns("id").AutoIncrementSeed = -1
        DS.Tables(4).Columns("id").AutoIncrementStep = -1


        'FGBOMTX
        Dim pk5(0) As DataColumn
        pk5(0) = DS.Tables(5).Columns("id")
        DS.Tables(5).PrimaryKey = pk5
        DS.Tables(5).Columns("id").AutoIncrement = True
        DS.Tables(5).Columns("id").AutoIncrementSeed = -1
        DS.Tables(5).Columns("id").AutoIncrementStep = -1

        'FGBOMUsageTX
        Dim pk6(0) As DataColumn
        pk6(0) = DS.Tables(6).Columns("id")
        DS.Tables(6).PrimaryKey = pk6
        DS.Tables(6).Columns("id").AutoIncrement = True
        DS.Tables(6).Columns("id").AutoIncrementSeed = -1
        DS.Tables(6).Columns("id").AutoIncrementStep = -1

        'Division
        Dim pk7(0) As DataColumn
        pk7(0) = DS.Tables(7).Columns("cvalue")
        DS.Tables(7).PrimaryKey = pk7

        'Main Reason
        Dim pk8(0) As DataColumn
        pk8(0) = DS.Tables(8).Columns("cvalue")
        DS.Tables(8).PrimaryKey = pk8

        'Component Category
        Dim pk9(0) As DataColumn
        pk9(0) = DS.Tables(9).Columns("cvalue")
        DS.Tables(9).PrimaryKey = pk9

        'VendorSAP
        Dim pk10(0) As DataColumn
        pk10(0) = DS.Tables(10).Columns("vendorcode")  'Vendor
        DS.Tables(10).PrimaryKey = pk10


        'Create Relation HD-DT
        Dim rel As DataRelation
        Dim hcol As DataColumn
        Dim dcol As DataColumn

        'FGProjectFamily - FGBOM
        hcol = DS.Tables("FGProjectFamily").Columns("id") 'id in table FGProjectFamily
        dcol = DS.Tables("FGBOM").Columns("fgprojectfamilyid") 'headerid in table FGBOM
        rel = New DataRelation("PF_BOM", hcol, dcol)
        DS.Relations.Add(rel)

        'FGBOM - FGBOMTX
        hcol = DS.Tables("FGBOM").Columns("id") 'id in table header
        dcol = DS.Tables("FGBOMTX").Columns("fgbomid") 'headerid in table detail
        rel = New DataRelation("B_BTX", hcol, dcol)
        DS.Relations.Add(rel)

        'FGBOM - FGBOMUsage
        hcol = DS.Tables("FGBOM").Columns("id") 'id in table header
        dcol = DS.Tables("FGBOMUsage").Columns("fgbomid") 'headerid in table detail
        rel = New DataRelation("B_BU", hcol, dcol)
        DS.Relations.Add(rel)

        'FGBOMUsage - FGBOMUsageTX
        hcol = DS.Tables("FGBOMUsage").Columns("id") 'id in table header
        dcol = DS.Tables("FGBOMUsageTx").Columns("fgbomusageid") 'headerid in table detail
        rel = New DataRelation("BU_BUTX", hcol, dcol)
        DS.Relations.Add(rel)

    End Sub


    Public Function save() As Boolean
        Dim myret As Boolean = False
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    For i = 0 To ds2.Tables.Count - 1
                        If ds2.Tables(i).Rows.Count > 0 Then
                            DS.Tables(i).Merge(ds2.Tables(i))
                        End If
                    Next
                    DS.AcceptChanges()
                    'MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If

        Return myret
    End Function

    Public Function Save(ByVal mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        If myModel.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function

End Class


Public Class FGBOM
    Public Property id As String
    Public Property fgprojectfamilyid As String
    Public Property createddate As String
    Public Property componentcategoryid As String
    Public Property partnumber As String
    Public Property componentdescription As String
    Public Property modelversion As String
    Public Property leadtime As String
    Public Property moq As String
    Public Property materialvendorname As String
    Public Property materialvendorlocation As String
    Public Property mainreasonid As String
    Public Property unit As String
    Public Property division As String
End Class

Public Class FGBOMUsage
    Public Property id As String
    Public Property fgbomid As String
    Public Property cmmf As String
End Class