Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Delegate Sub GenerateReportDelegate(ByRef sender As Object, ByRef e As eventargs)

Public Class GenerateReport

    Private FileNameFullPath As String = String.Empty
    Private sqlstr As String = String.Empty
    Public Parent As Object = Nothing
    Private ErrorMessage As New StringBuilder
    Dim status As Boolean = False
    Private myTemplate As String = String.Empty
    Public Property QueryList As List(Of QueryWorkSheet)
    Private DataSheet As Integer = 1
    Private Location As String = "A1"
    Private GenerateReportCallback As GenerateReportDelegate
    Private RawDataCallback As GenerateReportDelegate
    Private FieldNames As Boolean = True
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function
    Public sqlstrphoto As String
    Public Title As String

    Public Sub New()

    End Sub


    Public Sub New(ByRef Parent As Object, ByVal sqlstr As String, ByRef FileNameFullPath As String, ByVal GenerateReportCallback As GenerateReportDelegate, ByVal RawDataCallBack As GenerateReportDelegate, Optional DataSheet As Integer = 1, Optional ByVal template As String = "\template\ExcelTemplate.xltx", Optional Location As String = "A1", Optional FieldNames As Boolean = True, Optional sqlstrphoto As String = "", Optional title As String = "")
        Me.FileNameFullPath = FileNameFullPath
        Me.sqlstr = sqlstr
        Me.Parent = Parent
        Me.myTemplate = template
        Me.DataSheet = DataSheet
        Me.Location = Location
        Me.RawDataCallback = RawDataCallBack
        Me.GenerateReportCallback = GenerateReportCallback
        Me.FieldNames = FieldNames
        Me.sqlstrphoto = sqlstrphoto
        Me.Title = title
    End Sub

    Public Sub Run()
        Dim sw As New Stopwatch
        sw.Start()
        Parent.progressReport(1, "Generating Report...")
        Parent.progressReport(6, "Marques..")

        status = GenerateReport()
        sw.Stop()
        If status Then
            Parent.ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            If MsgBox("File name: " & FileNameFullPath & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(FileNameFullPath)
            End If
        Else
            Parent.progressReport(1, ErrorMessage.ToString)
        End If
        Parent.progressReport(5, "Continuous")
    End Sub

    Private Function GenerateReport() As Boolean
        Dim myRet As Boolean = False
        Dim oxl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim hwnd As System.IntPtr = Nothing

        Try
            Parent.progressReport(1, "Create Object")
            oxl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oxl.Hwnd

            oxl.DisplayAlerts = False
            Parent.progressReport(1, "Opening Template")
            owb = oxl.Workbooks.Open(Application.StartupPath & myTemplate)
            oxl.Visible = False

            Parent.Progressreport(1, "Creating Worksheet")

            If IsNothing(QueryList) Then
                'Only One RawData Query
                owb.Worksheets(DataSheet).select()
                oSheet = owb.Worksheets(DataSheet)
                If sqlstr.Length > 0 Then
                    Parent.ProgressReport(1, "Get records..")
                    FillWorksheet(oSheet, sqlstr)
                    Dim orange = oSheet.Range(Location)
                    Dim lastrow = GetLastRow(oxl, oSheet, orange)
                    If lastrow > 1 Then
                        Parent.ProgressReport(1, "Raw Data Callback..")
                        RawDataCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    RawDataCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'More than One RawDataQuery
                For i = 0 To QueryList.Count - 1
                    Dim myquery = CType(QueryList(i), QueryWorkSheet)
                    owb.Worksheets(myquery.DataSheet).select()
                    oSheet = owb.Worksheets(myquery.DataSheet)
                    oSheet.Name = myquery.SheetName
                    Parent.ProgressReport(1, "Get records..")

                    FillWorksheet(oSheet, myquery.Sqlstr)
                    Parent.ProgressReport(1, "Get Last Row..")
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oxl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification                       
                        Parent.ProgressReport(2, "Raw Data Callback..")
                        RawDataCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next

            End If
            Parent.ProgressReport(1, "Generate Report Callback..")
            GenerateReportCallback.Invoke(owb, New EventArgs)

            Dim success As Boolean = False
            Dim intI As Integer = 0
            While (Not success)
                intI += 1
                For i = 0 To owb.Connections.Count - 1
                    Try
                        owb.Connections(1).Delete()
                        success = True
                    Catch ex As Exception
                        Parent.ProgressReport(1, String.Format("Remove Connections failed.. {0}", intI))
                    End Try
                Next
            End While

            Parent.ProgressReport(1, "Saving File ..." & FileNameFullPath)

            If FileNameFullPath.Contains("xlsm") Then
                owb.SaveAs(FileNameFullPath, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                owb.SaveAs(FileNameFullPath)
            End If
            myRet = True
        Catch ex As Exception
            ErrorMessage.Append(ex.Message)
        Finally
            Try
                oxl.Quit()
                releaseComObject(oSheet)
                releaseComObject(owb)
                releaseComObject(oxl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
        End Try

        Return myRet
    End Function

    Private Sub FillWorksheet(oSheet As Excel.Worksheet, sqlstr As String)
        Dim oExCon As String = My.Settings.oExCon ' My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        oExCon = oExCon.Insert(oExCon.Length, "UID=" & DataAccess.GetUserName & ";PWD=" & DataAccess.GetPassword)
        Dim oRange As Excel.Range
        oRange = oSheet.Range(Location)
        With oSheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            .CommandText = sqlstr
            .FieldNames = FieldNames
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With
        oRange = Nothing
        oRange = oSheet.Range(Location)
        oRange.Select()
        oSheet.Application.Selection.autofilter()
        oSheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Function GetLastRow(oxl As Excel.Application, oSheet As Excel.Worksheet, orange As Excel.Range) As Object
        Dim lastrow As Long = 1
        oxl.ScreenUpdating = False
        Try
            lastrow = oSheet.Cells.Find("*", orange, , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
        Catch ex As Exception
        End Try
        Return lastrow
        oxl.ScreenUpdating = True
    End Function

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Public Sub MakeBorder(oSheet As Excel.Worksheet, RowStart As Integer, ColumnStart As Integer, RowEnd As Integer, ColumnEnd As Integer, sWeight As Integer, Linestyle As Integer)
        With oSheet.Range(oSheet.Cells(RowStart, ColumnStart), oSheet.Cells(RowEnd, ColumnEnd))
            .Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Linestyle
            .Borders(Excel.XlBordersIndex.xlEdgeTop).Weight = sWeight
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Linestyle
            .Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = sWeight
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Linestyle
            .Borders(Excel.XlBordersIndex.xlEdgeLeft).Weight = sWeight
            .Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Linestyle
            .Borders(Excel.XlBordersIndex.xlEdgeRight).Weight = sWeight
            .VerticalAlignment = Excel.Constants.xlTop
        End With
    End Sub

    Public Sub MakeInsideBorder(oSheet As Excel.Worksheet, RowStart As Integer, ColumnStart As Integer, RowEnd As Integer, ColumnEnd As Integer, sWeight As Integer, Linestyle As Integer)
        With oSheet.Range(oSheet.Cells(RowStart, ColumnStart), oSheet.Cells(RowEnd, ColumnEnd))
            If RowStart <> RowEnd Then
                .Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Linestyle
                .Borders(Excel.XlBordersIndex.xlInsideHorizontal).Weight = sWeight
            End If
            .Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Linestyle
            .Borders(Excel.XlBordersIndex.xlInsideVertical).Weight = sWeight
            .VerticalAlignment = Excel.Constants.xlTop
        End With
    End Sub

    Public Function FirstOf(myrecord As BindingSource, MyIndexfield As Integer) As Boolean
        Dim drv As DataRowView = myrecord.Current
        Dim i As Integer
        Dim sbOri As New StringBuilder
        Dim sbCompare As New StringBuilder

        'Check First Record
        If myrecord.Position = 0 Then
            FirstOf = True
            Exit Function
        End If


        For i = 0 To MyIndexfield
            sbOri.Append(String.Format("{0}", IIf(IsDBNull(drv.Row.Item(i)), "", drv.Row.Item(i).ToString.Trim)))
        Next
        myrecord.MovePrevious()
        drv = myrecord.Current
        For i = 0 To MyIndexfield
            sbCompare.Append(String.Format("{0}", IIf(IsDBNull(drv.Row.Item(i)), "", drv.Row.Item(i).ToString.Trim)))
        Next
        FirstOf = False
        If sbOri.ToString <> sbCompare.ToString Then
            FirstOf = True
        End If
        myrecord.MoveNext()
    End Function

    Public Function LastOf(myrecord As BindingSource, MyIndexfield As Integer) As Boolean
        Dim drv As DataRowView = myrecord.Current
        Dim i As Integer
        Dim sbOri As New StringBuilder
        Dim sbCompare As New StringBuilder

        'Check Last Record
        If myrecord.Position = myrecord.Count - 1 Then
            LastOf = True
            myrecord.MovePrevious()
            Exit Function
        End If

        For i = 0 To MyIndexfield
            sbOri.Append(String.Format("{0}", IIf(IsDBNull(drv.Row.Item(i)), "", drv.Row.Item(i).ToString.Trim)))
        Next

        myrecord.MoveNext()
        drv = myrecord.Current

        For i = 0 To MyIndexfield
            sbCompare.Append(String.Format("{0}", IIf(IsDBNull(drv.Row.Item(i)), "", drv.Row.Item(i).ToString.Trim)))
        Next
        LastOf = False
        If sbOri.ToString <> sbCompare.ToString Then
            LastOf = True
        End If
        myrecord.MovePrevious()

    End Function


    Sub FormatRawData()

    End Sub

    Sub GenerateReportWithRange(ByRef sender As Object, ByRef e As System.EventArgs)
        Dim oWb = DirectCast(sender, Excel.Workbook)
        Dim oxl As Excel.Application = oWb.Parent
        Dim oBC = oWb.Worksheets(1)
        oBC.Cells.EntireColumn.AutoFit()
        Parent.ProgressReport(1, "Data Formatting..")
        oBC.Range("H:H").ColumnWidth = 20
        oBC.Range("R:R").NumberFormat = "dd-MMM-yyyy"
        oBC.Range("A7:AV1000").VerticalAlignment = Excel.Constants.xlTop
        oBC.Range("N:N").ColumnWidth = 30


        oBC.Range("Q:AB").ColumnWidth = 11
        oBC.Range("S:S").ColumnWidth = 8
        oBC.Range("U:U").ColumnWidth = 8
        oBC.Range("W:W").ColumnWidth = 8
        oBC.Range("Y:Y").ColumnWidth = 8
        oBC.Range("AA:AA").ColumnWidth = 8
        oBC.Range("G:G").ColumnWidth = 15
        oBC.Cells.EntireColumn.WrapText = True



        'assign Photo
        Parent.ProgressReport(1, "Get Image Path info...")
        Dim DT As DataTable = DataAccess.GetDataTable(sqlstrphoto, CommandType.Text)
        Parent.ProgressReport(1, "Put Picture(s)...")
        oBC.Rows("7:" & DT.Rows.Count + 6).Interior.ColorIndex = 36
        Call MakeBorder(oBC, 7, 1, DT.Rows.Count + 6, 49, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)
        Call MakeInsideBorder(oBC, 7, 1, DT.Rows.Count + 6, 49, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)


        Call PutPicture(DT, oxl, oBC)
       
        Parent.ProgressReport(1, "AutoFit ....")
        Application.DoEvents()
        oBC.Cells.EntireColumn.AutoFit()
        Parent.ProgressReport(1, "Interior Pattern ....")
        oBC.Range(oBC.Columns("AX:AX"), oBC.Columns("AX:AX").End(Excel.XlDirection.xlToRight)).Interior.Pattern = Excel.Constants.xlNone
        Parent.ProgressReport(1, "Set Range ....")

        oBC.Range("H:H").ColumnWidth = 20
        oBC.Range("R:R").NumberFormat = "dd-MMM-yyyy"
        oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
        oBC.Range("N:N").ColumnWidth = 30
        oBC.Range("Q:AB").ColumnWidth = 11
        oBC.Range("S:S").ColumnWidth = 8
        oBC.Range("U:U").ColumnWidth = 8
        oBC.Range("W:W").ColumnWidth = 8
        oBC.Range("Y:Y").ColumnWidth = 8
        oBC.Range("AA:AA").ColumnWidth = 8
        oBC.Range("G:G").ColumnWidth = 15
        Parent.ProgressReport(1, "Wrap Text ....")        
        oBC.Cells.EntireColumn.WrapText = True
        oBC.Range("G1:I1").WrapText = False
        oBC.Cells(1, 8) = Date.Today
        oBC.Cells(1, 9) = Title
        oWb.Worksheets(1).Select()
        oxl.DisplayAlerts = False
        oBC.PageSetup.PrintArea = "$A$1:$AJ" & DT.Rows.Count + 6
    End Sub

    Sub FormatRawDataWithRange()
        'Throw New NotImplementedException
    End Sub

    Sub GenerateReportNoRange(ByRef sender As Object, ByRef e As System.EventArgs)       
        Dim oWb = DirectCast(sender, Excel.Workbook)
        Dim oxl As Excel.Application = oWb.Parent
        Dim oBC = oWb.Worksheets(1)
        Application.DoEvents()

        oBC.Cells.EntireColumn.AutoFit()
        Parent.ProgressReport(1, "Data Formatting...")
        oBC.Range("H:H").ColumnWidth = 20
        oBC.Range("Q:Q").NumberFormat = "dd-MMM-yyyy"
        oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
        oBC.Range("M:M").ColumnWidth = 30
        oBC.Range("P:AA").ColumnWidth = 11
        oBC.Range("R:R").ColumnWidth = 8
        oBC.Range("T:T").ColumnWidth = 8
        oBC.Range("V:V").ColumnWidth = 8
        oBC.Range("X:X").ColumnWidth = 8
        oBC.Range("Z:Z").ColumnWidth = 8
        oBC.Range("G:G").ColumnWidth = 15
        oBC.Cells.EntireColumn.WrapText = True



        'assign Photo
        Parent.ProgressReport(1, "Get Image Path info...")
        Dim DT As DataTable = DataAccess.GetDataTable(sqlstrphoto, CommandType.Text)
        Parent.ProgressReport(1, "Put Picture(s)...")
        Dim iRow As Long       
        iRow = 7      
        oBC.Rows("7:" & DT.Rows.Count + 6).Interior.ColorIndex = 36


        Call MakeBorder(oBC, 7, 1, DT.Rows.Count + 6, 48, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)
        Call MakeInsideBorder(oBC, 7, 1, DT.Rows.Count + 6, 48, Excel.XlBorderWeight.xlThin, Excel.XlLineStyle.xlContinuous)

        Call PutPicture(DT, oxl, oBC)

        Parent.ProgressReport(1, "AutoFit ....")
        Application.DoEvents()
        oBC.Cells.EntireColumn.AutoFit()
        Parent.ProgressReport(1, "Interior Pattern ....")
        oBC.Range(oBC.Columns("AW:AW"), oBC.Columns("AW:AW").End(Excel.XlDirection.xlToRight)).Interior.Pattern = Excel.Constants.xlNone
        Parent.ProgressReport(1, "Set Range ....")

        oBC.Range("H:H").ColumnWidth = 20
        oBC.Range("Q:Q").NumberFormat = "dd-MMM-yyyy"
        oBC.Range("A7:AP1000").VerticalAlignment = Excel.Constants.xlTop
        oBC.Range("M:M").ColumnWidth = 30

        oBC.Range("M:M").ColumnWidth = 30
        oBC.Range("P:AA").ColumnWidth = 11
        oBC.Range("R:R").ColumnWidth = 8
        oBC.Range("T:T").ColumnWidth = 8
        oBC.Range("V:V").ColumnWidth = 8
        oBC.Range("X:X").ColumnWidth = 8
        oBC.Range("Z:Z").ColumnWidth = 8
        oBC.Range("G:G").ColumnWidth = 15
        Parent.ProgressReport(1, "Wrap Text ....")
        oBC.Cells.EntireColumn.WrapText = True
        oBC.Range("G1:I1").WrapText = False
        oBC.Cells(1, 8) = Date.Today
        oBC.Cells(1, 9) = Title
        oWb.Worksheets(1).Select()
        oxl.DisplayAlerts = False
    End Sub

    Private Sub PutPicture(DT As DataTable, oxl As Excel.Application, oBC As Excel.Worksheet)
        Dim oWF As Excel.WorksheetFunction       
        oWF = oxl.WorksheetFunction
        Dim iRow As Long
        Dim Span As Double
        Dim j As Integer
        Dim Pict As Object
        Dim myfirstRow As Long
        Dim mylastRow As Long

        Dim myWD As Double
        Dim ganjil As Long
        Dim reccount As Long
        Dim BS As New BindingSource
        BS.DataSource = DT
        Dim myfilename As String
        Dim myImagePath = My.Settings.MyImagePath
        iRow = 7
        For i = 0 To DT.Rows.Count - 1
            Parent.ProgressReport(1, String.Format("Put Picture(s).. {0}/{1}", reccount, DT.Rows.Count - 1))
            If i = DT.Rows.Count - 1 Then
                Debug.Print("Last Record")
            End If
            If IsDBNull(DT.Rows(i).Item("projectname")) AndAlso IsDBNull(DT.Rows(i).Item("rangename")) Then

            Else
                If FirstOf(BS, 1) Then
                    myfirstRow = iRow
                End If
                If i = DT.Rows.Count - 1 Then
                    Debug.Print("Last Record")
                End If
                If LastOf(BS, 1) Then
                    mylastRow = iRow
                    ganjil = ganjil + 1
                    If ganjil Mod 2 = 0 Then
                        oBC.Rows(myfirstRow & ":" & mylastRow).Interior.ColorIndex = 40
                    End If
                    If Not (IsDBNull(DT.Rows(i).Item("imagepath"))) Then

                        myfilename = myImagePath & "\" & Path.GetFileName(Trim(DT.Rows(i).Item("imagepath")))

                        If File.Exists(myfilename) Then
                            If oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height < 100 Then
                                'Reserve height for cell
                                If myfirstRow = mylastRow Then
                                    oBC.Cells(myfirstRow, 2).RowHeight = 100
                                Else
                                    Span = (100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height) / (mylastRow + 1 - myfirstRow)
                                    For j = 0 To (mylastRow - myfirstRow)
                                        If j = mylastRow - myfirstRow Then
                                            Span = 100 - oBC.Range("H" & myfirstRow & ":H" & mylastRow).Height
                                        End If
                                        oBC.Cells(myfirstRow + j, 2).RowHeight = oBC.Cells(myfirstRow + j, 2).RowHeight + Span
                                    Next j
                                End If
                            End If
                            oBC.Range("H" & myfirstRow & ":H" & mylastRow).Merge()
                            'Print
                            Dim PictHT As Single, PictWD As Single
                            Dim RngHT As Single, RngWD As Single
                            Dim myScale As Double, MyL As Double, MyT As Double
                            Parent.ProgressReport(1, String.Format("Put Picture(s)...{0}/{1} {2}", reccount, DT.Rows.Count - 1, myfilename))

                            Pict = oBC.Pictures.Insert(myfilename)
                            Pict.Placement = 1

                            'do pic scale
                            Application.DoEvents()

                            PictHT = Pict.Height
                            PictWD = Pict.Width

                            'Size of Template Picture
                            RngHT = 98
                            RngWD = 98

                            myScale = oWF.Min(RngHT / PictHT, RngWD / PictWD)

                            With Pict
                                .Width = PictWD * myScale
                                .Height = PictHT * myScale

                            End With
                            'Put Picture Location
                            With oBC.Range("H" & myfirstRow)
                                MyT = .Top
                                MyL = .Left
                            End With

                            'Center Pict H & V
                            If PictHT > PictWD Then
                                MyL = MyL + (RngWD - Pict.Width) / 2
                            Else
                                MyT = MyT + (RngHT - Pict.Height) / 2
                            End If
                            myWD = 0

                            myWD = (oBC.Range("B" & iRow).Offset(0, 1).Left) - (oBC.Range("B" & iRow).Left) - 2

                            With Pict
                                .Top = MyT + 1
                                .Left = MyL + 2
                            End With
                            Pict = Nothing
                        End If
                    End If
                End If
            End If
            iRow = iRow + 1
            reccount = reccount + 1
            BS.MoveNext()
        Next
    End Sub

End Class
