Imports Microsoft.Office.Interop
Imports System.Threading

Public Class FormGenerateReportExposureRawData
    Dim myThread As New System.Threading.Thread(AddressOf DoQuery)
    Public Shared myForm As FormGenerateReportExposureRawData
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New ReportExposureController
    Private ViewAllData As Boolean = False
    Dim PeriodBS As BindingSource

    Public Sub New(ByVal ViewAllData As Boolean)
        InitializeComponent()
        Me.ViewAllData = ViewAllData
    End Sub
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myForm = New FormGenerateReportExposureRawData
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportExposureRawData
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex >= 0 Then
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            SaveFileDialog1.FileName = String.Format("DataBase.xlsx", Date.Today)
            Dim drv As DataRowView = ComboBox1.SelectedItem
            Dim criteria As String = String.Empty
            If ComboBox1.SelectedIndex > 0 Then
                criteria = String.Format("and btx.txdate = '{0:yyyy-MM-dd}'", drv.Item("txdate"))
            End If
            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                'Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myModel.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\ExcelTemplate.xltx")
                Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.myModel.GetSQLSTRReport(criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 1, "\templates\ExcelTemplate.xltx")              
                ' myReport.run(Me, New EventArgs)
                myReport.ExternalFilename = "CP.xlsx,FG.xlsx"
                myReport.mytemplate2 = "\templates\ExposureTemplate01.xltx"
                myReport.CreateExternalData(Me, New EventArgs)
            End If
        Else
            MessageBox.Show("Please select from period.")
        End If

    End Sub

    Private Sub FormatReport(ByRef sender As Object, ByRef e As EventArgs)
        'Throw New NotImplementedException
        ' Debug.Print("ok")
        Dim mye As ExternalDataArgs = TryCast(e, ExternalDataArgs)
        If Not IsNothing(mye) Then
            Dim oXl As Excel.Application = Nothing
            Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
            oXl = owb.Parent
            owb.Worksheets(1).select()
            '.Connection = {"OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Junk\PivotTable\DataBaseWRK.xlsx;Mode=Share Deny Write;Extended Properties=""HDR=YES;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=37;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"}
            With owb.Connections("DataBase").OLEDBConnection
                .BackgroundQuery = False
                .CommandText = {"RawData$"}
                .Connection = {String.Format("OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source={0};Mode=Share Deny Write;Extended Properties=""HDR=YES;"";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=37;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", mye.DBFile)}
                .RefreshOnFileOpen = False
                .SavePassword = False
                .SourceConnectionFile = ""
                .ServerCredentialsMethod = Excel.XlCredentialsMethod.xlCredentialsMethodIntegrated
                .AlwaysUseConnectionFile = False
            End With
            With owb.Connections("DataBase")
                .Name = "DataBase"
                .Description = ""
            End With
            owb.Connections("DataBase").Refresh()
            Select Case IO.Path.GetFileName(mye.FileName)
                Case "CP.xlsx"
                    owb.Worksheets(2).select()
                    Dim osheet = owb.Worksheets(2)
                    With osheet.PivotTables("PivotTable1").PivotFields("groupdivision")
                        .PivotItems("Finish Goods").Visible = False
                    End With
                    osheet.PivotTables("PivotTable1").PivotFields("groupcomponentcategory").ShowDetail = False

                    owb.Worksheets(3).select()
                    osheet = owb.Worksheets(3)
                    With osheet.PivotTables("PivotTable3").PivotFields("groupdivision")
                        .PivotItems("Finish Goods").Visible = False
                    End With
                    osheet.PivotTables("PivotTable3").PivotFields("groupcomponentcategory").ShowDetail = False

                    owb.Worksheets(4).select()
                    osheet = owb.Worksheets(4)
                    With osheet.PivotTables("PivotTable4").PivotFields("groupdivision")
                        .PivotItems("Finish Goods").Visible = False
                    End With
                    osheet.PivotTables("PivotTable4").PivotFields("vendorname").ShowDetail = False

                    owb.Worksheets(5).select()
                    osheet = owb.Worksheets(5)
                    With osheet.PivotTables("PivotTable5").PivotFields("groupdivision")
                        .PivotItems("Finish Goods").Visible = False
                    End With                 
                    osheet.PivotTables("PivotTable5").PivotFields("componentcategory").ShowDetail = False

                    owb.Worksheets(6).select()
                    osheet = owb.Worksheets(6)
                    With osheet.PivotTables("PivotTable1").PivotFields("groupdivision")
                        .PivotItems("Finish Goods").Visible = False
                    End With
                    osheet.PivotTables("PivotTable1").PivotFields("divisionname").ShowDetail = False
                Case "FG.xlsx"
                    owb.Worksheets(2).select()
                    Dim osheet = owb.Worksheets(2)
                    With osheet.PivotTables("PivotTable1").PivotFields("groupdivision")
                        .PivotItems("Component").Visible = False
                    End With
                    osheet.PivotTables("PivotTable1").PivotFields("groupcomponentcategory").ShowDetail = False

                    owb.Worksheets(3).select()
                    osheet = owb.Worksheets(3)
                    With osheet.PivotTables("PivotTable3").PivotFields("groupdivision")
                        .PivotItems("Component").Visible = False
                    End With
                    osheet.PivotTables("PivotTable3").PivotFields("groupcomponentcategory").ShowDetail = False

                    owb.Worksheets(4).select()
                    osheet = owb.Worksheets(4)
                    With osheet.PivotTables("PivotTable4").PivotFields("groupdivision")
                        .PivotItems("Component").Visible = False
                    End With
                    osheet.PivotTables("PivotTable4").PivotFields("vendorname").ShowDetail = False

                    owb.Worksheets(5).select()
                    osheet = owb.Worksheets(5)
                    With osheet.PivotTables("PivotTable5").PivotFields("groupdivision")
                        .PivotItems("Component").Visible = False
                    End With
                    osheet.PivotTables("PivotTable5").PivotFields("componentcategory").ShowDetail = False

                    owb.Worksheets(6).select()
                    osheet = owb.Worksheets(6)
                    With osheet.PivotTables("PivotTable1").PivotFields("groupdivision")
                        .PivotItems("Component").Visible = False
                    End With
                    osheet.PivotTables("PivotTable1").PivotFields("divisionname").ShowDetail = False
            End Select
        End If
        

    End Sub

    Private Sub PivotCallback(ByRef sender As Object, ByRef e As EventArgs)
        'MessageBox.Show("Call back")
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(1).select()
        Dim osheet = owb.Worksheets(1)
        Dim orange = osheet.Range("A2")

        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If

        osheet.name = "RawData"
        osheet.Columns("Z:Z").NumberFormat = "m/d/yyyy"
        osheet.Columns("AH:AH").NumberFormat = "m/d/yyyy"
        osheet.Columns("AL:AL").NumberFormat = "m/d/yyyy"

        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")

       
    End Sub

    Private Sub InitData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoQuery)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoQuery()
        Try
            ProgressReport(1, "Preparing Data...Please wait.")
            PeriodBS = New BindingSource
            PeriodBS = myController.myModel.GetPeriodAll
            ProgressReport(4, "Init Data")
            ProgressReport(1, String.Format("Loading...Done."))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    ComboBox1.DisplayMember = "txdatestring"
                    ComboBox1.ValueMember = "txdate"
                    ComboBox1.DataSource = PeriodBS
                    ComboBox1.SelectedIndex = 0
            End Select
        End If
    End Sub

    Private Sub FormGenerateReportExposure_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitData()
    End Sub
End Class