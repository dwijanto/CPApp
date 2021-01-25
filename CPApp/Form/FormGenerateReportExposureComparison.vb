Imports Microsoft.Office.Interop
Imports System.Threading

Public Class FormGenerateReportExposureComparison
    Dim myThread As New System.Threading.Thread(AddressOf DoQuery)
    Public Shared myForm As FormGenerateReportExposureComparison
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New ReportExposureController
    'Dim myModel As New ExposureModel
    'Private myIdentity As UserController = User.getIdentity
    Private ViewAllData As Boolean = False
    Dim stPeriodBS As BindingSource
    Dim ndPeriodBS As BindingSource

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
            myForm = New FormGenerateReportExposureComparison
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportExposureComparison
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox1.SelectedIndex >= 0 Then
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            SaveFileDialog1.FileName = String.Format("ExposureComparison-{0:yyyyMMdd}.xlsx", Date.Today)
            Dim stdrv As DataRowView = ComboBox1.SelectedItem
            Dim nddrv As DataRowView = ComboBox2.SelectedItem
            Dim Criteria As String = String.Format("and (btx.txdate in ('{0:yyyy-MM-dd}','{1:yyy-MM-dd}') ) ", stdrv.Item("txdate"), nddrv("txdate"))

            If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
                'Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myModel.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\ExcelTemplate.xltx")
                Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.myModel.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 6, "\templates\ExposureTemplate.xltx")
                myReport.Run(Me, New EventArgs)
            End If
        Else
            MessageBox.Show("Please select from period.")
        End If

    End Sub

    Private Sub FormatReport()
        'Throw New NotImplementedException
    End Sub

    Private Sub PivotCallback(ByRef sender As Object, ByRef e As EventArgs)
        'MessageBox.Show("Call back")
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(6).select()
        Dim osheet = owb.Worksheets(6)
        Dim orange = osheet.Range("A2")

        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If

        osheet.name = "RawData"
        osheet.Columns("Z:Z").NumberFormat = "m/d/yyyy"
        osheet.Columns("AH:AH").NumberFormat = "m/d/yyyy"
        osheet.Columns("AL:AL").NumberFormat = "m/d/yyyy"

        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")

        owb.Worksheets(1).select()
        osheet = owb.Worksheets(1)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        '----osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        osheet.pivottables("PivotTable1").PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")


        owb.Worksheets(2).select()
        osheet = owb.Worksheets(2)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        'Threading.Thread.Sleep(100)

        'MessageBox.Show("refresh1")
        '----osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        osheet.pivottables("PivotTable1").PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(3).select()
        osheet = owb.Worksheets(3)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        '----osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        osheet.pivottables("PivotTable1").PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(4).select()
        osheet = owb.Worksheets(4)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        '----osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        osheet.pivottables("PivotTable1").PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(5).select()
        osheet = owb.Worksheets(5)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        '---osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        osheet.pivottables("PivotTable1").PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        'owb.RefreshAll()
        'osheet.Cells.EntireColumn.AutoFit()

        'Threading.Thread.Sleep(5000)
        owb.Worksheets(1).select()
        'Threading.Thread.Sleep(1000)
        osheet = owb.Worksheets(1)
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
            stPeriodBS = New BindingSource
            stPeriodBS = myController.myModel.GetPeriod
            ndPeriodBS = New BindingSource
            ndPeriodBS = myController.myModel.GetPeriod
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
                    ComboBox1.DataSource = stPeriodBS
                    ComboBox1.SelectedIndex = 0

                    ComboBox2.DisplayMember = "txdatestring"
                    ComboBox2.ValueMember = "txdate"
                    ComboBox2.DataSource = ndPeriodBS
                    ComboBox2.SelectedIndex = 0
            End Select
        End If
    End Sub

    Private Sub FormGenerateReportExposure_Load(sender As Object, e As EventArgs) Handles Me.Load
        InitData()
    End Sub
End Class