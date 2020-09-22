Imports Microsoft.Office.Interop
Public Class FormGenerateReportExposure

    Public Shared myForm As FormGenerateReportExposure
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myModel As New ExposureModel
    'Private myIdentity As UserController = User.getIdentity
    Private ViewAllData As Boolean = False
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
            myForm = New FormGenerateReportExposure
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportExposure
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ToolStripStatusLabel1.Text = ""
        ToolStripStatusLabel2.Text = ""
        SaveFileDialog1.FileName = String.Format("Exposure-{0:yyyyMMdd}.xlsx", Date.Today)
        Dim Criteria As String = String.Empty

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            'Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myModel.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\ExcelTemplate.xltx")
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myModel.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 6, "\templates\ExposureTemplate.xltx")
            myReport.Run(Me, New EventArgs)
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
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")


        owb.Worksheets(2).select()
        osheet = owb.Worksheets(2)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        Threading.Thread.Sleep(100)

        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(3).select()
        osheet = owb.Worksheets(3)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        osheet.PivotTables("PivotTable1").PivotFields("monthly").AutoSort(Excel.XlSortOrder.xlAscending, "monthly")
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(4).select()
        osheet = owb.Worksheets(4)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        owb.Worksheets(5).select()
        osheet = owb.Worksheets(5)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")

        'owb.RefreshAll()
        'osheet.Cells.EntireColumn.AutoFit()
        Threading.Thread.Sleep(5000)
        owb.Worksheets(1).select()
        Threading.Thread.Sleep(1000)
        osheet = owb.Worksheets(1)
    End Sub
End Class