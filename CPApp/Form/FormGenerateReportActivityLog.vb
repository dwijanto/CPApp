Imports Microsoft.Office.Interop

Public Class FormGenerateReportActivityLog

    Public Shared myForm As FormGenerateReportActivityLog
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New ActivityLogController
    Private myIdentity As UserController = User.getIdentity
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
            myForm = New FormGenerateReportActivityLog
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportActivityLog
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myController = New ActivityLogController
        SaveFileDialog1.FileName = String.Format("ActivityLog-{0:yyyyMMdd}.xlsx", Date.Today)
        Dim Criteria As String = String.Empty
        If User.can("View Activity Log All Data") And ViewAllData Then
            SaveFileDialog1.FileName = String.Format("ActivityLogAll-{0:yyyyMMdd}.xlsx", Date.Today)
            Criteria = String.Format(" where u.activitydate >= '{0:yyyy-MM-dd}' and u.activitydate <= '{1:yyyy-MM-dd}'", DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)
        Else
            Criteria = String.Format(" where u.userid in (select quality.getsubordinate('{0}')) and u.activitydate >= '{1:yyyy-MM-dd}' and u.activitydate <= '{2:yyyy-MM-dd}'", myIdentity.userid.ToLower, DateTimePicker1.Value.Date, DateTimePicker2.Value.Date)
        End If

        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.GetSQLSTRReport(criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\activitylog.xltx")
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
        owb.Worksheets(2).select()
        Dim osheet = owb.Worksheets(2)
        Dim orange = osheet.Range("A2")

        If osheet.cells(2, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If
        osheet.Columns("B:B").NumberFormat = "dd-MMM-yyyy"
        osheet.name = "RawData"


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
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")
        owb.RefreshAll()
        Threading.Thread.Sleep(100)
    End Sub

End Class