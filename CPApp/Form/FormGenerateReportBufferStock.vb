Imports Microsoft.Office.Interop
Public Class FormGenerateReportBufferStock

    Public Shared myForm As FormGenerateReportBufferStock
    Private SaveFileName As String = String.Empty
    Private SaveFileDialog1 As New SaveFileDialog
    Dim myController As New BufferStockController
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
            myForm = New FormGenerateReportBufferStock
        ElseIf myForm.IsDisposed Then
            myForm = New FormGenerateReportBufferStock
        End If
        Return myForm
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        myController = New BufferStockController
        SaveFileDialog1.FileName = String.Format("BufferStock-{0:yyyyMMdd}.xlsx", Date.Today)
        Dim Criteria As String = String.Empty
        
        If SaveFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim myReport As ExportToExcelFile = New ExportToExcelFile(Me, myController.GetSQLSTRReport(Criteria), IO.Path.GetDirectoryName(SaveFileDialog1.FileName), IO.Path.GetFileName(SaveFileDialog1.FileName), AddressOf FormatReport, AddressOf PivotCallback, 2, "\templates\BufferStock.xltx")
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
        'Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")
        owb.RefreshAll()
        'osheet.Cells.EntireColumn.AutoFit()
        Threading.Thread.Sleep(100)
    End Sub
End Class