Imports System.Threading
Public Class FormImportBufferStock
    Public Shared myform As FormImportBufferStock
    Dim myThread As New Thread(AddressOf DoImportFile)
    Public Property Filename As String
    Public DS As DataSet
    Dim BStockController As BufferStockController = New BufferStockController


    Public Shared Function getInstance()
        If myForm Is Nothing Then
            myform = New FormImportBufferStock
        ElseIf myForm.IsDisposed Then
            myform = New FormImportBufferStock
        End If
        Return myForm
    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Me.Filename = OpenFileDialog1.FileName
            DoImport()
        End If

    End Sub

    Private Sub DoImport()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoImportFile)
            myThread.Start()
        Else
            ProgressReport(1, "Please wait until the current process finished.")
        End If
    End Sub

    Sub DoImportFile()
        'Get Existing Data
        Try
            ProgressReport(6, "")
            DS = BStockController.GetDSForImport
            Dim myImport = New ImportBufferStock(Me)
            If Not myImport.run() Then
                ProgressReport(1, String.Format("Error: {0}", myImport.ErrorMessage))
            End If
        Catch ex As Exception
            ProgressReport(5, "")
            MessageBox.Show(ex.Message)
        Finally
            ProgressReport(5, "")
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

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

            End Select
        End If
    End Sub
End Class