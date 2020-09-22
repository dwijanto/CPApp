Imports System.Threading
Imports System.IO

Public Class FormImportExposure

    Public Shared myform As FormImportExposure
    Dim myThread As New Thread(AddressOf DoImportFile)
    Public Property Filename As String
    Public DS As DataSet
    ' Dim BStockController As BufferStockController = New BufferStockController


    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormImportExposure
        ElseIf myform.IsDisposed Then
            myform = New FormImportExposure
        End If
        Return myform
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
            Dim sw As New Stopwatch
            sw.Start()
            ProgressReport(6, "")
            Dim myImport = New ImportExposure(Me)
            If Not myImport.run() Then
                sw.Stop()
                If myImport.ErrorMessage.Length > 100 Then
                    Dim myfilenameerror As String = String.Format("{0}\{1}Error.txt", Path.GetDirectoryName(Me.Filename), Path.GetFileNameWithoutExtension(Me.Filename))
                    Using mystream As New StreamWriter(myfilenameerror)
                        mystream.WriteLine(myImport.ErrorMessage)
                    End Using
                    Process.Start(myfilenameerror)
                    ProgressReport(1, String.Format("Error Found. "))
                Else
                    ProgressReport(1, String.Format("Error: {0}", myImport.ErrorMessage))
                End If

                Exit Sub
            End If
            sw.Stop()
            ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
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