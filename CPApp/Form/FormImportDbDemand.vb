Imports System.Threading
Public Class FormImportDbDemand
    Public Shared myform As FormImportDbDemand
    Dim myThread As New Thread(AddressOf DoImportFile)
    Public Property Filename As String
    Public DS As DataSet
    Dim MyController As DBDemandController = New DBDemandController
    Dim FGSelected As Boolean


    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormImportDbDemand
        ElseIf myform.IsDisposed Then
            myform = New FormImportDbDemand
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
        Dim sw As New Stopwatch
        sw.Start()
        Try


            ProgressReport(6, "")
            'DS = MyController.GetDSForImport
            Dim myimport As Object
            If FGSelected Then
                myimport = New ImportDbDemand(Me)
            Else
                myimport = New ImportCPDbDemand(Me)
            End If

            If Not myImport.run() Then
                ProgressReport(1, String.Format("Error: {0}", myImport.errMsgSB.ToString))
            Else
                sw.Stop()
                ProgressReport(1, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            End If
        Catch ex As Exception
            ProgressReport(5, "")
            MessageBox.Show(ex.Message)
        Finally
            sw.Stop()
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

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        FGSelected = RadioButton1.Checked
    End Sub
End Class