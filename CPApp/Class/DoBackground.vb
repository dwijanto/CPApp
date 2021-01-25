Imports System.Threading
'Public Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
Public Delegate Sub ProgressReportCallback(ByRef sender As Object, ByRef e As EventArgs)

'Callback can be called by Event or Delegate
'Parent must have :
'- Toolstripstatuslabel1,
'- Toolstripstatuslabel2,
'- ToolStripProgressBar1


Public Class DoBackground
    Private Parent As Object
    Private _progressreportCallback As ProgressReportCallback
    Private _myThread As Thread

    Public Shared Event CallBack(ByVal sender As Object, ByVal e As EventArgs)

    Public Property myThread As Thread
        Get
            Return _myThread
        End Get
        Set(value As Thread)
            _myThread = value
        End Set
    End Property

    Public Property ProgressReportCallback As ProgressReportCallback
        Get
            Return _progressreportCallback
        End Get
        Set(value As ProgressReportCallback)
            _progressreportCallback = value
        End Set
    End Property

    Public Sub New(Parent As Object, myCallBack As ProgressReportCallback)
        Me.Parent = Parent
        Me.ProgressReportCallback = myCallBack

    End Sub

    Public Sub New(Parent As Object)
        Me.Parent = Parent
    End Sub

    Public Sub doThread(start As System.Threading.ThreadStart)
        If IsNothing(myThread) Then
            run(start)
        Else
            If Not myThread.IsAlive Then
                run(start)
            Else
                MessageBox.Show("Please wait until the current process is finished.")
            End If
        End If
    End Sub

    Sub run(start As System.Threading.ThreadStart)
        myThread = New Thread(start)
        myThread.SetApartmentState(ApartmentState.STA)
        myThread.Start()
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Parent.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Parent.Toolstripstatuslabel1.text = message
                Case 2
                    Parent.Toolstripstatuslabel2.text = message
                Case 4
                    Try
                        ProgressReportCallback.Invoke(Me, New System.EventArgs)
                    Catch ex As Exception
                    End Try
                    RaiseEvent CallBack(4, New EventArgs)
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 8
                    RaiseEvent CallBack(8, New EventArgs)
            End Select
        End If
    End Sub



End Class
