Imports System.Threading

Public Class FormParameters

    Private Shared myform As FormParameters

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormParameters
        ElseIf myform.IsDisposed Then
            myform = New FormParameters
        End If
        Return myform
    End Function

    Dim myController As ParamAdapter
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        myController = New ParamAdapter
        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, String.Format("Loading...Done."))
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try

    End Sub

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
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
                    Dim parentid As Long
                    parentid = myController.GetParentid("Component Category")
                    UcdgvParam1.DataGridView1.Columns(0).HeaderText = "Name"
                    UcdgvParam1.DataGridView1.Columns(0).DataPropertyName = "cvalue"
                    UcdgvParam1.DataGridView1.Columns(0).ReadOnly = False                    
                    UcdgvParam1.BindingControl(Me, myController.BS, parentid, "")

                    parentid = myController.GetParentid("Main Reason")
                    UcdgvParam2.DataGridView1.Columns(0).HeaderText = "Name"
                    UcdgvParam2.DataGridView1.Columns(0).DataPropertyName = "cvalue"
                    UcdgvParam2.DataGridView1.Columns(0).ReadOnly = False                    
                    UcdgvParam2.BindingControl(Me, myController.BS2, parentid, "")

                    parentid = myController.GetParentid("Division")
                    UcdgvParam3.DataGridView1.Columns(0).HeaderText = "Name"
                    UcdgvParam3.DataGridView1.Columns(0).DataPropertyName = "cvalue"
                    UcdgvParam3.DataGridView1.Columns(0).ReadOnly = False                   
                    UcdgvParam3.BindingControl(Me, myController.BS3, parentid, "")
            End Select
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Me.Validate()
        myController.save()
    End Sub


    Private Sub FormParameters_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()

    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        LoadData()
    End Sub
End Class