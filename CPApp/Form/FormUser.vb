Imports System.Threading

Public Class FormUser
    Private Shared myform As FormUser
    Dim myController As UserController

    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormUser
        ElseIf myform.IsDisposed Then
            myform = New FormUser
        End If
        Return myform
    End Function

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim drv As DataRowView

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub RefreshMYInterface()
        DataGridView1.Invalidate()
    End Sub

    Sub DoWork()
        myController = New UserController
        Try
            ProgressReport(1, "Loading...Please wait.")
            If myController.loaddata() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, String.Format("Loading...Done. Records {0}", myController.BS.Count))
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
                    DataGridView1.AutoGenerateColumns = False
                    DataGridView1.DataSource = myController.BS
            End Select
        End If
    End Sub


    Private Sub AddNewToolStripButton1_Click(sender As Object, e As EventArgs) Handles AddToolStripButton.Click
        showTX(TxEnum.NewRecord)
    End Sub

    Private Sub DeleteToolStripButton2_Click(sender As Object, e As EventArgs) Handles DeleteToolStripButton.Click
        If Not IsNothing(myController.GetCurrentRecord) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    myController.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub

    Private Sub FormITAssets_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
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

    Private Sub RefreshToolStripButton4_Click(sender As Object, e As EventArgs) Handles RefreshToolStripButton.Click
        LoadData()
    End Sub



    Private Sub CommitToolStripButton3_Click(sender As Object, e As EventArgs) Handles CommitToolStripButton.Click
        Me.Validate()
        myController.save()
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        Try
            myController.ApplyFilter = ToolStripTextBox1.Text
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub UpdateToolStripButton_Click(sender As Object, e As EventArgs) Handles UpdateToolStripButton.Click
        showTX(TxEnum.UpdateRecord)
    End Sub

    Private Sub showTX(txEnum As TxEnum)
        Dim ApprovalBS As New BindingSource
        Dim DeptBS As New BindingSource
       
        Dim drv As DataRowView = Nothing
        Select Case txEnum
            Case CPApp.TxEnum.NewRecord
                drv = myController.GetNewRecord
            Case CPApp.TxEnum.UpdateRecord
                drv = myController.GetCurrentRecord
        End Select
        drv.BeginEdit()

        Dim myform = New DialogAddUpdUser(drv)
        myform.ShowDialog()

    End Sub

    Private Sub DataGridView1_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        UpdateToolStripButton.PerformClick()
    End Sub

    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs) Handles ToolStripTextBox1.Click

    End Sub
End Class