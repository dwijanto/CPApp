Public Class UCDGVParam
    Private myform As Object
    Private WithEvents BS As BindingSource
    Private ParentId As Long
    Private ParamName As String
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
    End Sub
    Public Sub BindingControl(ByVal myform As Object, ByRef bs As BindingSource, ByVal parentid As Long, ByVal paramname As String)
        Me.ParentId = parentid
        Me.myform = myform
        Me.BS = bs
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
        Me.ParamName = paramname
    End Sub
    Public Sub BindingControl(ByVal myform As Object, ByRef bs As BindingSource, ByVal parentid As Long)
        Me.ParentId = parentid
        Me.myform = myform
        Me.BS = bs
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
    End Sub
    Private Sub NewRecordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewRecordToolStripMenuItem.Click
        Dim drv = BS.AddNew
        drv.item("paramhdid") = ParentId
        drv.item("paramname") = Me.ParamName

    End Sub

    Private Sub DeleteRecordToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteRecordToolStripMenuItem.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                    BS.RemoveAt(drv.Index)
                Next
            End If
        End If
    End Sub
End Class
