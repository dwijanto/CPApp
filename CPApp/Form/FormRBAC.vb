Imports System.ComponentModel
Public Class FormRBAC
    Implements INotifyPropertyChanged

    Private Shared myform As FormRBAC
    Public Shared Function getInstance()
        If myform Is Nothing Then
            myform = New FormRBAC
        ElseIf myform.IsDisposed Then
            myform = New FormRBAC
        End If
        Return myform
    End Function

    Public Event PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Implements INotifyPropertyChanged.PropertyChanged

    Public Property ItemType As Integer
        Get
            If RadioButton1.Checked = True Then
                Return Int(TypeEnum.TYPE_ROLE)
            ElseIf RadioButton2.Checked = True Then
                Return Int(TypeEnum.TYPE_PERMISSION)
            Else
                Return 0
            End If
        End Get
        Set(value As Integer)
            If value = Int(TypeEnum.TYPE_ROLE) Then
                RadioButton1.Checked = True
            ElseIf value = Int(TypeEnum.TYPE_PERMISSION) Then
                RadioButton2.Checked = True
            End If
        End Set
    End Property


    Dim myRBAC As DbManager = New DbManager



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim myItem As Item
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MessageBox.Show("Textbox cannot be empty.")
            Exit Sub
        End If
        If ItemType = Int(TypeEnum.TYPE_ROLE) Then
            myItem = New Role
        ElseIf ItemType = Int(TypeEnum.TYPE_PERMISSION) Then
            myItem = New Permission
        Else
            MessageBox.Show("Please select Type.")
            Exit Sub
        End If
        myItem.name = TextBox1.Text
        myItem.description = TextBox2.Text
        myItem.type = ItemType
        myItem.createdAt = Date.Now
        myItem.updatedAt = Date.Now
        If myRBAC.addItem(myItem) Then
            MessageBox.Show(String.Format("[{0}] - [{1}] has been added", TextBox1.Text, TextBox2.Text))
            TextBox1.Clear()
            TextBox2.Clear()
        End If

    End Sub

   

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        onPropertyChanged("ItemType")
    End Sub

    Private Sub onPropertyChanged(Name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(Name))
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            'Dim parent As Item = myRBAC.getItem(TextBox3.Text)
            'Dim child As Item = myRBAC.getItem(TextBox4.Text)
            Dim parent As Item = ComboBox3.SelectedItem
            Dim child As Item = ComboBox4.SelectedItem
            If Not IsNothing(parent) Or IsNothing(child) Then
                Try
                    myRBAC.addChild(parent, child)
                    MessageBox.Show(String.Format("[{0}] - [{1}] has been added", parent.name, child.name))
                    'TextBox3.Clear()
                    'TextBox4.Clear()
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim myitem As Item = ComboBox1.SelectedItem
            Dim myRole As Role = myitem
            Dim myUser As DataRowView = ComboBox2.SelectedItem
            myRBAC.revoke(myRole, myUser.Item("id"))
            MessageBox.Show(String.Format("[{0}] - [{1}] has been revoked", myitem.name, myUser.Item("username")))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim myitem As Item = ComboBox1.SelectedItem
            Dim myRole As Role = myitem
            Dim myUser As DataRowView = ComboBox2.SelectedItem
            myRBAC.assign(myRole, myUser.Item("id"))
            MessageBox.Show(String.Format("[{0}] - [{1}] has been added", myitem.name, myUser.Item("username")))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub FormRBAC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadCombobox()
    End Sub

    Private Sub loadCombobox()
        Dim myrole As List(Of Item) = myRBAC.getItems(TypeEnum.TYPE_ROLE)
        If myrole.Count = 0 Then
            Exit Sub
        End If
        ComboBox1.Items.Clear()
        ComboBox1.DisplayMember = "name"
        ComboBox1.ValueMember = "id"

        For Each Role As Item In myrole
            ComboBox1.Items.Add(Role)
        Next
        ComboBox1.SelectedIndex = 0

        Dim UC1 As New UserController
        Dim myuser As DataSet = UC1.findAll(other:="order by username")
        Dim bs As New BindingSource
        bs.DataSource = myuser.Tables(0)

        ComboBox2.DisplayMember = "username"
        ComboBox2.ValueMember = "id"
        ComboBox2.DataSource = bs
        ComboBox2.SelectedIndex = 0

        Dim AllItems As List(Of Item) = myRBAC.getAllItems
        ComboBox3.Items.Clear()
        ComboBox3.DisplayMember = "name"
        ComboBox3.ValueMember = "id"

        For Each item As Item In AllItems
            ComboBox3.Items.Add(item)
        Next
        ComboBox3.SelectedIndex = 0

        ComboBox4.Items.Clear()
        ComboBox4.DisplayMember = "name"
        ComboBox4.ValueMember = "id"

        For Each item As Item In AllItems
            ComboBox4.Items.Add(item)
        Next
        ComboBox4.SelectedIndex = 0
    End Sub

 
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        loadCombobox()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim mydrv As DataRowView = ComboBox2.SelectedItem
        Dim assignments As List(Of Assignment) = myRBAC.getAssignments(mydrv.Row.Item("id"))
        ListBox1.Items.Clear()
        For Each a In assignments
            ListBox1.Items.Add(a.rolename)
        Next
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Dim myrole As List(Of Item) = myRBAC.getItems(TypeEnum.TYPE_ROLE)
        ListBox2.Items.Clear()
        For Each a In myrole
            ListBox2.Items.Add(a.name)
        Next

        Dim mypermission As List(Of Item) = myRBAC.getItems(TypeEnum.TYPE_PERMISSION)
        ListBox3.Items.Clear()
        For Each a In mypermission
            ListBox3.Items.Add(a.name)
        Next

        Dim DS As New DataSet
        'myAdapter.GetDataset("select * from quality.auth_item_child;", ds)
        DS = DataAccess.GetDataSet("select * from cp.auth_item_child;", CommandType.Text, Nothing)
        Dim bs As New BindingSource
        bs.DataSource = ds.Tables(0)

        DataGridView1.DataSource = bs


    End Sub


    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            'Dim parent As Item = myRBAC.getItem(TextBox3.Text)
            'Dim child As Item = myRBAC.getItem(TextBox4.Text)
            Dim parent As Item = ComboBox3.SelectedItem
            Dim child As Item = ComboBox4.SelectedItem
            If Not IsNothing(parent) Or IsNothing(child) Then
                Try
                    myRBAC.removeChild(parent, child)
                    MessageBox.Show(String.Format("[{0}] - [{1}] has been removed", parent.name, child.name))
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class