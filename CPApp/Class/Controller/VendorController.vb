Public Class VendorController
    Private Model As New VendorModel
    Public BS As BindingSource
    Private DS As DataSet

    Public Function getVendors() As List(Of VendorModel)
        Return Model.GetVendors
    End Function
    Public Function getVendorsCustom() As List(Of VendorModel)
        Return Model.GetVendorsCustom
    End Function
    Public Function GetDataset() As DataSet
        DS = Model.GetDataSet
        Return DS
    End Function

    Public Function save() As Boolean
        Dim myret As Boolean = False
        BS.EndEdit()
        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

    Public Function Save(mye As ContentBaseEventArgs) As Boolean
        Dim myret As Boolean = False
        If Model.save(Me, mye) Then
            myret = True
        End If
        Return myret
    End Function



End Class
