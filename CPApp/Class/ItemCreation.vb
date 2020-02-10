Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Public Class ItemCreation
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function

    Private Filename As String = String.Empty
    Private DRV As DataRowView
    Private Parent As Object
    Public Sub New(ByVal Parent As Object, ByVal FileName As String, ByVal DRV As DataRowView)
        Me.Filename = FileName
        Me.DRV = DRV
        Me.Parent = Parent
    End Sub

    Public Function GenerateExcel() As Boolean
        Dim myret As Boolean = False
        Dim oxl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim hwnd As System.IntPtr
        Try
            


            Parent.ProgressReport(1, "Creating Excel Automation...")
            oxl = CreateObject("Excel.Application")
            hwnd = oxl.Hwnd
            owb = oxl.Workbooks.Open(Application.StartupPath & "\Template\ItemCreation.xltx")

            oSheet = owb.Worksheets(1)
            oxl.Visible = False

            oSheet.Cells(11, 2).Value = Parent.NameOfRequester
            oSheet.Cells(12, 2).Value = IIf(IsNothing(Parent.DateOfRequest), "", Parent.DateOfRequest)
            oSheet.Cells(13, 2).Value = Parent.TypeOfRequest
            oSheet.Cells(16, 2).Value = DRV.Row.Item("cmmf") 'Trim(rsTmp!cmmf) 'CMMF
            oSheet.Cells(17, 2).Value = DRV.Row.Item("materialcode") 'Trim(rsTmp!materialcode) 'Item Code
            oSheet.Cells(18, 2).Value = DRV.Row.Item("Description") 'Trim(rsTmp!Description) 'Description
            oSheet.Cells(19, 2).Value = IIf(Parent.NewProject, "Yes", "No") 'IIf(Option1.Value, "Yes", "No")
            oSheet.Cells(20, 2).Value = IIf(Parent.PlatformProject, "Yes", "No") 'IIf(Option4.Value, "Yes", "No")
            oSheet.Cells(21, 2).Value = IIf(Parent.DelegatedItem, "Yes", "No") 'IIf(Option5.Value, "Yes", "No")
            oSheet.Cells(22, 2).Value = DRV.Row.Item("projectname") 'Trim(rsTmp!projectname)
            oSheet.Cells(23, 2).Value = DRV.Row.Item("brandname") 'Trim(rsTmp!brandname)
            oSheet.Cells(24, 2).Value = DRV.Row.Item("rangename") 'Trim(rsTmp!rangename)

            oSheet.Cells(25, 2).Value = String.Format("{0} {1}", Parent.Supplier, Parent.VendorCode) 'Trim(Combo2.Text) & " " & Combo2.ItemData(Combo2.ListIndex)   'Supplier
            oSheet.Cells(26, 2).Value = String.Format("{0} {1}", DRV.Row.Item("loadinggroup"), DRV.Row.Item("loadingname")) 'Trim(rsTmp!loadinggroup) & " " & Trim(rsTmp!loadingname) 'Loading Group
            oSheet.Cells(27, 2).Value = String.Format("{0} {1}", DRV.Row.Item("purchasinggroup"), DRV.Row.Item("typeofitem")) 'Trim(rsTmp!purchasinggroup) & " " & Trim(rsTmp!typeofitem)
            oSheet.Cells(29, 2).Value = DRV.Row.Item("netprice") 'Trim(rsTmp!netprice)
            oSheet.Cells(30, 2).Value = DRV.Row.Item("amort") 'Trim(rsTmp!amort)
            oSheet.Cells(31, 2).Value = DRV.Row.Item("contractno") 'Trim(rsTmp!contractno)
            oSheet.Cells(32, 2).Value = DRV.Row.Item("sppet") 'Trim(rsTmp!sppet)
            oSheet.Cells(33, 2).Value = DRV.Row.Item("spps") 'Trim(rsTmp!spps)

            oSheet.Cells(34, 2).Value = String.Format("{0} {1}", DRV.Row.Item("mrpcontrollercode"), DRV.Row.Item("spm")) 'Trim(rsTmp!mrpcontrollercode) & " " & Trim(rsTmp!spm)
            oSheet.Cells(35, 2).Value = DRV.Row.Item("leadtime") 'Trim(rsTmp!leadtime)
            oSheet.Cells(36, 2).Value = DRV.Row.Item("qty20") 'Trim(rsTmp!qty20)
            oSheet.Cells(37, 2).Value = DRV.Row.Item("qty40") 'Trim(rsTmp!qty40)
            oSheet.Cells(38, 2).Value = DRV.Row.Item("qty40hq") 'Trim(rsTmp!qty40hq)
            oSheet.Cells(39, 2).Value = DRV.Row.Item("moq") 'Trim(rsTmp!moq)
            oSheet.Cells(40, 2).Value = DRV.Row.Item("pcspercartoon") 'Trim(rsTmp!pcspercartoon)
            oSheet.Cells(41, 2).Value = DRV.Row.Item("Length") 'Trim(rsTmp!Length)
            oSheet.Cells(42, 2).Value = DRV.Row.Item("Width") 'Trim(rsTmp!Width)
            oSheet.Cells(43, 2).Value = DRV.Row.Item("Height") 'Trim(rsTmp!Height)
            oSheet.Cells(44, 2).Value = DRV.Row.Item("lengthbox") 'Trim(rsTmp!lengthbox)
            oSheet.Cells(45, 2).Value = DRV.Row.Item("widthbox") 'Trim(rsTmp!widthbox)
            oSheet.Cells(46, 2).Value = DRV.Row.Item("heightbox") 'Trim(rsTmp!heightbox)
            oSheet.Cells(47, 2).Value = String.Format("{0:#,##0.00}", DRV.Row.Item("weightwo")) 'Trim(rsTmp!weightwo)
            oSheet.Cells(48, 2).Value = String.Format("{0:#,##0.00}", DRV.Row.Item("weightwi")) 'Trim(rsTmp!weightwi)
            oSheet.Cells(49, 2).Value = String.Format("{0:#,##0.00}", DRV.Row.Item("grossweight")) 'Trim(rsTmp!grossweight)
            oSheet.Cells(50, 2).Value = Parent.VendorCode 'Trim(Combo2.ItemData(Combo2.ListIndex))

            owb.SaveAs(Filename)

        Catch ex As Exception
            Parent.ProgressReport(1, ex.Message)
        Finally
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
        End Try

        Return myret
    End Function
    Public Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub
End Class
