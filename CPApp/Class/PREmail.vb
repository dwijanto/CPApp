Imports System.Text
Imports System.Net.Mail
Imports System.Net.Mime

Public Class PREmail
    Inherits Email

    Private sendtoname As String
    Private drv As DataRowView
    Public _errorMessage As String = String.Empty
    Private statusname As String = String.Empty
    Private dtlbs As BindingSource
    Private Function GetBodyMessage() As String
        Dim sb As New StringBuilder

        sb.Append("<!DOCTYPE html><html><head><meta name='description' content='[ProductRequest]' /><meta http-equiv='Content-Type' content='text/html; charset=us-ascii'></head>" &
                  "<style>td,th {padding-left:5px;padding-right:10px;  }  th {background-color:red;    color:white;text-align:left;}  .defaultfont{    font-size:11.0pt; font-family:'Calibri','sans-serif';    }</style><body class='defaultfont'>")
        sb.Append(String.Format("<p>Dear {0},</p> <p>Please be informed that we have the following product request need your approval/follow up: <table border = 0 cellspacing = 0 >", sendtoname))
        sb.Append(String.Format("<tr><td>Request Number<td><td>:</td><td>{0}</td></tr> " &
                                "<tr><td>Status<td><td>:</td><td>{1}</td></tr><tr>" &
                                "<td>Delivery Details Name<td><td>:</td><td>{2}</td></tr>" &
                                "<tr><td>Applicant Date<td><td>:</td><td>{3:dd-MMM-yyyy}</td></tr>" &
                                "<tr><td>Applicant Name<td><td>:</td><td>{4}</td></tr>" &
                                "<tr><td>Approval From Dept<td><td>:</td><td>{5}</td></tr></table><br>", drv.Item("refnumber"), statusname, drv.Item("bpartnerfullname"), drv.Item("applicantdate"), drv.Item("applicantname"), drv.Item("deptapproval")))

        sb.Append("<table border=1 cellspacing=0 class='defaultfont'><tr><th>CMMF</th> <th>Product Name</th> <th>Quantity</th><th>Price</th><th>Total Cost HKD</th><th>Cost Element / Cost Center Record</th></tr>")
        Dim mytotal As Decimal = 0
        For Each mydrv In dtlbs.List
            Dim total As Decimal
            Dim myqty As Integer
            If IsDBNull(mydrv.item("confirmedqty")) Then
                total = mydrv.item("qty") * mydrv.item("price")
                myqty = mydrv.item("qty")
            Else
                total = mydrv.item("confirmedqty") * mydrv.item("price")
                myqty = mydrv.item("confirmedqty")
            End If

            mytotal += total
            sb.Append(String.Format("<tr><td>{0}</td><td>{1}</td><td align=right>{2}</td><td align=right>{3:#,##0.00}</td><td align=right>{4:#,##0.00}</td><td>{5}</td></tr>",
                                mydrv.item("cmmf"), mydrv.item("localdescription"), myqty, mydrv.item("price"), total, mydrv.item("expensesname")))
        Next

        sb.Append(String.Format("</table><br>Total: {0:#,##0.00} <p>Thank you.<br><br>You can access the system in RD Web Access by below link:<br>   <a href=""https://sw07e601/RDWeb"">Product Request Application.</a></p><br></body></html>", mytotal))

        Return sb.ToString
    End Function

    Public ReadOnly Property ErrorMessage As String
        Get
            Return _errorMessage
        End Get
    End Property


    Public Function Execute(ByVal sendto As String, ByVal sendtoname As String, ByVal statusname As String, ByVal drv As DataRowView, ByVal DTLBS As BindingSource, Optional ByVal cc As String = "") As Boolean
        Dim myret As Boolean = False
        Try
            'Prepare Email
            Me.statusname = statusname
            Me.sendtoname = sendtoname
            Me.drv = drv
            Me.sendto = Trim(sendto)
            Me.subject = String.Format("Product Request Application: Tasks status. ({0:dd-MMM-yyyy}).", Today.Date)
            Me.dtlbs = DTLBS

            If Not IsNothing(Me.sendto) Then

                Dim mycontent = GetBodyMessage()

                Dim htmlView As AlternateView = AlternateView.CreateAlternateViewFromString(String.Format("{0} <br>Or click the Product Request Application icon on your desktop: <br><p> <img src=cid:myLogo> <br></p><p>Produt Request System Administrator</p></body></html>", mycontent), Nothing, MediaTypeNames.Text.Html)

                Dim logo As New LinkedResource(Application.StartupPath & "\PR.png")
                logo.ContentId = "myLogo"
                htmlView.LinkedResources.Add(logo)

                Me.htmlView = htmlView
                Me.isBodyHtml = True
                Me.sender = "no-reply@groupeseb.com"
                Me.body = mycontent
                Me.cc = String.Format("{0}", cc)
                myret = Me.send(ErrorMessage)
            End If
        Catch ex As Exception
            Logger.log(ex.Message)
            MessageBox.Show(ex.Message)
        End Try

        Return myret
    End Function
End Class
