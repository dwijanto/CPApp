<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ToolStripContainer1 = New System.Windows.Forms.ToolStripContainer()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RBACToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportBufferStockToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportForecastToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportExposureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BufferStockToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExposureToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExposureComparisonToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CustomerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.VendorToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ParameterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExposureRawDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripContainer1.BottomToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.TopToolStripPanel.SuspendLayout()
        Me.ToolStripContainer1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStripContainer1
        '
        '
        'ToolStripContainer1.BottomToolStripPanel
        '
        Me.ToolStripContainer1.BottomToolStripPanel.Controls.Add(Me.StatusStrip1)
        '
        'ToolStripContainer1.ContentPanel
        '
        Me.ToolStripContainer1.ContentPanel.Size = New System.Drawing.Size(594, 79)
        Me.ToolStripContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ToolStripContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStripContainer1.Name = "ToolStripContainer1"
        Me.ToolStripContainer1.Size = New System.Drawing.Size(594, 125)
        Me.ToolStripContainer1.TabIndex = 1
        Me.ToolStripContainer1.Text = "ToolStripContainer1"
        '
        'ToolStripContainer1.TopToolStripPanel
        '
        Me.ToolStripContainer1.TopToolStripPanel.Controls.Add(Me.MenuStrip1)
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 0)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(594, 22)
        Me.StatusStrip1.TabIndex = 0
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Dock = System.Windows.Forms.DockStyle.None
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AdminToolStripMenuItem, Me.ActionsToolStripMenuItem, Me.ReportToolStripMenuItem, Me.MasterToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(594, 24)
        Me.MenuStrip1.TabIndex = 0
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RBACToolStripMenuItem})
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.AdminToolStripMenuItem.Text = "Admin"
        '
        'RBACToolStripMenuItem
        '
        Me.RBACToolStripMenuItem.Name = "RBACToolStripMenuItem"
        Me.RBACToolStripMenuItem.Size = New System.Drawing.Size(104, 22)
        Me.RBACToolStripMenuItem.Tag = "FormRBAC"
        Me.RBACToolStripMenuItem.Text = "RBAC"
        '
        'ActionsToolStripMenuItem
        '
        Me.ActionsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportBufferStockToolStripMenuItem, Me.ImportForecastToolStripMenuItem, Me.ImportExposureToolStripMenuItem})
        Me.ActionsToolStripMenuItem.Name = "ActionsToolStripMenuItem"
        Me.ActionsToolStripMenuItem.Size = New System.Drawing.Size(59, 20)
        Me.ActionsToolStripMenuItem.Text = "Actions"
        '
        'ImportBufferStockToolStripMenuItem
        '
        Me.ImportBufferStockToolStripMenuItem.Name = "ImportBufferStockToolStripMenuItem"
        Me.ImportBufferStockToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ImportBufferStockToolStripMenuItem.Tag = "FormImportBufferStock"
        Me.ImportBufferStockToolStripMenuItem.Text = "Import Buffer Stock"
        '
        'ImportForecastToolStripMenuItem
        '
        Me.ImportForecastToolStripMenuItem.Name = "ImportForecastToolStripMenuItem"
        Me.ImportForecastToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ImportForecastToolStripMenuItem.Tag = "FormImportDbDemand"
        Me.ImportForecastToolStripMenuItem.Text = "Import DB Demand"
        '
        'ImportExposureToolStripMenuItem
        '
        Me.ImportExposureToolStripMenuItem.Name = "ImportExposureToolStripMenuItem"
        Me.ImportExposureToolStripMenuItem.Size = New System.Drawing.Size(177, 22)
        Me.ImportExposureToolStripMenuItem.Tag = "FormImportExposure"
        Me.ImportExposureToolStripMenuItem.Text = "Import Exposure"
        '
        'ReportToolStripMenuItem
        '
        Me.ReportToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BufferStockToolStripMenuItem, Me.ExposureToolStripMenuItem, Me.ExposureComparisonToolStripMenuItem, Me.ExposureRawDataToolStripMenuItem})
        Me.ReportToolStripMenuItem.Name = "ReportToolStripMenuItem"
        Me.ReportToolStripMenuItem.Size = New System.Drawing.Size(54, 20)
        Me.ReportToolStripMenuItem.Text = "Report"
        '
        'BufferStockToolStripMenuItem
        '
        Me.BufferStockToolStripMenuItem.Name = "BufferStockToolStripMenuItem"
        Me.BufferStockToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.BufferStockToolStripMenuItem.Tag = "FormGenerateReportBufferStock"
        Me.BufferStockToolStripMenuItem.Text = "Buffer Stock"
        '
        'ExposureToolStripMenuItem
        '
        Me.ExposureToolStripMenuItem.Name = "ExposureToolStripMenuItem"
        Me.ExposureToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.ExposureToolStripMenuItem.Tag = "FormGenerateReportExposure"
        Me.ExposureToolStripMenuItem.Text = "Exposure"
        '
        'ExposureComparisonToolStripMenuItem
        '
        Me.ExposureComparisonToolStripMenuItem.Name = "ExposureComparisonToolStripMenuItem"
        Me.ExposureComparisonToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.ExposureComparisonToolStripMenuItem.Tag = "FormGenerateReportExposureComparison"
        Me.ExposureComparisonToolStripMenuItem.Text = "Exposure Comparison"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserToolStripMenuItem, Me.CustomerToolStripMenuItem, Me.VendorToolStripMenuItem, Me.ParameterToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.UserToolStripMenuItem.Tag = "FormUser"
        Me.UserToolStripMenuItem.Text = "User"
        '
        'CustomerToolStripMenuItem
        '
        Me.CustomerToolStripMenuItem.Name = "CustomerToolStripMenuItem"
        Me.CustomerToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.CustomerToolStripMenuItem.Text = "Customer"
        '
        'VendorToolStripMenuItem
        '
        Me.VendorToolStripMenuItem.Name = "VendorToolStripMenuItem"
        Me.VendorToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.VendorToolStripMenuItem.Text = "Vendor"
        '
        'ParameterToolStripMenuItem
        '
        Me.ParameterToolStripMenuItem.Name = "ParameterToolStripMenuItem"
        Me.ParameterToolStripMenuItem.Size = New System.Drawing.Size(128, 22)
        Me.ParameterToolStripMenuItem.Tag = "FormParameters"
        Me.ParameterToolStripMenuItem.Text = "Parameter"
        '
        'ExposureRawDataToolStripMenuItem
        '
        Me.ExposureRawDataToolStripMenuItem.Name = "ExposureRawDataToolStripMenuItem"
        Me.ExposureRawDataToolStripMenuItem.Size = New System.Drawing.Size(190, 22)
        Me.ExposureRawDataToolStripMenuItem.Tag = "FormGenerateReportExposureRawData"
        Me.ExposureRawDataToolStripMenuItem.Text = "Exposure RawData"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(594, 125)
        Me.Controls.Add(Me.ToolStripContainer1)
        Me.Name = "FormMenu"
        Me.Text = "FormMenu"
        Me.ToolStripContainer1.BottomToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.BottomToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.TopToolStripPanel.ResumeLayout(False)
        Me.ToolStripContainer1.TopToolStripPanel.PerformLayout()
        Me.ToolStripContainer1.ResumeLayout(False)
        Me.ToolStripContainer1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ToolStripContainer1 As System.Windows.Forms.ToolStripContainer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents ActionsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RBACToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportBufferStockToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CustomerToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents VendorToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BufferStockToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportForecastToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportExposureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExposureToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ParameterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExposureComparisonToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExposureRawDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
