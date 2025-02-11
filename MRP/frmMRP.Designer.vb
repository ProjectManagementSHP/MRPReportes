<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMRP
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMRP))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.TabControlMRPGlobal = New System.Windows.Forms.TabControl()
        Me.TabPageMRP = New System.Windows.Forms.TabPage()
        Me.GroupBoxBudgetInformation = New System.Windows.Forms.GroupBox()
        Me.lblTotalTotal2 = New System.Windows.Forms.Label()
        Me.lblRecordsPerVendor = New System.Windows.Forms.Label()
        Me.lblRecordsPerWeek = New System.Windows.Forms.Label()
        Me.lblTPerVendor = New System.Windows.Forms.Label()
        Me.lblTPerWeek = New System.Windows.Forms.Label()
        Me.btnCloseBudgetInformation = New System.Windows.Forms.Button()
        Me.GridPerVendor = New System.Windows.Forms.DataGridView()
        Me.GridPerWeek = New System.Windows.Forms.DataGridView()
        Me.GroupBoxPurchasingOrderHistory = New System.Windows.Forms.GroupBox()
        Me.btnRefreshPurchasingOrderItemsHistory = New System.Windows.Forms.Button()
        Me.btnClosePurchasingOrderItemsHistory = New System.Windows.Forms.Button()
        Me.lblRecordsPurchasingOrderItemsHistory = New System.Windows.Forms.Label()
        Me.lblTItems = New System.Windows.Forms.Label()
        Me.GridPurchasingOrderItemsHistory = New System.Windows.Forms.DataGridView()
        Me.lblTotal = New System.Windows.Forms.LinkLabel()
        Me.txbExchangeRate = New System.Windows.Forms.TextBox()
        Me.lblExchangeRate = New System.Windows.Forms.Label()
        Me.cmb10Percent = New System.Windows.Forms.ComboBox()
        Me.btnHelp = New System.Windows.Forms.Button()
        Me.cmbFilter = New System.Windows.Forms.ComboBox()
        Me.GroupApproved = New System.Windows.Forms.GroupBox()
        Me.lblApprovedMessage = New System.Windows.Forms.Label()
        Me.txbPasswordApprove = New System.Windows.Forms.TextBox()
        Me.txbUserApprove = New System.Windows.Forms.TextBox()
        Me.lblPasswordA = New System.Windows.Forms.Label()
        Me.lblUserIDA = New System.Windows.Forms.Label()
        Me.btnLoadMRP = New System.Windows.Forms.Button()
        Me.GroupBoxSaved = New System.Windows.Forms.GroupBox()
        Me.rdoViewOnly = New System.Windows.Forms.RadioButton()
        Me.rdoSaveReport = New System.Windows.Forms.RadioButton()
        Me.GridMRP = New System.Windows.Forms.DataGridView()
        Me.GroupBoxOption = New System.Windows.Forms.GroupBox()
        Me.GroupWipSalesOrder = New System.Windows.Forms.GroupBox()
        Me.lblRecordsSalesOrder = New System.Windows.Forms.Label()
        Me.lblRecordsWip = New System.Windows.Forms.Label()
        Me.lblTSalesOrder = New System.Windows.Forms.Label()
        Me.lblTWIP = New System.Windows.Forms.Label()
        Me.btnRefreshSalesOrders = New System.Windows.Forms.Button()
        Me.btnCloseAddIems = New System.Windows.Forms.Button()
        Me.GridSalesOrder = New System.Windows.Forms.DataGridView()
        Me.GridWIP = New System.Windows.Forms.DataGridView()
        Me.ckbConfirmed = New System.Windows.Forms.CheckBox()
        Me.ckbPastDue = New System.Windows.Forms.CheckBox()
        Me.lblWeekTo = New System.Windows.Forms.Label()
        Me.lblWeekFrom = New System.Windows.Forms.Label()
        Me.lblTo = New System.Windows.Forms.Label()
        Me.lblFrom = New System.Windows.Forms.Label()
        Me.dtpTo = New System.Windows.Forms.DateTimePicker()
        Me.dtpFrom = New System.Windows.Forms.DateTimePicker()
        Me.rdoSpecificDates = New System.Windows.Forms.RadioButton()
        Me.rdoAllWeeks = New System.Windows.Forms.RadioButton()
        Me.btnCalculate = New System.Windows.Forms.Button()
        Me.GroupBoxFind = New System.Windows.Forms.GroupBox()
        Me.lblFind = New System.Windows.Forms.Label()
        Me.txbFind = New System.Windows.Forms.TextBox()
        Me.btnFind = New System.Windows.Forms.Button()
        Me.btnExportToExcel = New System.Windows.Forms.Button()
        Me.lblQty = New System.Windows.Forms.Label()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.txbQty = New System.Windows.Forms.TextBox()
        Me.lblMRPReference = New System.Windows.Forms.Label()
        Me.lblTMRPReference = New System.Windows.Forms.Label()
        Me.lblRecordsMRP = New System.Windows.Forms.Label()
        Me.lblMRP = New System.Windows.Forms.Label()
        Me.TabPageBOMWIP = New System.Windows.Forms.TabPage()
        Me.GroupBoxAUBOMWIP = New System.Windows.Forms.GroupBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cmbWIPBOMWIP = New System.Windows.Forms.ComboBox()
        Me.cmbAUBOMWIP = New System.Windows.Forms.ComboBox()
        Me.cmbRevBOMWIP = New System.Windows.Forms.ComboBox()
        Me.lblTRevBOMWIP = New System.Windows.Forms.Label()
        Me.lblTAUBOMWIP = New System.Windows.Forms.Label()
        Me.GroupBoxPNBOMWIP = New System.Windows.Forms.GroupBox()
        Me.txbBOMWIP = New System.Windows.Forms.TextBox()
        Me.lblTPNBOMWIP = New System.Windows.Forms.Label()
        Me.btnFindBOMWIP = New System.Windows.Forms.Button()
        Me.lblRecordsBOMWIP = New System.Windows.Forms.Label()
        Me.GridBOMWIP = New System.Windows.Forms.DataGridView()
        Me.TabPageBOMENG = New System.Windows.Forms.TabPage()
        Me.GroupBoxByAUBOMENG = New System.Windows.Forms.GroupBox()
        Me.cmbAUBOMENG = New System.Windows.Forms.ComboBox()
        Me.cmbRevBOMENG = New System.Windows.Forms.ComboBox()
        Me.lblTAUBOMENG = New System.Windows.Forms.Label()
        Me.lblTRevBOMENG = New System.Windows.Forms.Label()
        Me.GroupBoxPNBOMENG = New System.Windows.Forms.GroupBox()
        Me.txbPNBOMENG = New System.Windows.Forms.TextBox()
        Me.lblTPNBOMENG = New System.Windows.Forms.Label()
        Me.btnFindBOMENG = New System.Windows.Forms.Button()
        Me.lblRecordsBOMENG = New System.Windows.Forms.Label()
        Me.GridBOMENG = New System.Windows.Forms.DataGridView()
        Me.TabPageMyTable = New System.Windows.Forms.TabPage()
        Me.GroupBoxPNMyTable = New System.Windows.Forms.GroupBox()
        Me.cmbPNMyTable = New System.Windows.Forms.ComboBox()
        Me.lblRecordsMyTable = New System.Windows.Forms.Label()
        Me.GridMyTable = New System.Windows.Forms.DataGridView()
        Me.TabPageSalesOrder = New System.Windows.Forms.TabPage()
        Me.GroupBoxSalesOrderControl = New System.Windows.Forms.GroupBox()
        Me.GroupBoxSalesOrderStatus = New System.Windows.Forms.GroupBox()
        Me.rdoAllSalesOrderByAU = New System.Windows.Forms.RadioButton()
        Me.rdoOpenSalesOrderByAU = New System.Windows.Forms.RadioButton()
        Me.rdoCancelSalesOrderByAU = New System.Windows.Forms.RadioButton()
        Me.rdoCloseSalesOrderByAU = New System.Windows.Forms.RadioButton()
        Me.btnFindSalesOrder = New System.Windows.Forms.Button()
        Me.lblTAUSalesOrder = New System.Windows.Forms.Label()
        Me.lblTrevSalesOrder = New System.Windows.Forms.Label()
        Me.txbAUSalesOrder = New System.Windows.Forms.TextBox()
        Me.cmbRevSalesOrder = New System.Windows.Forms.ComboBox()
        Me.lblRecordsGridSalesOrder = New System.Windows.Forms.Label()
        Me.GridAUSalesOrderFind = New System.Windows.Forms.DataGridView()
        Me.TabPageWIPByAU = New System.Windows.Forms.TabPage()
        Me.GroupBoxWIPByAU = New System.Windows.Forms.GroupBox()
        Me.GroupBoxStatusWIPByAU = New System.Windows.Forms.GroupBox()
        Me.rdoAllWipByAU = New System.Windows.Forms.RadioButton()
        Me.rdoOpenWipByAU = New System.Windows.Forms.RadioButton()
        Me.rdoCancelWipByAU = New System.Windows.Forms.RadioButton()
        Me.rdoCloseWipByAU = New System.Windows.Forms.RadioButton()
        Me.btnFindWipByAU = New System.Windows.Forms.Button()
        Me.lblTAUWipByAU = New System.Windows.Forms.Label()
        Me.lblTRevWipByAU = New System.Windows.Forms.Label()
        Me.txbAUWipByAU = New System.Windows.Forms.TextBox()
        Me.cmbRevWipByAU = New System.Windows.Forms.ComboBox()
        Me.lblRecordsWipByAU = New System.Windows.Forms.Label()
        Me.GridWipByAU = New System.Windows.Forms.DataGridView()
        Me.GroupBoxUserMRP = New System.Windows.Forms.GroupBox()
        Me.btnCancelLoginEng = New System.Windows.Forms.Button()
        Me.btnLoginMRP = New System.Windows.Forms.Button()
        Me.txbUserMRP = New System.Windows.Forms.TextBox()
        Me.lblTEngPassword = New System.Windows.Forms.Label()
        Me.txbUserMRPPassword = New System.Windows.Forms.TextBox()
        Me.lblTEngUser = New System.Windows.Forms.Label()
        Me.cmbPONoAprovadas = New System.Windows.Forms.ComboBox()
        Me.txbUser = New System.Windows.Forms.TextBox()
        Me.TabControlMRPGlobal.SuspendLayout()
        Me.TabPageMRP.SuspendLayout()
        Me.GroupBoxBudgetInformation.SuspendLayout()
        CType(Me.GridPerVendor, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridPerWeek, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxPurchasingOrderHistory.SuspendLayout()
        CType(Me.GridPurchasingOrderItemsHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupApproved.SuspendLayout()
        Me.GroupBoxSaved.SuspendLayout()
        CType(Me.GridMRP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxOption.SuspendLayout()
        Me.GroupWipSalesOrder.SuspendLayout()
        CType(Me.GridSalesOrder, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.GridWIP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxFind.SuspendLayout()
        Me.TabPageBOMWIP.SuspendLayout()
        Me.GroupBoxAUBOMWIP.SuspendLayout()
        Me.GroupBoxPNBOMWIP.SuspendLayout()
        CType(Me.GridBOMWIP, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageBOMENG.SuspendLayout()
        Me.GroupBoxByAUBOMENG.SuspendLayout()
        Me.GroupBoxPNBOMENG.SuspendLayout()
        CType(Me.GridBOMENG, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageMyTable.SuspendLayout()
        Me.GroupBoxPNMyTable.SuspendLayout()
        CType(Me.GridMyTable, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageSalesOrder.SuspendLayout()
        Me.GroupBoxSalesOrderControl.SuspendLayout()
        Me.GroupBoxSalesOrderStatus.SuspendLayout()
        CType(Me.GridAUSalesOrderFind, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPageWIPByAU.SuspendLayout()
        Me.GroupBoxWIPByAU.SuspendLayout()
        Me.GroupBoxStatusWIPByAU.SuspendLayout()
        CType(Me.GridWipByAU, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBoxUserMRP.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControlMRPGlobal
        '
        Me.TabControlMRPGlobal.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageMRP)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageBOMWIP)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageBOMENG)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageMyTable)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageSalesOrder)
        Me.TabControlMRPGlobal.Controls.Add(Me.TabPageWIPByAU)
        Me.TabControlMRPGlobal.Location = New System.Drawing.Point(8, 3)
        Me.TabControlMRPGlobal.Margin = New System.Windows.Forms.Padding(2)
        Me.TabControlMRPGlobal.Name = "TabControlMRPGlobal"
        Me.TabControlMRPGlobal.SelectedIndex = 0
        Me.TabControlMRPGlobal.Size = New System.Drawing.Size(934, 478)
        Me.TabControlMRPGlobal.TabIndex = 5328
        Me.TabControlMRPGlobal.Visible = False
        '
        'TabPageMRP
        '
        Me.TabPageMRP.Controls.Add(Me.GroupBoxBudgetInformation)
        Me.TabPageMRP.Controls.Add(Me.GroupBoxPurchasingOrderHistory)
        Me.TabPageMRP.Controls.Add(Me.lblTotal)
        Me.TabPageMRP.Controls.Add(Me.txbExchangeRate)
        Me.TabPageMRP.Controls.Add(Me.lblExchangeRate)
        Me.TabPageMRP.Controls.Add(Me.cmb10Percent)
        Me.TabPageMRP.Controls.Add(Me.btnHelp)
        Me.TabPageMRP.Controls.Add(Me.cmbFilter)
        Me.TabPageMRP.Controls.Add(Me.GroupApproved)
        Me.TabPageMRP.Controls.Add(Me.btnLoadMRP)
        Me.TabPageMRP.Controls.Add(Me.GroupBoxSaved)
        Me.TabPageMRP.Controls.Add(Me.GridMRP)
        Me.TabPageMRP.Controls.Add(Me.GroupBoxOption)
        Me.TabPageMRP.Controls.Add(Me.GroupBoxFind)
        Me.TabPageMRP.Controls.Add(Me.btnExportToExcel)
        Me.TabPageMRP.Controls.Add(Me.lblQty)
        Me.TabPageMRP.Controls.Add(Me.btnClear)
        Me.TabPageMRP.Controls.Add(Me.txbQty)
        Me.TabPageMRP.Controls.Add(Me.lblMRPReference)
        Me.TabPageMRP.Controls.Add(Me.lblTMRPReference)
        Me.TabPageMRP.Controls.Add(Me.lblRecordsMRP)
        Me.TabPageMRP.Controls.Add(Me.lblMRP)
        Me.TabPageMRP.Location = New System.Drawing.Point(4, 22)
        Me.TabPageMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageMRP.Name = "TabPageMRP"
        Me.TabPageMRP.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageMRP.Size = New System.Drawing.Size(926, 452)
        Me.TabPageMRP.TabIndex = 0
        Me.TabPageMRP.Text = "MRP Report"
        Me.TabPageMRP.UseVisualStyleBackColor = True
        '
        'GroupBoxBudgetInformation
        '
        Me.GroupBoxBudgetInformation.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.lblTotalTotal2)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.lblRecordsPerVendor)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.lblRecordsPerWeek)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.lblTPerVendor)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.lblTPerWeek)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.btnCloseBudgetInformation)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.GridPerVendor)
        Me.GroupBoxBudgetInformation.Controls.Add(Me.GridPerWeek)
        Me.GroupBoxBudgetInformation.Location = New System.Drawing.Point(2, 10)
        Me.GroupBoxBudgetInformation.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxBudgetInformation.Name = "GroupBoxBudgetInformation"
        Me.GroupBoxBudgetInformation.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxBudgetInformation.Size = New System.Drawing.Size(1076, 470)
        Me.GroupBoxBudgetInformation.TabIndex = 5320
        Me.GroupBoxBudgetInformation.TabStop = False
        Me.GroupBoxBudgetInformation.Text = "Budget Information "
        Me.GroupBoxBudgetInformation.Visible = False
        '
        'lblTotalTotal2
        '
        Me.lblTotalTotal2.AutoSize = True
        Me.lblTotalTotal2.Location = New System.Drawing.Point(262, 15)
        Me.lblTotalTotal2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotalTotal2.Name = "lblTotalTotal2"
        Me.lblTotalTotal2.Size = New System.Drawing.Size(43, 13)
        Me.lblTotalTotal2.TabIndex = 1116
        Me.lblTotalTotal2.Text = "Total: 0"
        '
        'lblRecordsPerVendor
        '
        Me.lblRecordsPerVendor.AutoSize = True
        Me.lblRecordsPerVendor.Location = New System.Drawing.Point(500, 33)
        Me.lblRecordsPerVendor.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsPerVendor.Name = "lblRecordsPerVendor"
        Me.lblRecordsPerVendor.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsPerVendor.TabIndex = 1115
        Me.lblRecordsPerVendor.Text = "Records: 0"
        '
        'lblRecordsPerWeek
        '
        Me.lblRecordsPerWeek.AutoSize = True
        Me.lblRecordsPerWeek.Location = New System.Drawing.Point(118, 33)
        Me.lblRecordsPerWeek.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsPerWeek.Name = "lblRecordsPerWeek"
        Me.lblRecordsPerWeek.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsPerWeek.TabIndex = 1114
        Me.lblRecordsPerWeek.Text = "Records: 0"
        '
        'lblTPerVendor
        '
        Me.lblTPerVendor.AutoSize = True
        Me.lblTPerVendor.Location = New System.Drawing.Point(396, 33)
        Me.lblTPerVendor.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTPerVendor.Name = "lblTPerVendor"
        Me.lblTPerVendor.Size = New System.Drawing.Size(63, 13)
        Me.lblTPerVendor.TabIndex = 1113
        Me.lblTPerVendor.Text = "Per Vendor:"
        '
        'lblTPerWeek
        '
        Me.lblTPerWeek.AutoSize = True
        Me.lblTPerWeek.Location = New System.Drawing.Point(4, 33)
        Me.lblTPerWeek.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTPerWeek.Name = "lblTPerWeek"
        Me.lblTPerWeek.Size = New System.Drawing.Size(58, 13)
        Me.lblTPerWeek.TabIndex = 1112
        Me.lblTPerWeek.Text = "Per Week:"
        '
        'btnCloseBudgetInformation
        '
        Me.btnCloseBudgetInformation.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCloseBudgetInformation.Image = CType(resources.GetObject("btnCloseBudgetInformation.Image"), System.Drawing.Image)
        Me.btnCloseBudgetInformation.Location = New System.Drawing.Point(979, 15)
        Me.btnCloseBudgetInformation.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCloseBudgetInformation.Name = "btnCloseBudgetInformation"
        Me.btnCloseBudgetInformation.Size = New System.Drawing.Size(34, 31)
        Me.btnCloseBudgetInformation.TabIndex = 553
        Me.btnCloseBudgetInformation.UseVisualStyleBackColor = True
        '
        'GridPerVendor
        '
        Me.GridPerVendor.AllowUserToAddRows = False
        Me.GridPerVendor.AllowUserToDeleteRows = False
        Me.GridPerVendor.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridPerVendor.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPerVendor.Location = New System.Drawing.Point(395, 50)
        Me.GridPerVendor.Margin = New System.Windows.Forms.Padding(2)
        Me.GridPerVendor.Name = "GridPerVendor"
        Me.GridPerVendor.RowTemplate.Height = 24
        Me.GridPerVendor.Size = New System.Drawing.Size(674, 180)
        Me.GridPerVendor.TabIndex = 1
        '
        'GridPerWeek
        '
        Me.GridPerWeek.AllowUserToAddRows = False
        Me.GridPerWeek.AllowUserToDeleteRows = False
        Me.GridPerWeek.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridPerWeek.Location = New System.Drawing.Point(3, 50)
        Me.GridPerWeek.Margin = New System.Windows.Forms.Padding(2)
        Me.GridPerWeek.Name = "GridPerWeek"
        Me.GridPerWeek.RowTemplate.Height = 24
        Me.GridPerWeek.Size = New System.Drawing.Size(369, 180)
        Me.GridPerWeek.TabIndex = 0
        '
        'GroupBoxPurchasingOrderHistory
        '
        Me.GroupBoxPurchasingOrderHistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBoxPurchasingOrderHistory.Controls.Add(Me.btnRefreshPurchasingOrderItemsHistory)
        Me.GroupBoxPurchasingOrderHistory.Controls.Add(Me.btnClosePurchasingOrderItemsHistory)
        Me.GroupBoxPurchasingOrderHistory.Controls.Add(Me.lblRecordsPurchasingOrderItemsHistory)
        Me.GroupBoxPurchasingOrderHistory.Controls.Add(Me.lblTItems)
        Me.GroupBoxPurchasingOrderHistory.Controls.Add(Me.GridPurchasingOrderItemsHistory)
        Me.GroupBoxPurchasingOrderHistory.Location = New System.Drawing.Point(6, 15)
        Me.GroupBoxPurchasingOrderHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPurchasingOrderHistory.Name = "GroupBoxPurchasingOrderHistory"
        Me.GroupBoxPurchasingOrderHistory.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPurchasingOrderHistory.Size = New System.Drawing.Size(924, 444)
        Me.GroupBoxPurchasingOrderHistory.TabIndex = 5320
        Me.GroupBoxPurchasingOrderHistory.TabStop = False
        Me.GroupBoxPurchasingOrderHistory.Text = "Purchasing Order History:"
        Me.GroupBoxPurchasingOrderHistory.Visible = False
        '
        'btnRefreshPurchasingOrderItemsHistory
        '
        Me.btnRefreshPurchasingOrderItemsHistory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefreshPurchasingOrderItemsHistory.Image = CType(resources.GetObject("btnRefreshPurchasingOrderItemsHistory.Image"), System.Drawing.Image)
        Me.btnRefreshPurchasingOrderItemsHistory.Location = New System.Drawing.Point(824, 15)
        Me.btnRefreshPurchasingOrderItemsHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRefreshPurchasingOrderItemsHistory.Name = "btnRefreshPurchasingOrderItemsHistory"
        Me.btnRefreshPurchasingOrderItemsHistory.Size = New System.Drawing.Size(28, 32)
        Me.btnRefreshPurchasingOrderItemsHistory.TabIndex = 5313
        Me.btnRefreshPurchasingOrderItemsHistory.UseVisualStyleBackColor = True
        Me.btnRefreshPurchasingOrderItemsHistory.Visible = False
        '
        'btnClosePurchasingOrderItemsHistory
        '
        Me.btnClosePurchasingOrderItemsHistory.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClosePurchasingOrderItemsHistory.Image = CType(resources.GetObject("btnClosePurchasingOrderItemsHistory.Image"), System.Drawing.Image)
        Me.btnClosePurchasingOrderItemsHistory.Location = New System.Drawing.Point(856, 15)
        Me.btnClosePurchasingOrderItemsHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClosePurchasingOrderItemsHistory.Name = "btnClosePurchasingOrderItemsHistory"
        Me.btnClosePurchasingOrderItemsHistory.Size = New System.Drawing.Size(34, 31)
        Me.btnClosePurchasingOrderItemsHistory.TabIndex = 5312
        Me.btnClosePurchasingOrderItemsHistory.UseVisualStyleBackColor = True
        '
        'lblRecordsPurchasingOrderItemsHistory
        '
        Me.lblRecordsPurchasingOrderItemsHistory.AutoSize = True
        Me.lblRecordsPurchasingOrderItemsHistory.Location = New System.Drawing.Point(166, 37)
        Me.lblRecordsPurchasingOrderItemsHistory.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsPurchasingOrderItemsHistory.Name = "lblRecordsPurchasingOrderItemsHistory"
        Me.lblRecordsPurchasingOrderItemsHistory.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsPurchasingOrderItemsHistory.TabIndex = 5311
        Me.lblRecordsPurchasingOrderItemsHistory.Text = "Records: 0"
        '
        'lblTItems
        '
        Me.lblTItems.AutoSize = True
        Me.lblTItems.Location = New System.Drawing.Point(4, 37)
        Me.lblTItems.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTItems.Name = "lblTItems"
        Me.lblTItems.Size = New System.Drawing.Size(35, 13)
        Me.lblTItems.TabIndex = 5306
        Me.lblTItems.Text = "Items:"
        '
        'GridPurchasingOrderItemsHistory
        '
        Me.GridPurchasingOrderItemsHistory.AllowUserToAddRows = False
        Me.GridPurchasingOrderItemsHistory.AllowUserToDeleteRows = False
        Me.GridPurchasingOrderItemsHistory.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridPurchasingOrderItemsHistory.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.GridPurchasingOrderItemsHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridPurchasingOrderItemsHistory.DefaultCellStyle = DataGridViewCellStyle2
        Me.GridPurchasingOrderItemsHistory.Location = New System.Drawing.Point(4, 54)
        Me.GridPurchasingOrderItemsHistory.Margin = New System.Windows.Forms.Padding(2)
        Me.GridPurchasingOrderItemsHistory.Name = "GridPurchasingOrderItemsHistory"
        Me.GridPurchasingOrderItemsHistory.ReadOnly = True
        Me.GridPurchasingOrderItemsHistory.RowTemplate.Height = 24
        Me.GridPurchasingOrderItemsHistory.Size = New System.Drawing.Size(913, 377)
        Me.GridPurchasingOrderItemsHistory.TabIndex = 5308
        '
        'lblTotal
        '
        Me.lblTotal.AutoSize = True
        Me.lblTotal.Location = New System.Drawing.Point(604, 42)
        Me.lblTotal.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(43, 13)
        Me.lblTotal.TabIndex = 5326
        Me.lblTotal.TabStop = True
        Me.lblTotal.Text = "Total: 0"
        '
        'txbExchangeRate
        '
        Me.txbExchangeRate.Location = New System.Drawing.Point(831, 80)
        Me.txbExchangeRate.Margin = New System.Windows.Forms.Padding(2)
        Me.txbExchangeRate.Name = "txbExchangeRate"
        Me.txbExchangeRate.Size = New System.Drawing.Size(41, 20)
        Me.txbExchangeRate.TabIndex = 5324
        '
        'lblExchangeRate
        '
        Me.lblExchangeRate.AutoSize = True
        Me.lblExchangeRate.Location = New System.Drawing.Point(754, 82)
        Me.lblExchangeRate.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblExchangeRate.Name = "lblExchangeRate"
        Me.lblExchangeRate.Size = New System.Drawing.Size(84, 13)
        Me.lblExchangeRate.TabIndex = 5323
        Me.lblExchangeRate.Text = "Exchange Rate:"
        '
        'cmb10Percent
        '
        Me.cmb10Percent.AutoCompleteCustomSource.AddRange(New String() {"Primary Without Bin Balance", "Primary With Bin Balance", "No Primary Without Bin Balance", "No Primary With Bin Balance", "ALL", "Only Bin Balance"})
        Me.cmb10Percent.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.AllSystemSources
        Me.cmb10Percent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmb10Percent.FormattingEnabled = True
        Me.cmb10Percent.Items.AddRange(New Object() {"ALL", "10%"})
        Me.cmb10Percent.Location = New System.Drawing.Point(210, 85)
        Me.cmb10Percent.Margin = New System.Windows.Forms.Padding(2)
        Me.cmb10Percent.Name = "cmb10Percent"
        Me.cmb10Percent.Size = New System.Drawing.Size(58, 21)
        Me.cmb10Percent.TabIndex = 5322
        '
        'btnHelp
        '
        Me.btnHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnHelp.Image = CType(resources.GetObject("btnHelp.Image"), System.Drawing.Image)
        Me.btnHelp.Location = New System.Drawing.Point(884, 62)
        Me.btnHelp.Margin = New System.Windows.Forms.Padding(2)
        Me.btnHelp.Name = "btnHelp"
        Me.btnHelp.Size = New System.Drawing.Size(40, 40)
        Me.btnHelp.TabIndex = 5321
        Me.btnHelp.UseVisualStyleBackColor = True
        '
        'cmbFilter
        '
        Me.cmbFilter.AutoCompleteCustomSource.AddRange(New String() {"Primary Without Bin Balance", "Primary With Bin Balance", "No Primary Without Bin Balance", "No Primary With Bin Balance", "ALL", "Only Bin Balance"})
        Me.cmbFilter.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.AllSystemSources
        Me.cmbFilter.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFilter.FormattingEnabled = True
        Me.cmbFilter.Items.AddRange(New Object() {"Only Primary Without Bin Balance", "Only Primary With Bin Balance", "All Without Bin Balance", "ALL", "Only Bin Balance"})
        Me.cmbFilter.Location = New System.Drawing.Point(40, 85)
        Me.cmbFilter.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbFilter.Name = "cmbFilter"
        Me.cmbFilter.Size = New System.Drawing.Size(166, 21)
        Me.cmbFilter.TabIndex = 5320
        '
        'GroupApproved
        '
        Me.GroupApproved.Controls.Add(Me.lblApprovedMessage)
        Me.GroupApproved.Controls.Add(Me.txbPasswordApprove)
        Me.GroupApproved.Controls.Add(Me.txbUserApprove)
        Me.GroupApproved.Controls.Add(Me.lblPasswordA)
        Me.GroupApproved.Controls.Add(Me.lblUserIDA)
        Me.GroupApproved.Location = New System.Drawing.Point(330, 203)
        Me.GroupApproved.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupApproved.Name = "GroupApproved"
        Me.GroupApproved.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupApproved.Size = New System.Drawing.Size(246, 116)
        Me.GroupApproved.TabIndex = 5317
        Me.GroupApproved.TabStop = False
        Me.GroupApproved.Text = "Approved By"
        Me.GroupApproved.Visible = False
        '
        'lblApprovedMessage
        '
        Me.lblApprovedMessage.AutoSize = True
        Me.lblApprovedMessage.Location = New System.Drawing.Point(48, 65)
        Me.lblApprovedMessage.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblApprovedMessage.Name = "lblApprovedMessage"
        Me.lblApprovedMessage.Size = New System.Drawing.Size(153, 39)
        Me.lblApprovedMessage.TabIndex = 5
        Me.lblApprovedMessage.Text = "To approve, please,  add your " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "  Windows  account-password " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "        and press E" &
    "nter key  "
        '
        'txbPasswordApprove
        '
        Me.txbPasswordApprove.Location = New System.Drawing.Point(80, 38)
        Me.txbPasswordApprove.Margin = New System.Windows.Forms.Padding(2)
        Me.txbPasswordApprove.Name = "txbPasswordApprove"
        Me.txbPasswordApprove.Size = New System.Drawing.Size(122, 20)
        Me.txbPasswordApprove.TabIndex = 14
        Me.txbPasswordApprove.UseSystemPasswordChar = True
        '
        'txbUserApprove
        '
        Me.txbUserApprove.Location = New System.Drawing.Point(80, 15)
        Me.txbUserApprove.Margin = New System.Windows.Forms.Padding(2)
        Me.txbUserApprove.Name = "txbUserApprove"
        Me.txbUserApprove.Size = New System.Drawing.Size(122, 20)
        Me.txbUserApprove.TabIndex = 13
        '
        'lblPasswordA
        '
        Me.lblPasswordA.AutoSize = True
        Me.lblPasswordA.Location = New System.Drawing.Point(20, 41)
        Me.lblPasswordA.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPasswordA.Name = "lblPasswordA"
        Me.lblPasswordA.Size = New System.Drawing.Size(56, 13)
        Me.lblPasswordA.TabIndex = 1
        Me.lblPasswordA.Text = "Password:"
        '
        'lblUserIDA
        '
        Me.lblUserIDA.AutoSize = True
        Me.lblUserIDA.Location = New System.Drawing.Point(34, 18)
        Me.lblUserIDA.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblUserIDA.Name = "lblUserIDA"
        Me.lblUserIDA.Size = New System.Drawing.Size(43, 13)
        Me.lblUserIDA.TabIndex = 0
        Me.lblUserIDA.Text = "UserID:"
        '
        'btnLoadMRP
        '
        Me.btnLoadMRP.BackColor = System.Drawing.SystemColors.Control
        Me.btnLoadMRP.Location = New System.Drawing.Point(604, 76)
        Me.btnLoadMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoadMRP.Name = "btnLoadMRP"
        Me.btnLoadMRP.Size = New System.Drawing.Size(65, 25)
        Me.btnLoadMRP.TabIndex = 5318
        Me.btnLoadMRP.Text = "Load MRP"
        Me.btnLoadMRP.UseVisualStyleBackColor = False
        '
        'GroupBoxSaved
        '
        Me.GroupBoxSaved.Controls.Add(Me.rdoViewOnly)
        Me.GroupBoxSaved.Controls.Add(Me.rdoSaveReport)
        Me.GroupBoxSaved.Location = New System.Drawing.Point(510, 39)
        Me.GroupBoxSaved.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSaved.Name = "GroupBoxSaved"
        Me.GroupBoxSaved.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSaved.Size = New System.Drawing.Size(89, 65)
        Me.GroupBoxSaved.TabIndex = 5317
        Me.GroupBoxSaved.TabStop = False
        '
        'rdoViewOnly
        '
        Me.rdoViewOnly.AutoSize = True
        Me.rdoViewOnly.Location = New System.Drawing.Point(4, 38)
        Me.rdoViewOnly.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoViewOnly.Name = "rdoViewOnly"
        Me.rdoViewOnly.Size = New System.Drawing.Size(72, 17)
        Me.rdoViewOnly.TabIndex = 1
        Me.rdoViewOnly.Text = "View Only"
        Me.rdoViewOnly.UseVisualStyleBackColor = True
        '
        'rdoSaveReport
        '
        Me.rdoSaveReport.AutoSize = True
        Me.rdoSaveReport.Checked = True
        Me.rdoSaveReport.Location = New System.Drawing.Point(4, 14)
        Me.rdoSaveReport.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoSaveReport.Name = "rdoSaveReport"
        Me.rdoSaveReport.Size = New System.Drawing.Size(85, 17)
        Me.rdoSaveReport.TabIndex = 0
        Me.rdoSaveReport.TabStop = True
        Me.rdoSaveReport.Text = "Save Report"
        Me.rdoSaveReport.UseVisualStyleBackColor = True
        '
        'GridMRP
        '
        Me.GridMRP.AllowUserToAddRows = False
        Me.GridMRP.AllowUserToDeleteRows = False
        Me.GridMRP.AllowUserToOrderColumns = True
        Me.GridMRP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.GridMRP.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.GridMRP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle4.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridMRP.DefaultCellStyle = DataGridViewCellStyle4
        Me.GridMRP.Location = New System.Drawing.Point(4, 111)
        Me.GridMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.GridMRP.Name = "GridMRP"
        Me.GridMRP.RowTemplate.Height = 24
        Me.GridMRP.Size = New System.Drawing.Size(920, 324)
        Me.GridMRP.TabIndex = 5299
        '
        'GroupBoxOption
        '
        Me.GroupBoxOption.Controls.Add(Me.GroupWipSalesOrder)
        Me.GroupBoxOption.Controls.Add(Me.ckbConfirmed)
        Me.GroupBoxOption.Controls.Add(Me.ckbPastDue)
        Me.GroupBoxOption.Controls.Add(Me.lblWeekTo)
        Me.GroupBoxOption.Controls.Add(Me.lblWeekFrom)
        Me.GroupBoxOption.Controls.Add(Me.lblTo)
        Me.GroupBoxOption.Controls.Add(Me.lblFrom)
        Me.GroupBoxOption.Controls.Add(Me.dtpTo)
        Me.GroupBoxOption.Controls.Add(Me.dtpFrom)
        Me.GroupBoxOption.Controls.Add(Me.rdoSpecificDates)
        Me.GroupBoxOption.Controls.Add(Me.rdoAllWeeks)
        Me.GroupBoxOption.Controls.Add(Me.btnCalculate)
        Me.GroupBoxOption.Location = New System.Drawing.Point(4, 5)
        Me.GroupBoxOption.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxOption.Name = "GroupBoxOption"
        Me.GroupBoxOption.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxOption.Size = New System.Drawing.Size(400, 77)
        Me.GroupBoxOption.TabIndex = 5290
        Me.GroupBoxOption.TabStop = False
        Me.GroupBoxOption.Text = "Option"
        '
        'GroupWipSalesOrder
        '
        Me.GroupWipSalesOrder.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupWipSalesOrder.Controls.Add(Me.lblRecordsSalesOrder)
        Me.GroupWipSalesOrder.Controls.Add(Me.lblRecordsWip)
        Me.GroupWipSalesOrder.Controls.Add(Me.lblTSalesOrder)
        Me.GroupWipSalesOrder.Controls.Add(Me.lblTWIP)
        Me.GroupWipSalesOrder.Controls.Add(Me.btnRefreshSalesOrders)
        Me.GroupWipSalesOrder.Controls.Add(Me.btnCloseAddIems)
        Me.GroupWipSalesOrder.Controls.Add(Me.GridSalesOrder)
        Me.GroupWipSalesOrder.Controls.Add(Me.GridWIP)
        Me.GroupWipSalesOrder.Location = New System.Drawing.Point(4, 4)
        Me.GroupWipSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupWipSalesOrder.Name = "GroupWipSalesOrder"
        Me.GroupWipSalesOrder.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupWipSalesOrder.Size = New System.Drawing.Size(922, 444)
        Me.GroupWipSalesOrder.TabIndex = 5319
        Me.GroupWipSalesOrder.TabStop = False
        Me.GroupWipSalesOrder.Text = "Information"
        Me.GroupWipSalesOrder.Visible = False
        '
        'lblRecordsSalesOrder
        '
        Me.lblRecordsSalesOrder.AutoSize = True
        Me.lblRecordsSalesOrder.Location = New System.Drawing.Point(118, 272)
        Me.lblRecordsSalesOrder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsSalesOrder.Name = "lblRecordsSalesOrder"
        Me.lblRecordsSalesOrder.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsSalesOrder.TabIndex = 1115
        Me.lblRecordsSalesOrder.Text = "Records: 0"
        '
        'lblRecordsWip
        '
        Me.lblRecordsWip.AutoSize = True
        Me.lblRecordsWip.Location = New System.Drawing.Point(118, 32)
        Me.lblRecordsWip.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsWip.Name = "lblRecordsWip"
        Me.lblRecordsWip.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsWip.TabIndex = 1114
        Me.lblRecordsWip.Text = "Records: 0"
        '
        'lblTSalesOrder
        '
        Me.lblTSalesOrder.AutoSize = True
        Me.lblTSalesOrder.Location = New System.Drawing.Point(4, 272)
        Me.lblTSalesOrder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTSalesOrder.Name = "lblTSalesOrder"
        Me.lblTSalesOrder.Size = New System.Drawing.Size(65, 13)
        Me.lblTSalesOrder.TabIndex = 1113
        Me.lblTSalesOrder.Text = "Sales Order:"
        '
        'lblTWIP
        '
        Me.lblTWIP.AutoSize = True
        Me.lblTWIP.Location = New System.Drawing.Point(4, 33)
        Me.lblTWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTWIP.Name = "lblTWIP"
        Me.lblTWIP.Size = New System.Drawing.Size(31, 13)
        Me.lblTWIP.TabIndex = 1112
        Me.lblTWIP.Text = "WIP:"
        '
        'btnRefreshSalesOrders
        '
        Me.btnRefreshSalesOrders.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnRefreshSalesOrders.Image = CType(resources.GetObject("btnRefreshSalesOrders.Image"), System.Drawing.Image)
        Me.btnRefreshSalesOrders.Location = New System.Drawing.Point(793, 15)
        Me.btnRefreshSalesOrders.Margin = New System.Windows.Forms.Padding(2)
        Me.btnRefreshSalesOrders.Name = "btnRefreshSalesOrders"
        Me.btnRefreshSalesOrders.Size = New System.Drawing.Size(28, 32)
        Me.btnRefreshSalesOrders.TabIndex = 1111
        Me.btnRefreshSalesOrders.UseVisualStyleBackColor = True
        '
        'btnCloseAddIems
        '
        Me.btnCloseAddIems.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnCloseAddIems.Image = CType(resources.GetObject("btnCloseAddIems.Image"), System.Drawing.Image)
        Me.btnCloseAddIems.Location = New System.Drawing.Point(825, 15)
        Me.btnCloseAddIems.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCloseAddIems.Name = "btnCloseAddIems"
        Me.btnCloseAddIems.Size = New System.Drawing.Size(34, 31)
        Me.btnCloseAddIems.TabIndex = 553
        Me.btnCloseAddIems.UseVisualStyleBackColor = True
        '
        'GridSalesOrder
        '
        Me.GridSalesOrder.AllowUserToAddRows = False
        Me.GridSalesOrder.AllowUserToDeleteRows = False
        Me.GridSalesOrder.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridSalesOrder.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridSalesOrder.Location = New System.Drawing.Point(2, 288)
        Me.GridSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.GridSalesOrder.Name = "GridSalesOrder"
        Me.GridSalesOrder.RowTemplate.Height = 24
        Me.GridSalesOrder.Size = New System.Drawing.Size(915, 137)
        Me.GridSalesOrder.TabIndex = 1
        '
        'GridWIP
        '
        Me.GridWIP.AllowUserToAddRows = False
        Me.GridWIP.AllowUserToDeleteRows = False
        Me.GridWIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridWIP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridWIP.Location = New System.Drawing.Point(3, 50)
        Me.GridWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.GridWIP.Name = "GridWIP"
        Me.GridWIP.RowTemplate.Height = 24
        Me.GridWIP.Size = New System.Drawing.Size(914, 220)
        Me.GridWIP.TabIndex = 0
        '
        'ckbConfirmed
        '
        Me.ckbConfirmed.AutoSize = True
        Me.ckbConfirmed.Location = New System.Drawing.Point(326, 8)
        Me.ckbConfirmed.Margin = New System.Windows.Forms.Padding(2)
        Me.ckbConfirmed.Name = "ckbConfirmed"
        Me.ckbConfirmed.Size = New System.Drawing.Size(73, 17)
        Me.ckbConfirmed.TabIndex = 9
        Me.ckbConfirmed.Text = "Confirmed"
        Me.ckbConfirmed.UseVisualStyleBackColor = True
        '
        'ckbPastDue
        '
        Me.ckbPastDue.AutoSize = True
        Me.ckbPastDue.Location = New System.Drawing.Point(326, 26)
        Me.ckbPastDue.Margin = New System.Windows.Forms.Padding(2)
        Me.ckbPastDue.Name = "ckbPastDue"
        Me.ckbPastDue.Size = New System.Drawing.Size(70, 17)
        Me.ckbPastDue.TabIndex = 8
        Me.ckbPastDue.Text = "Past Due"
        Me.ckbPastDue.UseVisualStyleBackColor = True
        '
        'lblWeekTo
        '
        Me.lblWeekTo.AutoSize = True
        Me.lblWeekTo.Location = New System.Drawing.Point(304, 51)
        Me.lblWeekTo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblWeekTo.Name = "lblWeekTo"
        Me.lblWeekTo.Size = New System.Drawing.Size(10, 13)
        Me.lblWeekTo.TabIndex = 7
        Me.lblWeekTo.Text = "-"
        '
        'lblWeekFrom
        '
        Me.lblWeekFrom.AutoSize = True
        Me.lblWeekFrom.Location = New System.Drawing.Point(304, 24)
        Me.lblWeekFrom.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblWeekFrom.Name = "lblWeekFrom"
        Me.lblWeekFrom.Size = New System.Drawing.Size(10, 13)
        Me.lblWeekFrom.TabIndex = 6
        Me.lblWeekFrom.Text = "-"
        '
        'lblTo
        '
        Me.lblTo.AutoSize = True
        Me.lblTo.Location = New System.Drawing.Point(92, 51)
        Me.lblTo.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTo.Name = "lblTo"
        Me.lblTo.Size = New System.Drawing.Size(23, 13)
        Me.lblTo.TabIndex = 5
        Me.lblTo.Text = "To:"
        '
        'lblFrom
        '
        Me.lblFrom.AutoSize = True
        Me.lblFrom.Location = New System.Drawing.Point(81, 24)
        Me.lblFrom.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFrom.Name = "lblFrom"
        Me.lblFrom.Size = New System.Drawing.Size(33, 13)
        Me.lblFrom.TabIndex = 4
        Me.lblFrom.Text = "From:"
        '
        'dtpTo
        '
        Me.dtpTo.Enabled = False
        Me.dtpTo.Location = New System.Drawing.Point(118, 50)
        Me.dtpTo.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpTo.Name = "dtpTo"
        Me.dtpTo.Size = New System.Drawing.Size(182, 20)
        Me.dtpTo.TabIndex = 3
        '
        'dtpFrom
        '
        Me.dtpFrom.Enabled = False
        Me.dtpFrom.Location = New System.Drawing.Point(118, 23)
        Me.dtpFrom.Margin = New System.Windows.Forms.Padding(2)
        Me.dtpFrom.Name = "dtpFrom"
        Me.dtpFrom.Size = New System.Drawing.Size(182, 20)
        Me.dtpFrom.TabIndex = 2
        '
        'rdoSpecificDates
        '
        Me.rdoSpecificDates.AutoSize = True
        Me.rdoSpecificDates.Location = New System.Drawing.Point(4, 47)
        Me.rdoSpecificDates.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoSpecificDates.Name = "rdoSpecificDates"
        Me.rdoSpecificDates.Size = New System.Drawing.Size(94, 17)
        Me.rdoSpecificDates.TabIndex = 1
        Me.rdoSpecificDates.Text = "Specific Dates"
        Me.rdoSpecificDates.UseVisualStyleBackColor = True
        '
        'rdoAllWeeks
        '
        Me.rdoAllWeeks.AutoSize = True
        Me.rdoAllWeeks.Checked = True
        Me.rdoAllWeeks.Location = New System.Drawing.Point(4, 27)
        Me.rdoAllWeeks.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoAllWeeks.Name = "rdoAllWeeks"
        Me.rdoAllWeeks.Size = New System.Drawing.Size(73, 17)
        Me.rdoAllWeeks.TabIndex = 0
        Me.rdoAllWeeks.TabStop = True
        Me.rdoAllWeeks.Text = "All Weeks"
        Me.rdoAllWeeks.UseVisualStyleBackColor = True
        '
        'btnCalculate
        '
        Me.btnCalculate.BackColor = System.Drawing.SystemColors.Highlight
        Me.btnCalculate.Location = New System.Drawing.Point(326, 45)
        Me.btnCalculate.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCalculate.Name = "btnCalculate"
        Me.btnCalculate.Size = New System.Drawing.Size(65, 25)
        Me.btnCalculate.TabIndex = 0
        Me.btnCalculate.Text = "&Calculate"
        Me.btnCalculate.UseVisualStyleBackColor = False
        '
        'GroupBoxFind
        '
        Me.GroupBoxFind.Controls.Add(Me.lblFind)
        Me.GroupBoxFind.Controls.Add(Me.txbFind)
        Me.GroupBoxFind.Controls.Add(Me.btnFind)
        Me.GroupBoxFind.Location = New System.Drawing.Point(412, 17)
        Me.GroupBoxFind.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxFind.Name = "GroupBoxFind"
        Me.GroupBoxFind.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxFind.Size = New System.Drawing.Size(92, 87)
        Me.GroupBoxFind.TabIndex = 5313
        Me.GroupBoxFind.TabStop = False
        Me.GroupBoxFind.Text = "Find"
        '
        'lblFind
        '
        Me.lblFind.AutoSize = True
        Me.lblFind.Location = New System.Drawing.Point(4, 15)
        Me.lblFind.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFind.Name = "lblFind"
        Me.lblFind.Size = New System.Drawing.Size(87, 13)
        Me.lblFind.TabIndex = 5294
        Me.lblFind.Text = "MRP Reference:"
        '
        'txbFind
        '
        Me.txbFind.Location = New System.Drawing.Point(7, 33)
        Me.txbFind.Margin = New System.Windows.Forms.Padding(2)
        Me.txbFind.Name = "txbFind"
        Me.txbFind.Size = New System.Drawing.Size(83, 20)
        Me.txbFind.TabIndex = 5291
        '
        'btnFind
        '
        Me.btnFind.Location = New System.Drawing.Point(7, 56)
        Me.btnFind.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFind.Name = "btnFind"
        Me.btnFind.Size = New System.Drawing.Size(82, 25)
        Me.btnFind.TabIndex = 5292
        Me.btnFind.Text = "&Find"
        Me.btnFind.UseVisualStyleBackColor = True
        '
        'btnExportToExcel
        '
        Me.btnExportToExcel.Enabled = False
        Me.btnExportToExcel.Image = CType(resources.GetObject("btnExportToExcel.Image"), System.Drawing.Image)
        Me.btnExportToExcel.Location = New System.Drawing.Point(672, 62)
        Me.btnExportToExcel.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExportToExcel.Name = "btnExportToExcel"
        Me.btnExportToExcel.Size = New System.Drawing.Size(38, 40)
        Me.btnExportToExcel.TabIndex = 5295
        Me.btnExportToExcel.UseVisualStyleBackColor = True
        '
        'lblQty
        '
        Me.lblQty.AutoSize = True
        Me.lblQty.Location = New System.Drawing.Point(287, 86)
        Me.lblQty.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblQty.Name = "lblQty"
        Me.lblQty.Size = New System.Drawing.Size(26, 13)
        Me.lblQty.TabIndex = 5312
        Me.lblQty.Text = "Qty:"
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.SystemColors.Control
        Me.btnClear.Image = CType(resources.GetObject("btnClear.Image"), System.Drawing.Image)
        Me.btnClear.Location = New System.Drawing.Point(714, 62)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(2)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(36, 40)
        Me.btnClear.TabIndex = 5296
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'txbQty
        '
        Me.txbQty.Location = New System.Drawing.Point(320, 84)
        Me.txbQty.Margin = New System.Windows.Forms.Padding(2)
        Me.txbQty.Name = "txbQty"
        Me.txbQty.Size = New System.Drawing.Size(76, 20)
        Me.txbQty.TabIndex = 5311
        '
        'lblMRPReference
        '
        Me.lblMRPReference.AutoSize = True
        Me.lblMRPReference.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMRPReference.Location = New System.Drawing.Point(661, 13)
        Me.lblMRPReference.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMRPReference.Name = "lblMRPReference"
        Me.lblMRPReference.Size = New System.Drawing.Size(16, 24)
        Me.lblMRPReference.TabIndex = 5297
        Me.lblMRPReference.Text = "-"
        '
        'lblTMRPReference
        '
        Me.lblTMRPReference.AutoSize = True
        Me.lblTMRPReference.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTMRPReference.Location = New System.Drawing.Point(515, 13)
        Me.lblTMRPReference.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTMRPReference.Name = "lblTMRPReference"
        Me.lblTMRPReference.Size = New System.Drawing.Size(155, 25)
        Me.lblTMRPReference.TabIndex = 5298
        Me.lblTMRPReference.Text = "MRP Reference:"
        '
        'lblRecordsMRP
        '
        Me.lblRecordsMRP.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRecordsMRP.AutoSize = True
        Me.lblRecordsMRP.Location = New System.Drawing.Point(686, 438)
        Me.lblRecordsMRP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsMRP.Name = "lblRecordsMRP"
        Me.lblRecordsMRP.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsMRP.TabIndex = 5300
        Me.lblRecordsMRP.Text = "Records: 0"
        '
        'lblMRP
        '
        Me.lblMRP.AutoSize = True
        Me.lblMRP.Location = New System.Drawing.Point(4, 90)
        Me.lblMRP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblMRP.Name = "lblMRP"
        Me.lblMRP.Size = New System.Drawing.Size(34, 13)
        Me.lblMRP.TabIndex = 5301
        Me.lblMRP.Text = "MRP:"
        '
        'TabPageBOMWIP
        '
        Me.TabPageBOMWIP.Controls.Add(Me.GroupBoxAUBOMWIP)
        Me.TabPageBOMWIP.Controls.Add(Me.GroupBoxPNBOMWIP)
        Me.TabPageBOMWIP.Controls.Add(Me.lblRecordsBOMWIP)
        Me.TabPageBOMWIP.Controls.Add(Me.GridBOMWIP)
        Me.TabPageBOMWIP.Location = New System.Drawing.Point(4, 22)
        Me.TabPageBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageBOMWIP.Name = "TabPageBOMWIP"
        Me.TabPageBOMWIP.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageBOMWIP.Size = New System.Drawing.Size(926, 452)
        Me.TabPageBOMWIP.TabIndex = 1
        Me.TabPageBOMWIP.Text = "BOM WIP"
        Me.TabPageBOMWIP.UseVisualStyleBackColor = True
        '
        'GroupBoxAUBOMWIP
        '
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.Label1)
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.cmbWIPBOMWIP)
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.cmbAUBOMWIP)
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.cmbRevBOMWIP)
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.lblTRevBOMWIP)
        Me.GroupBoxAUBOMWIP.Controls.Add(Me.lblTAUBOMWIP)
        Me.GroupBoxAUBOMWIP.Location = New System.Drawing.Point(8, 105)
        Me.GroupBoxAUBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxAUBOMWIP.Name = "GroupBoxAUBOMWIP"
        Me.GroupBoxAUBOMWIP.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxAUBOMWIP.Size = New System.Drawing.Size(200, 113)
        Me.GroupBoxAUBOMWIP.TabIndex = 1134
        Me.GroupBoxAUBOMWIP.TabStop = False
        Me.GroupBoxAUBOMWIP.Text = "By AU WIP"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 78)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(31, 13)
        Me.Label1.TabIndex = 1133
        Me.Label1.Text = "WIP:"
        '
        'cmbWIPBOMWIP
        '
        Me.cmbWIPBOMWIP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbWIPBOMWIP.FormattingEnabled = True
        Me.cmbWIPBOMWIP.Location = New System.Drawing.Point(62, 76)
        Me.cmbWIPBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbWIPBOMWIP.Name = "cmbWIPBOMWIP"
        Me.cmbWIPBOMWIP.Size = New System.Drawing.Size(92, 21)
        Me.cmbWIPBOMWIP.TabIndex = 1132
        '
        'cmbAUBOMWIP
        '
        Me.cmbAUBOMWIP.FormattingEnabled = True
        Me.cmbAUBOMWIP.Location = New System.Drawing.Point(62, 27)
        Me.cmbAUBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAUBOMWIP.Name = "cmbAUBOMWIP"
        Me.cmbAUBOMWIP.Size = New System.Drawing.Size(92, 21)
        Me.cmbAUBOMWIP.TabIndex = 1127
        '
        'cmbRevBOMWIP
        '
        Me.cmbRevBOMWIP.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRevBOMWIP.FormattingEnabled = True
        Me.cmbRevBOMWIP.Location = New System.Drawing.Point(62, 51)
        Me.cmbRevBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbRevBOMWIP.Name = "cmbRevBOMWIP"
        Me.cmbRevBOMWIP.Size = New System.Drawing.Size(92, 21)
        Me.cmbRevBOMWIP.TabIndex = 1128
        '
        'lblTRevBOMWIP
        '
        Me.lblTRevBOMWIP.AutoSize = True
        Me.lblTRevBOMWIP.Location = New System.Drawing.Point(29, 54)
        Me.lblTRevBOMWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTRevBOMWIP.Name = "lblTRevBOMWIP"
        Me.lblTRevBOMWIP.Size = New System.Drawing.Size(30, 13)
        Me.lblTRevBOMWIP.TabIndex = 1131
        Me.lblTRevBOMWIP.Text = "Rev:"
        '
        'lblTAUBOMWIP
        '
        Me.lblTAUBOMWIP.AutoSize = True
        Me.lblTAUBOMWIP.Location = New System.Drawing.Point(34, 29)
        Me.lblTAUBOMWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTAUBOMWIP.Name = "lblTAUBOMWIP"
        Me.lblTAUBOMWIP.Size = New System.Drawing.Size(25, 13)
        Me.lblTAUBOMWIP.TabIndex = 1130
        Me.lblTAUBOMWIP.Text = "AU:"
        '
        'GroupBoxPNBOMWIP
        '
        Me.GroupBoxPNBOMWIP.Controls.Add(Me.txbBOMWIP)
        Me.GroupBoxPNBOMWIP.Controls.Add(Me.lblTPNBOMWIP)
        Me.GroupBoxPNBOMWIP.Controls.Add(Me.btnFindBOMWIP)
        Me.GroupBoxPNBOMWIP.Location = New System.Drawing.Point(8, 28)
        Me.GroupBoxPNBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNBOMWIP.Name = "GroupBoxPNBOMWIP"
        Me.GroupBoxPNBOMWIP.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNBOMWIP.Size = New System.Drawing.Size(200, 72)
        Me.GroupBoxPNBOMWIP.TabIndex = 1133
        Me.GroupBoxPNBOMWIP.TabStop = False
        Me.GroupBoxPNBOMWIP.Text = "By PN"
        '
        'txbBOMWIP
        '
        Me.txbBOMWIP.Location = New System.Drawing.Point(32, 25)
        Me.txbBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.txbBOMWIP.Name = "txbBOMWIP"
        Me.txbBOMWIP.Size = New System.Drawing.Size(121, 20)
        Me.txbBOMWIP.TabIndex = 1126
        '
        'lblTPNBOMWIP
        '
        Me.lblTPNBOMWIP.AutoSize = True
        Me.lblTPNBOMWIP.Location = New System.Drawing.Point(4, 28)
        Me.lblTPNBOMWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTPNBOMWIP.Name = "lblTPNBOMWIP"
        Me.lblTPNBOMWIP.Size = New System.Drawing.Size(25, 13)
        Me.lblTPNBOMWIP.TabIndex = 1129
        Me.lblTPNBOMWIP.Text = "PN:"
        '
        'btnFindBOMWIP
        '
        Me.btnFindBOMWIP.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindBOMWIP.Image = CType(resources.GetObject("btnFindBOMWIP.Image"), System.Drawing.Image)
        Me.btnFindBOMWIP.Location = New System.Drawing.Point(157, 17)
        Me.btnFindBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindBOMWIP.Name = "btnFindBOMWIP"
        Me.btnFindBOMWIP.Size = New System.Drawing.Size(38, 38)
        Me.btnFindBOMWIP.TabIndex = 1132
        Me.btnFindBOMWIP.Text = "         Find"
        Me.btnFindBOMWIP.UseVisualStyleBackColor = True
        '
        'lblRecordsBOMWIP
        '
        Me.lblRecordsBOMWIP.AutoSize = True
        Me.lblRecordsBOMWIP.Location = New System.Drawing.Point(220, 11)
        Me.lblRecordsBOMWIP.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsBOMWIP.Name = "lblRecordsBOMWIP"
        Me.lblRecordsBOMWIP.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsBOMWIP.TabIndex = 1116
        Me.lblRecordsBOMWIP.Text = "Records: 0"
        '
        'GridBOMWIP
        '
        Me.GridBOMWIP.AllowUserToAddRows = False
        Me.GridBOMWIP.AllowUserToDeleteRows = False
        Me.GridBOMWIP.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridBOMWIP.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridBOMWIP.Location = New System.Drawing.Point(212, 28)
        Me.GridBOMWIP.Margin = New System.Windows.Forms.Padding(2)
        Me.GridBOMWIP.Name = "GridBOMWIP"
        Me.GridBOMWIP.RowTemplate.Height = 24
        Me.GridBOMWIP.Size = New System.Drawing.Size(704, 407)
        Me.GridBOMWIP.TabIndex = 1115
        '
        'TabPageBOMENG
        '
        Me.TabPageBOMENG.Controls.Add(Me.GroupBoxByAUBOMENG)
        Me.TabPageBOMENG.Controls.Add(Me.GroupBoxPNBOMENG)
        Me.TabPageBOMENG.Controls.Add(Me.lblRecordsBOMENG)
        Me.TabPageBOMENG.Controls.Add(Me.GridBOMENG)
        Me.TabPageBOMENG.Location = New System.Drawing.Point(4, 22)
        Me.TabPageBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageBOMENG.Name = "TabPageBOMENG"
        Me.TabPageBOMENG.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageBOMENG.Size = New System.Drawing.Size(926, 452)
        Me.TabPageBOMENG.TabIndex = 2
        Me.TabPageBOMENG.Text = "BOM ENG"
        Me.TabPageBOMENG.UseVisualStyleBackColor = True
        '
        'GroupBoxByAUBOMENG
        '
        Me.GroupBoxByAUBOMENG.Controls.Add(Me.cmbAUBOMENG)
        Me.GroupBoxByAUBOMENG.Controls.Add(Me.cmbRevBOMENG)
        Me.GroupBoxByAUBOMENG.Controls.Add(Me.lblTAUBOMENG)
        Me.GroupBoxByAUBOMENG.Controls.Add(Me.lblTRevBOMENG)
        Me.GroupBoxByAUBOMENG.Location = New System.Drawing.Point(4, 105)
        Me.GroupBoxByAUBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxByAUBOMENG.Name = "GroupBoxByAUBOMENG"
        Me.GroupBoxByAUBOMENG.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxByAUBOMENG.Size = New System.Drawing.Size(203, 76)
        Me.GroupBoxByAUBOMENG.TabIndex = 1129
        Me.GroupBoxByAUBOMENG.TabStop = False
        Me.GroupBoxByAUBOMENG.Text = "By AU"
        '
        'cmbAUBOMENG
        '
        Me.cmbAUBOMENG.FormattingEnabled = True
        Me.cmbAUBOMENG.Location = New System.Drawing.Point(69, 17)
        Me.cmbAUBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAUBOMENG.Name = "cmbAUBOMENG"
        Me.cmbAUBOMENG.Size = New System.Drawing.Size(92, 21)
        Me.cmbAUBOMENG.TabIndex = 1120
        '
        'cmbRevBOMENG
        '
        Me.cmbRevBOMENG.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRevBOMENG.FormattingEnabled = True
        Me.cmbRevBOMENG.Location = New System.Drawing.Point(69, 41)
        Me.cmbRevBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbRevBOMENG.Name = "cmbRevBOMENG"
        Me.cmbRevBOMENG.Size = New System.Drawing.Size(92, 21)
        Me.cmbRevBOMENG.TabIndex = 1121
        '
        'lblTAUBOMENG
        '
        Me.lblTAUBOMENG.AutoSize = True
        Me.lblTAUBOMENG.Location = New System.Drawing.Point(41, 20)
        Me.lblTAUBOMENG.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTAUBOMENG.Name = "lblTAUBOMENG"
        Me.lblTAUBOMENG.Size = New System.Drawing.Size(25, 13)
        Me.lblTAUBOMENG.TabIndex = 1123
        Me.lblTAUBOMENG.Text = "AU:"
        '
        'lblTRevBOMENG
        '
        Me.lblTRevBOMENG.AutoSize = True
        Me.lblTRevBOMENG.Location = New System.Drawing.Point(37, 44)
        Me.lblTRevBOMENG.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTRevBOMENG.Name = "lblTRevBOMENG"
        Me.lblTRevBOMENG.Size = New System.Drawing.Size(30, 13)
        Me.lblTRevBOMENG.TabIndex = 1124
        Me.lblTRevBOMENG.Text = "Rev:"
        '
        'GroupBoxPNBOMENG
        '
        Me.GroupBoxPNBOMENG.Controls.Add(Me.txbPNBOMENG)
        Me.GroupBoxPNBOMENG.Controls.Add(Me.lblTPNBOMENG)
        Me.GroupBoxPNBOMENG.Controls.Add(Me.btnFindBOMENG)
        Me.GroupBoxPNBOMENG.Location = New System.Drawing.Point(4, 28)
        Me.GroupBoxPNBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNBOMENG.Name = "GroupBoxPNBOMENG"
        Me.GroupBoxPNBOMENG.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNBOMENG.Size = New System.Drawing.Size(203, 64)
        Me.GroupBoxPNBOMENG.TabIndex = 1128
        Me.GroupBoxPNBOMENG.TabStop = False
        Me.GroupBoxPNBOMENG.Text = "By PN"
        '
        'txbPNBOMENG
        '
        Me.txbPNBOMENG.Location = New System.Drawing.Point(32, 25)
        Me.txbPNBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.txbPNBOMENG.Name = "txbPNBOMENG"
        Me.txbPNBOMENG.Size = New System.Drawing.Size(106, 20)
        Me.txbPNBOMENG.TabIndex = 1119
        '
        'lblTPNBOMENG
        '
        Me.lblTPNBOMENG.AutoSize = True
        Me.lblTPNBOMENG.Location = New System.Drawing.Point(4, 28)
        Me.lblTPNBOMENG.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTPNBOMENG.Name = "lblTPNBOMENG"
        Me.lblTPNBOMENG.Size = New System.Drawing.Size(25, 13)
        Me.lblTPNBOMENG.TabIndex = 1122
        Me.lblTPNBOMENG.Text = "PN:"
        '
        'btnFindBOMENG
        '
        Me.btnFindBOMENG.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindBOMENG.Image = CType(resources.GetObject("btnFindBOMENG.Image"), System.Drawing.Image)
        Me.btnFindBOMENG.Location = New System.Drawing.Point(142, 17)
        Me.btnFindBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindBOMENG.Name = "btnFindBOMENG"
        Me.btnFindBOMENG.Size = New System.Drawing.Size(38, 38)
        Me.btnFindBOMENG.TabIndex = 1125
        Me.btnFindBOMENG.Text = "         Find"
        Me.btnFindBOMENG.UseVisualStyleBackColor = True
        '
        'lblRecordsBOMENG
        '
        Me.lblRecordsBOMENG.AutoSize = True
        Me.lblRecordsBOMENG.Location = New System.Drawing.Point(220, 11)
        Me.lblRecordsBOMENG.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsBOMENG.Name = "lblRecordsBOMENG"
        Me.lblRecordsBOMENG.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsBOMENG.TabIndex = 1118
        Me.lblRecordsBOMENG.Text = "Records: 0"
        '
        'GridBOMENG
        '
        Me.GridBOMENG.AllowUserToAddRows = False
        Me.GridBOMENG.AllowUserToDeleteRows = False
        Me.GridBOMENG.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridBOMENG.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridBOMENG.Location = New System.Drawing.Point(212, 28)
        Me.GridBOMENG.Margin = New System.Windows.Forms.Padding(2)
        Me.GridBOMENG.Name = "GridBOMENG"
        Me.GridBOMENG.RowTemplate.Height = 24
        Me.GridBOMENG.Size = New System.Drawing.Size(704, 407)
        Me.GridBOMENG.TabIndex = 1117
        '
        'TabPageMyTable
        '
        Me.TabPageMyTable.Controls.Add(Me.GroupBoxPNMyTable)
        Me.TabPageMyTable.Controls.Add(Me.lblRecordsMyTable)
        Me.TabPageMyTable.Controls.Add(Me.GridMyTable)
        Me.TabPageMyTable.Location = New System.Drawing.Point(4, 22)
        Me.TabPageMyTable.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageMyTable.Name = "TabPageMyTable"
        Me.TabPageMyTable.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageMyTable.Size = New System.Drawing.Size(926, 452)
        Me.TabPageMyTable.TabIndex = 3
        Me.TabPageMyTable.Text = "Search in my table"
        Me.TabPageMyTable.UseVisualStyleBackColor = True
        '
        'GroupBoxPNMyTable
        '
        Me.GroupBoxPNMyTable.Controls.Add(Me.cmbPNMyTable)
        Me.GroupBoxPNMyTable.Location = New System.Drawing.Point(8, 32)
        Me.GroupBoxPNMyTable.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNMyTable.Name = "GroupBoxPNMyTable"
        Me.GroupBoxPNMyTable.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxPNMyTable.Size = New System.Drawing.Size(203, 63)
        Me.GroupBoxPNMyTable.TabIndex = 1132
        Me.GroupBoxPNMyTable.TabStop = False
        Me.GroupBoxPNMyTable.Text = "By PN"
        '
        'cmbPNMyTable
        '
        Me.cmbPNMyTable.FormattingEnabled = True
        Me.cmbPNMyTable.Location = New System.Drawing.Point(4, 26)
        Me.cmbPNMyTable.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbPNMyTable.Name = "cmbPNMyTable"
        Me.cmbPNMyTable.Size = New System.Drawing.Size(195, 21)
        Me.cmbPNMyTable.TabIndex = 1133
        '
        'lblRecordsMyTable
        '
        Me.lblRecordsMyTable.AutoSize = True
        Me.lblRecordsMyTable.Location = New System.Drawing.Point(224, 15)
        Me.lblRecordsMyTable.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsMyTable.Name = "lblRecordsMyTable"
        Me.lblRecordsMyTable.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsMyTable.TabIndex = 1131
        Me.lblRecordsMyTable.Text = "Records: 0"
        '
        'GridMyTable
        '
        Me.GridMyTable.AllowUserToAddRows = False
        Me.GridMyTable.AllowUserToDeleteRows = False
        Me.GridMyTable.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridMyTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridMyTable.Location = New System.Drawing.Point(216, 32)
        Me.GridMyTable.Margin = New System.Windows.Forms.Padding(2)
        Me.GridMyTable.Name = "GridMyTable"
        Me.GridMyTable.RowTemplate.Height = 24
        Me.GridMyTable.Size = New System.Drawing.Size(704, 407)
        Me.GridMyTable.TabIndex = 1130
        '
        'TabPageSalesOrder
        '
        Me.TabPageSalesOrder.Controls.Add(Me.GroupBoxSalesOrderControl)
        Me.TabPageSalesOrder.Controls.Add(Me.lblRecordsGridSalesOrder)
        Me.TabPageSalesOrder.Controls.Add(Me.GridAUSalesOrderFind)
        Me.TabPageSalesOrder.Location = New System.Drawing.Point(4, 22)
        Me.TabPageSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageSalesOrder.Name = "TabPageSalesOrder"
        Me.TabPageSalesOrder.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageSalesOrder.Size = New System.Drawing.Size(926, 452)
        Me.TabPageSalesOrder.TabIndex = 4
        Me.TabPageSalesOrder.Text = "Sales Order"
        Me.TabPageSalesOrder.UseVisualStyleBackColor = True
        '
        'GroupBoxSalesOrderControl
        '
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.GroupBoxSalesOrderStatus)
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.btnFindSalesOrder)
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.lblTAUSalesOrder)
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.lblTrevSalesOrder)
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.txbAUSalesOrder)
        Me.GroupBoxSalesOrderControl.Controls.Add(Me.cmbRevSalesOrder)
        Me.GroupBoxSalesOrderControl.Location = New System.Drawing.Point(8, 32)
        Me.GroupBoxSalesOrderControl.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSalesOrderControl.Name = "GroupBoxSalesOrderControl"
        Me.GroupBoxSalesOrderControl.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSalesOrderControl.Size = New System.Drawing.Size(203, 164)
        Me.GroupBoxSalesOrderControl.TabIndex = 1135
        Me.GroupBoxSalesOrderControl.TabStop = False
        Me.GroupBoxSalesOrderControl.Text = "By AU"
        '
        'GroupBoxSalesOrderStatus
        '
        Me.GroupBoxSalesOrderStatus.Controls.Add(Me.rdoAllSalesOrderByAU)
        Me.GroupBoxSalesOrderStatus.Controls.Add(Me.rdoOpenSalesOrderByAU)
        Me.GroupBoxSalesOrderStatus.Controls.Add(Me.rdoCancelSalesOrderByAU)
        Me.GroupBoxSalesOrderStatus.Controls.Add(Me.rdoCloseSalesOrderByAU)
        Me.GroupBoxSalesOrderStatus.Location = New System.Drawing.Point(35, 64)
        Me.GroupBoxSalesOrderStatus.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSalesOrderStatus.Name = "GroupBoxSalesOrderStatus"
        Me.GroupBoxSalesOrderStatus.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxSalesOrderStatus.Size = New System.Drawing.Size(61, 80)
        Me.GroupBoxSalesOrderStatus.TabIndex = 1142
        Me.GroupBoxSalesOrderStatus.TabStop = False
        '
        'rdoAllSalesOrderByAU
        '
        Me.rdoAllSalesOrderByAU.AutoSize = True
        Me.rdoAllSalesOrderByAU.Location = New System.Drawing.Point(4, 60)
        Me.rdoAllSalesOrderByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoAllSalesOrderByAU.Name = "rdoAllSalesOrderByAU"
        Me.rdoAllSalesOrderByAU.Size = New System.Drawing.Size(36, 17)
        Me.rdoAllSalesOrderByAU.TabIndex = 1141
        Me.rdoAllSalesOrderByAU.Text = "All"
        Me.rdoAllSalesOrderByAU.UseVisualStyleBackColor = True
        '
        'rdoOpenSalesOrderByAU
        '
        Me.rdoOpenSalesOrderByAU.AutoSize = True
        Me.rdoOpenSalesOrderByAU.Checked = True
        Me.rdoOpenSalesOrderByAU.Location = New System.Drawing.Point(4, 9)
        Me.rdoOpenSalesOrderByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoOpenSalesOrderByAU.Name = "rdoOpenSalesOrderByAU"
        Me.rdoOpenSalesOrderByAU.Size = New System.Drawing.Size(51, 17)
        Me.rdoOpenSalesOrderByAU.TabIndex = 1138
        Me.rdoOpenSalesOrderByAU.TabStop = True
        Me.rdoOpenSalesOrderByAU.Text = "Open"
        Me.rdoOpenSalesOrderByAU.UseVisualStyleBackColor = True
        '
        'rdoCancelSalesOrderByAU
        '
        Me.rdoCancelSalesOrderByAU.AutoSize = True
        Me.rdoCancelSalesOrderByAU.Location = New System.Drawing.Point(4, 43)
        Me.rdoCancelSalesOrderByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoCancelSalesOrderByAU.Name = "rdoCancelSalesOrderByAU"
        Me.rdoCancelSalesOrderByAU.Size = New System.Drawing.Size(58, 17)
        Me.rdoCancelSalesOrderByAU.TabIndex = 1140
        Me.rdoCancelSalesOrderByAU.Text = "Cancel"
        Me.rdoCancelSalesOrderByAU.UseVisualStyleBackColor = True
        '
        'rdoCloseSalesOrderByAU
        '
        Me.rdoCloseSalesOrderByAU.AutoSize = True
        Me.rdoCloseSalesOrderByAU.Location = New System.Drawing.Point(4, 26)
        Me.rdoCloseSalesOrderByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoCloseSalesOrderByAU.Name = "rdoCloseSalesOrderByAU"
        Me.rdoCloseSalesOrderByAU.Size = New System.Drawing.Size(51, 17)
        Me.rdoCloseSalesOrderByAU.TabIndex = 1139
        Me.rdoCloseSalesOrderByAU.Text = "Close"
        Me.rdoCloseSalesOrderByAU.UseVisualStyleBackColor = True
        '
        'btnFindSalesOrder
        '
        Me.btnFindSalesOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindSalesOrder.Image = CType(resources.GetObject("btnFindSalesOrder.Image"), System.Drawing.Image)
        Me.btnFindSalesOrder.Location = New System.Drawing.Point(160, 18)
        Me.btnFindSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindSalesOrder.Name = "btnFindSalesOrder"
        Me.btnFindSalesOrder.Size = New System.Drawing.Size(38, 38)
        Me.btnFindSalesOrder.TabIndex = 1137
        Me.btnFindSalesOrder.Text = "         Find"
        Me.btnFindSalesOrder.UseVisualStyleBackColor = True
        '
        'lblTAUSalesOrder
        '
        Me.lblTAUSalesOrder.AutoSize = True
        Me.lblTAUSalesOrder.Location = New System.Drawing.Point(8, 20)
        Me.lblTAUSalesOrder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTAUSalesOrder.Name = "lblTAUSalesOrder"
        Me.lblTAUSalesOrder.Size = New System.Drawing.Size(25, 13)
        Me.lblTAUSalesOrder.TabIndex = 1136
        Me.lblTAUSalesOrder.Text = "AU:"
        '
        'lblTrevSalesOrder
        '
        Me.lblTrevSalesOrder.AutoSize = True
        Me.lblTrevSalesOrder.Location = New System.Drawing.Point(3, 42)
        Me.lblTrevSalesOrder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTrevSalesOrder.Name = "lblTrevSalesOrder"
        Me.lblTrevSalesOrder.Size = New System.Drawing.Size(30, 13)
        Me.lblTrevSalesOrder.TabIndex = 1135
        Me.lblTrevSalesOrder.Text = "Rev:"
        '
        'txbAUSalesOrder
        '
        Me.txbAUSalesOrder.Location = New System.Drawing.Point(35, 17)
        Me.txbAUSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.txbAUSalesOrder.Name = "txbAUSalesOrder"
        Me.txbAUSalesOrder.Size = New System.Drawing.Size(122, 20)
        Me.txbAUSalesOrder.TabIndex = 1134
        '
        'cmbRevSalesOrder
        '
        Me.cmbRevSalesOrder.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRevSalesOrder.FormattingEnabled = True
        Me.cmbRevSalesOrder.Location = New System.Drawing.Point(35, 40)
        Me.cmbRevSalesOrder.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbRevSalesOrder.Name = "cmbRevSalesOrder"
        Me.cmbRevSalesOrder.Size = New System.Drawing.Size(122, 21)
        Me.cmbRevSalesOrder.TabIndex = 1133
        '
        'lblRecordsGridSalesOrder
        '
        Me.lblRecordsGridSalesOrder.AutoSize = True
        Me.lblRecordsGridSalesOrder.Location = New System.Drawing.Point(224, 15)
        Me.lblRecordsGridSalesOrder.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsGridSalesOrder.Name = "lblRecordsGridSalesOrder"
        Me.lblRecordsGridSalesOrder.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsGridSalesOrder.TabIndex = 1134
        Me.lblRecordsGridSalesOrder.Text = "Records: 0"
        '
        'GridAUSalesOrderFind
        '
        Me.GridAUSalesOrderFind.AllowUserToAddRows = False
        Me.GridAUSalesOrderFind.AllowUserToDeleteRows = False
        Me.GridAUSalesOrderFind.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridAUSalesOrderFind.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridAUSalesOrderFind.Location = New System.Drawing.Point(216, 32)
        Me.GridAUSalesOrderFind.Margin = New System.Windows.Forms.Padding(2)
        Me.GridAUSalesOrderFind.Name = "GridAUSalesOrderFind"
        Me.GridAUSalesOrderFind.RowTemplate.Height = 24
        Me.GridAUSalesOrderFind.Size = New System.Drawing.Size(704, 407)
        Me.GridAUSalesOrderFind.TabIndex = 1133
        '
        'TabPageWIPByAU
        '
        Me.TabPageWIPByAU.Controls.Add(Me.GroupBoxWIPByAU)
        Me.TabPageWIPByAU.Controls.Add(Me.lblRecordsWipByAU)
        Me.TabPageWIPByAU.Controls.Add(Me.GridWipByAU)
        Me.TabPageWIPByAU.Location = New System.Drawing.Point(4, 22)
        Me.TabPageWIPByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.TabPageWIPByAU.Name = "TabPageWIPByAU"
        Me.TabPageWIPByAU.Padding = New System.Windows.Forms.Padding(2)
        Me.TabPageWIPByAU.Size = New System.Drawing.Size(926, 452)
        Me.TabPageWIPByAU.TabIndex = 5
        Me.TabPageWIPByAU.Text = "WIP"
        Me.TabPageWIPByAU.UseVisualStyleBackColor = True
        '
        'GroupBoxWIPByAU
        '
        Me.GroupBoxWIPByAU.Controls.Add(Me.GroupBoxStatusWIPByAU)
        Me.GroupBoxWIPByAU.Controls.Add(Me.btnFindWipByAU)
        Me.GroupBoxWIPByAU.Controls.Add(Me.lblTAUWipByAU)
        Me.GroupBoxWIPByAU.Controls.Add(Me.lblTRevWipByAU)
        Me.GroupBoxWIPByAU.Controls.Add(Me.txbAUWipByAU)
        Me.GroupBoxWIPByAU.Controls.Add(Me.cmbRevWipByAU)
        Me.GroupBoxWIPByAU.Location = New System.Drawing.Point(8, 32)
        Me.GroupBoxWIPByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxWIPByAU.Name = "GroupBoxWIPByAU"
        Me.GroupBoxWIPByAU.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxWIPByAU.Size = New System.Drawing.Size(203, 164)
        Me.GroupBoxWIPByAU.TabIndex = 1138
        Me.GroupBoxWIPByAU.TabStop = False
        Me.GroupBoxWIPByAU.Text = "By AU"
        '
        'GroupBoxStatusWIPByAU
        '
        Me.GroupBoxStatusWIPByAU.Controls.Add(Me.rdoAllWipByAU)
        Me.GroupBoxStatusWIPByAU.Controls.Add(Me.rdoOpenWipByAU)
        Me.GroupBoxStatusWIPByAU.Controls.Add(Me.rdoCancelWipByAU)
        Me.GroupBoxStatusWIPByAU.Controls.Add(Me.rdoCloseWipByAU)
        Me.GroupBoxStatusWIPByAU.Location = New System.Drawing.Point(35, 64)
        Me.GroupBoxStatusWIPByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxStatusWIPByAU.Name = "GroupBoxStatusWIPByAU"
        Me.GroupBoxStatusWIPByAU.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxStatusWIPByAU.Size = New System.Drawing.Size(61, 80)
        Me.GroupBoxStatusWIPByAU.TabIndex = 1142
        Me.GroupBoxStatusWIPByAU.TabStop = False
        '
        'rdoAllWipByAU
        '
        Me.rdoAllWipByAU.AutoSize = True
        Me.rdoAllWipByAU.Location = New System.Drawing.Point(4, 60)
        Me.rdoAllWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoAllWipByAU.Name = "rdoAllWipByAU"
        Me.rdoAllWipByAU.Size = New System.Drawing.Size(36, 17)
        Me.rdoAllWipByAU.TabIndex = 1141
        Me.rdoAllWipByAU.Text = "All"
        Me.rdoAllWipByAU.UseVisualStyleBackColor = True
        '
        'rdoOpenWipByAU
        '
        Me.rdoOpenWipByAU.AutoSize = True
        Me.rdoOpenWipByAU.Checked = True
        Me.rdoOpenWipByAU.Location = New System.Drawing.Point(4, 9)
        Me.rdoOpenWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoOpenWipByAU.Name = "rdoOpenWipByAU"
        Me.rdoOpenWipByAU.Size = New System.Drawing.Size(51, 17)
        Me.rdoOpenWipByAU.TabIndex = 1138
        Me.rdoOpenWipByAU.TabStop = True
        Me.rdoOpenWipByAU.Text = "Open"
        Me.rdoOpenWipByAU.UseVisualStyleBackColor = True
        '
        'rdoCancelWipByAU
        '
        Me.rdoCancelWipByAU.AutoSize = True
        Me.rdoCancelWipByAU.Location = New System.Drawing.Point(4, 43)
        Me.rdoCancelWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoCancelWipByAU.Name = "rdoCancelWipByAU"
        Me.rdoCancelWipByAU.Size = New System.Drawing.Size(58, 17)
        Me.rdoCancelWipByAU.TabIndex = 1140
        Me.rdoCancelWipByAU.Text = "Cancel"
        Me.rdoCancelWipByAU.UseVisualStyleBackColor = True
        '
        'rdoCloseWipByAU
        '
        Me.rdoCloseWipByAU.AutoSize = True
        Me.rdoCloseWipByAU.Location = New System.Drawing.Point(4, 26)
        Me.rdoCloseWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.rdoCloseWipByAU.Name = "rdoCloseWipByAU"
        Me.rdoCloseWipByAU.Size = New System.Drawing.Size(51, 17)
        Me.rdoCloseWipByAU.TabIndex = 1139
        Me.rdoCloseWipByAU.Text = "Close"
        Me.rdoCloseWipByAU.UseVisualStyleBackColor = True
        '
        'btnFindWipByAU
        '
        Me.btnFindWipByAU.Font = New System.Drawing.Font("Microsoft Sans Serif", 4.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFindWipByAU.Image = CType(resources.GetObject("btnFindWipByAU.Image"), System.Drawing.Image)
        Me.btnFindWipByAU.Location = New System.Drawing.Point(160, 18)
        Me.btnFindWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.btnFindWipByAU.Name = "btnFindWipByAU"
        Me.btnFindWipByAU.Size = New System.Drawing.Size(38, 38)
        Me.btnFindWipByAU.TabIndex = 1137
        Me.btnFindWipByAU.Text = "         Find"
        Me.btnFindWipByAU.UseVisualStyleBackColor = True
        '
        'lblTAUWipByAU
        '
        Me.lblTAUWipByAU.AutoSize = True
        Me.lblTAUWipByAU.Location = New System.Drawing.Point(8, 20)
        Me.lblTAUWipByAU.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTAUWipByAU.Name = "lblTAUWipByAU"
        Me.lblTAUWipByAU.Size = New System.Drawing.Size(25, 13)
        Me.lblTAUWipByAU.TabIndex = 1136
        Me.lblTAUWipByAU.Text = "AU:"
        '
        'lblTRevWipByAU
        '
        Me.lblTRevWipByAU.AutoSize = True
        Me.lblTRevWipByAU.Location = New System.Drawing.Point(3, 42)
        Me.lblTRevWipByAU.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTRevWipByAU.Name = "lblTRevWipByAU"
        Me.lblTRevWipByAU.Size = New System.Drawing.Size(30, 13)
        Me.lblTRevWipByAU.TabIndex = 1135
        Me.lblTRevWipByAU.Text = "Rev:"
        '
        'txbAUWipByAU
        '
        Me.txbAUWipByAU.Location = New System.Drawing.Point(35, 17)
        Me.txbAUWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.txbAUWipByAU.Name = "txbAUWipByAU"
        Me.txbAUWipByAU.Size = New System.Drawing.Size(122, 20)
        Me.txbAUWipByAU.TabIndex = 1134
        '
        'cmbRevWipByAU
        '
        Me.cmbRevWipByAU.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbRevWipByAU.FormattingEnabled = True
        Me.cmbRevWipByAU.Location = New System.Drawing.Point(35, 40)
        Me.cmbRevWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbRevWipByAU.Name = "cmbRevWipByAU"
        Me.cmbRevWipByAU.Size = New System.Drawing.Size(122, 21)
        Me.cmbRevWipByAU.TabIndex = 1133
        '
        'lblRecordsWipByAU
        '
        Me.lblRecordsWipByAU.AutoSize = True
        Me.lblRecordsWipByAU.Location = New System.Drawing.Point(224, 15)
        Me.lblRecordsWipByAU.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblRecordsWipByAU.Name = "lblRecordsWipByAU"
        Me.lblRecordsWipByAU.Size = New System.Drawing.Size(59, 13)
        Me.lblRecordsWipByAU.TabIndex = 1137
        Me.lblRecordsWipByAU.Text = "Records: 0"
        '
        'GridWipByAU
        '
        Me.GridWipByAU.AllowUserToAddRows = False
        Me.GridWipByAU.AllowUserToDeleteRows = False
        Me.GridWipByAU.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GridWipByAU.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridWipByAU.Location = New System.Drawing.Point(216, 32)
        Me.GridWipByAU.Margin = New System.Windows.Forms.Padding(2)
        Me.GridWipByAU.Name = "GridWipByAU"
        Me.GridWipByAU.RowTemplate.Height = 24
        Me.GridWipByAU.Size = New System.Drawing.Size(704, 407)
        Me.GridWipByAU.TabIndex = 1136
        '
        'GroupBoxUserMRP
        '
        Me.GroupBoxUserMRP.Controls.Add(Me.btnCancelLoginEng)
        Me.GroupBoxUserMRP.Controls.Add(Me.btnLoginMRP)
        Me.GroupBoxUserMRP.Controls.Add(Me.txbUserMRP)
        Me.GroupBoxUserMRP.Controls.Add(Me.lblTEngPassword)
        Me.GroupBoxUserMRP.Controls.Add(Me.txbUserMRPPassword)
        Me.GroupBoxUserMRP.Controls.Add(Me.lblTEngUser)
        Me.GroupBoxUserMRP.Location = New System.Drawing.Point(377, -1)
        Me.GroupBoxUserMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.GroupBoxUserMRP.Name = "GroupBoxUserMRP"
        Me.GroupBoxUserMRP.Padding = New System.Windows.Forms.Padding(2)
        Me.GroupBoxUserMRP.Size = New System.Drawing.Size(192, 107)
        Me.GroupBoxUserMRP.TabIndex = 5331
        Me.GroupBoxUserMRP.TabStop = False
        Me.GroupBoxUserMRP.Text = "Datos del Comprador"
        '
        'btnCancelLoginEng
        '
        Me.btnCancelLoginEng.Image = CType(resources.GetObject("btnCancelLoginEng.Image"), System.Drawing.Image)
        Me.btnCancelLoginEng.Location = New System.Drawing.Point(4, 74)
        Me.btnCancelLoginEng.Margin = New System.Windows.Forms.Padding(2)
        Me.btnCancelLoginEng.Name = "btnCancelLoginEng"
        Me.btnCancelLoginEng.Size = New System.Drawing.Size(38, 31)
        Me.btnCancelLoginEng.TabIndex = 12
        Me.btnCancelLoginEng.UseVisualStyleBackColor = True
        '
        'btnLoginMRP
        '
        Me.btnLoginMRP.Location = New System.Drawing.Point(126, 74)
        Me.btnLoginMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoginMRP.Name = "btnLoginMRP"
        Me.btnLoginMRP.Size = New System.Drawing.Size(56, 28)
        Me.btnLoginMRP.TabIndex = 11
        Me.btnLoginMRP.Text = "Entrar"
        Me.btnLoginMRP.UseVisualStyleBackColor = True
        '
        'txbUserMRP
        '
        Me.txbUserMRP.Location = New System.Drawing.Point(94, 28)
        Me.txbUserMRP.Margin = New System.Windows.Forms.Padding(2)
        Me.txbUserMRP.Name = "txbUserMRP"
        Me.txbUserMRP.Size = New System.Drawing.Size(76, 20)
        Me.txbUserMRP.TabIndex = 9
        '
        'lblTEngPassword
        '
        Me.lblTEngPassword.AutoSize = True
        Me.lblTEngPassword.Location = New System.Drawing.Point(26, 54)
        Me.lblTEngPassword.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTEngPassword.Name = "lblTEngPassword"
        Me.lblTEngPassword.Size = New System.Drawing.Size(64, 13)
        Me.lblTEngPassword.TabIndex = 5
        Me.lblTEngPassword.Text = "Contraseña:"
        '
        'txbUserMRPPassword
        '
        Me.txbUserMRPPassword.Location = New System.Drawing.Point(94, 51)
        Me.txbUserMRPPassword.Margin = New System.Windows.Forms.Padding(2)
        Me.txbUserMRPPassword.Name = "txbUserMRPPassword"
        Me.txbUserMRPPassword.Size = New System.Drawing.Size(76, 20)
        Me.txbUserMRPPassword.TabIndex = 10
        Me.txbUserMRPPassword.UseSystemPasswordChar = True
        '
        'lblTEngUser
        '
        Me.lblTEngUser.AutoSize = True
        Me.lblTEngUser.Location = New System.Drawing.Point(44, 31)
        Me.lblTEngUser.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblTEngUser.Name = "lblTEngUser"
        Me.lblTEngUser.Size = New System.Drawing.Size(46, 13)
        Me.lblTEngUser.TabIndex = 4
        Me.lblTEngUser.Text = "Usuario:"
        '
        'cmbPONoAprovadas
        '
        Me.cmbPONoAprovadas.FormattingEnabled = True
        Me.cmbPONoAprovadas.Location = New System.Drawing.Point(86, 486)
        Me.cmbPONoAprovadas.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbPONoAprovadas.Name = "cmbPONoAprovadas"
        Me.cmbPONoAprovadas.Size = New System.Drawing.Size(92, 21)
        Me.cmbPONoAprovadas.TabIndex = 5330
        Me.cmbPONoAprovadas.Visible = False
        '
        'txbUser
        '
        Me.txbUser.Location = New System.Drawing.Point(5, 486)
        Me.txbUser.Margin = New System.Windows.Forms.Padding(2)
        Me.txbUser.Name = "txbUser"
        Me.txbUser.Size = New System.Drawing.Size(76, 20)
        Me.txbUser.TabIndex = 5329
        Me.txbUser.Visible = False
        '
        'frmMRP
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(946, 506)
        Me.Controls.Add(Me.GroupBoxUserMRP)
        Me.Controls.Add(Me.TabControlMRPGlobal)
        Me.Controls.Add(Me.cmbPONoAprovadas)
        Me.Controls.Add(Me.txbUser)
        Me.Name = "frmMRP"
        Me.Text = "MRP"
        Me.TabControlMRPGlobal.ResumeLayout(False)
        Me.TabPageMRP.ResumeLayout(False)
        Me.TabPageMRP.PerformLayout()
        Me.GroupBoxBudgetInformation.ResumeLayout(False)
        Me.GroupBoxBudgetInformation.PerformLayout()
        CType(Me.GridPerVendor, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridPerWeek, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxPurchasingOrderHistory.ResumeLayout(False)
        Me.GroupBoxPurchasingOrderHistory.PerformLayout()
        CType(Me.GridPurchasingOrderItemsHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupApproved.ResumeLayout(False)
        Me.GroupApproved.PerformLayout()
        Me.GroupBoxSaved.ResumeLayout(False)
        Me.GroupBoxSaved.PerformLayout()
        CType(Me.GridMRP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxOption.ResumeLayout(False)
        Me.GroupBoxOption.PerformLayout()
        Me.GroupWipSalesOrder.ResumeLayout(False)
        Me.GroupWipSalesOrder.PerformLayout()
        CType(Me.GridSalesOrder, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.GridWIP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxFind.ResumeLayout(False)
        Me.GroupBoxFind.PerformLayout()
        Me.TabPageBOMWIP.ResumeLayout(False)
        Me.TabPageBOMWIP.PerformLayout()
        Me.GroupBoxAUBOMWIP.ResumeLayout(False)
        Me.GroupBoxAUBOMWIP.PerformLayout()
        Me.GroupBoxPNBOMWIP.ResumeLayout(False)
        Me.GroupBoxPNBOMWIP.PerformLayout()
        CType(Me.GridBOMWIP, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageBOMENG.ResumeLayout(False)
        Me.TabPageBOMENG.PerformLayout()
        Me.GroupBoxByAUBOMENG.ResumeLayout(False)
        Me.GroupBoxByAUBOMENG.PerformLayout()
        Me.GroupBoxPNBOMENG.ResumeLayout(False)
        Me.GroupBoxPNBOMENG.PerformLayout()
        CType(Me.GridBOMENG, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageMyTable.ResumeLayout(False)
        Me.TabPageMyTable.PerformLayout()
        Me.GroupBoxPNMyTable.ResumeLayout(False)
        CType(Me.GridMyTable, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageSalesOrder.ResumeLayout(False)
        Me.TabPageSalesOrder.PerformLayout()
        Me.GroupBoxSalesOrderControl.ResumeLayout(False)
        Me.GroupBoxSalesOrderControl.PerformLayout()
        Me.GroupBoxSalesOrderStatus.ResumeLayout(False)
        Me.GroupBoxSalesOrderStatus.PerformLayout()
        CType(Me.GridAUSalesOrderFind, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPageWIPByAU.ResumeLayout(False)
        Me.TabPageWIPByAU.PerformLayout()
        Me.GroupBoxWIPByAU.ResumeLayout(False)
        Me.GroupBoxWIPByAU.PerformLayout()
        Me.GroupBoxStatusWIPByAU.ResumeLayout(False)
        Me.GroupBoxStatusWIPByAU.PerformLayout()
        CType(Me.GridWipByAU, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBoxUserMRP.ResumeLayout(False)
        Me.GroupBoxUserMRP.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents TabControlMRPGlobal As TabControl
    Friend WithEvents TabPageMRP As TabPage
    Friend WithEvents GroupWipSalesOrder As GroupBox
    Friend WithEvents lblRecordsSalesOrder As Label
    Friend WithEvents lblRecordsWip As Label
    Friend WithEvents lblTSalesOrder As Label
    Friend WithEvents lblTWIP As Label
    Friend WithEvents btnRefreshSalesOrders As Button
    Friend WithEvents btnCloseAddIems As Button
    Friend WithEvents GridSalesOrder As DataGridView
    Friend WithEvents GridWIP As DataGridView
    Friend WithEvents GroupBoxBudgetInformation As GroupBox
    Friend WithEvents lblTotalTotal2 As Label
    Friend WithEvents lblRecordsPerVendor As Label
    Friend WithEvents lblRecordsPerWeek As Label
    Friend WithEvents lblTPerVendor As Label
    Friend WithEvents lblTPerWeek As Label
    Friend WithEvents btnCloseBudgetInformation As Button
    Friend WithEvents GridPerVendor As DataGridView
    Friend WithEvents GridPerWeek As DataGridView
    Friend WithEvents GroupBoxPurchasingOrderHistory As GroupBox
    Friend WithEvents btnRefreshPurchasingOrderItemsHistory As Button
    Friend WithEvents btnClosePurchasingOrderItemsHistory As Button
    Friend WithEvents lblRecordsPurchasingOrderItemsHistory As Label
    Friend WithEvents lblTItems As Label
    Friend WithEvents GridPurchasingOrderItemsHistory As DataGridView
    Friend WithEvents lblTotal As LinkLabel
    Friend WithEvents txbExchangeRate As TextBox
    Friend WithEvents lblExchangeRate As Label
    Friend WithEvents cmb10Percent As ComboBox
    Friend WithEvents btnHelp As Button
    Friend WithEvents cmbFilter As ComboBox
    Friend WithEvents GroupApproved As GroupBox
    Friend WithEvents lblApprovedMessage As Label
    Friend WithEvents txbPasswordApprove As TextBox
    Friend WithEvents txbUserApprove As TextBox
    Friend WithEvents lblPasswordA As Label
    Friend WithEvents lblUserIDA As Label
    Friend WithEvents btnLoadMRP As Button
    Friend WithEvents GroupBoxSaved As GroupBox
    Friend WithEvents rdoViewOnly As RadioButton
    Friend WithEvents rdoSaveReport As RadioButton
    Friend WithEvents GridMRP As DataGridView
    Friend WithEvents GroupBoxOption As GroupBox
    Friend WithEvents ckbPastDue As CheckBox
    Friend WithEvents lblWeekTo As Label
    Friend WithEvents lblWeekFrom As Label
    Friend WithEvents lblTo As Label
    Friend WithEvents lblFrom As Label
    Friend WithEvents dtpTo As DateTimePicker
    Friend WithEvents dtpFrom As DateTimePicker
    Friend WithEvents rdoSpecificDates As RadioButton
    Friend WithEvents rdoAllWeeks As RadioButton
    Friend WithEvents btnCalculate As Button
    Friend WithEvents GroupBoxFind As GroupBox
    Friend WithEvents lblFind As Label
    Friend WithEvents txbFind As TextBox
    Friend WithEvents btnFind As Button
    Friend WithEvents btnExportToExcel As Button
    Friend WithEvents lblQty As Label
    Friend WithEvents btnClear As Button
    Friend WithEvents txbQty As TextBox
    Friend WithEvents lblMRPReference As Label
    Friend WithEvents lblTMRPReference As Label
    Friend WithEvents lblRecordsMRP As Label
    Friend WithEvents lblMRP As Label
    Friend WithEvents TabPageBOMWIP As TabPage
    Friend WithEvents GroupBoxAUBOMWIP As GroupBox
    Friend WithEvents Label1 As Label
    Friend WithEvents cmbWIPBOMWIP As ComboBox
    Friend WithEvents cmbAUBOMWIP As ComboBox
    Friend WithEvents cmbRevBOMWIP As ComboBox
    Friend WithEvents lblTRevBOMWIP As Label
    Friend WithEvents lblTAUBOMWIP As Label
    Friend WithEvents GroupBoxPNBOMWIP As GroupBox
    Friend WithEvents txbBOMWIP As TextBox
    Friend WithEvents lblTPNBOMWIP As Label
    Friend WithEvents btnFindBOMWIP As Button
    Friend WithEvents lblRecordsBOMWIP As Label
    Friend WithEvents GridBOMWIP As DataGridView
    Friend WithEvents TabPageBOMENG As TabPage
    Friend WithEvents GroupBoxByAUBOMENG As GroupBox
    Friend WithEvents cmbAUBOMENG As ComboBox
    Friend WithEvents cmbRevBOMENG As ComboBox
    Friend WithEvents lblTAUBOMENG As Label
    Friend WithEvents lblTRevBOMENG As Label
    Friend WithEvents GroupBoxPNBOMENG As GroupBox
    Friend WithEvents txbPNBOMENG As TextBox
    Friend WithEvents lblTPNBOMENG As Label
    Friend WithEvents btnFindBOMENG As Button
    Friend WithEvents lblRecordsBOMENG As Label
    Friend WithEvents GridBOMENG As DataGridView
    Friend WithEvents TabPageMyTable As TabPage
    Friend WithEvents GroupBoxPNMyTable As GroupBox
    Friend WithEvents cmbPNMyTable As ComboBox
    Friend WithEvents lblRecordsMyTable As Label
    Friend WithEvents GridMyTable As DataGridView
    Friend WithEvents TabPageSalesOrder As TabPage
    Friend WithEvents GroupBoxSalesOrderControl As GroupBox
    Friend WithEvents GroupBoxSalesOrderStatus As GroupBox
    Friend WithEvents rdoAllSalesOrderByAU As RadioButton
    Friend WithEvents rdoOpenSalesOrderByAU As RadioButton
    Friend WithEvents rdoCancelSalesOrderByAU As RadioButton
    Friend WithEvents rdoCloseSalesOrderByAU As RadioButton
    Friend WithEvents btnFindSalesOrder As Button
    Friend WithEvents lblTAUSalesOrder As Label
    Friend WithEvents lblTrevSalesOrder As Label
    Friend WithEvents txbAUSalesOrder As TextBox
    Friend WithEvents cmbRevSalesOrder As ComboBox
    Friend WithEvents lblRecordsGridSalesOrder As Label
    Friend WithEvents GridAUSalesOrderFind As DataGridView
    Friend WithEvents TabPageWIPByAU As TabPage
    Friend WithEvents GroupBoxWIPByAU As GroupBox
    Friend WithEvents GroupBoxStatusWIPByAU As GroupBox
    Friend WithEvents rdoAllWipByAU As RadioButton
    Friend WithEvents rdoOpenWipByAU As RadioButton
    Friend WithEvents rdoCancelWipByAU As RadioButton
    Friend WithEvents rdoCloseWipByAU As RadioButton
    Friend WithEvents btnFindWipByAU As Button
    Friend WithEvents lblTAUWipByAU As Label
    Friend WithEvents lblTRevWipByAU As Label
    Friend WithEvents txbAUWipByAU As TextBox
    Friend WithEvents cmbRevWipByAU As ComboBox
    Friend WithEvents lblRecordsWipByAU As Label
    Friend WithEvents GridWipByAU As DataGridView
    Friend WithEvents GroupBoxUserMRP As GroupBox
    Friend WithEvents btnCancelLoginEng As Button
    Friend WithEvents btnLoginMRP As Button
    Friend WithEvents txbUserMRP As TextBox
    Friend WithEvents lblTEngPassword As Label
    Friend WithEvents txbUserMRPPassword As TextBox
    Friend WithEvents lblTEngUser As Label
    Friend WithEvents cmbPONoAprovadas As ComboBox
    Friend WithEvents txbUser As TextBox
    Friend WithEvents ckbConfirmed As CheckBox
End Class
