<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class OctoPart_API
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
        Me.dgvOctopartResults = New System.Windows.Forms.DataGridView()
        Me.dgvBuildPR = New System.Windows.Forms.DataGridView()
        Me.labelOctopartResults = New System.Windows.Forms.Label()
        Me.labelBuildPR = New System.Windows.Forms.Label()
        Me.labelKeyword = New System.Windows.Forms.Label()
        Me.tbKeyword = New System.Windows.Forms.TextBox()
        Me.cbCategories = New System.Windows.Forms.ComboBox()
        Me.cbSubcategories = New System.Windows.Forms.ComboBox()
        Me.labelCategories = New System.Windows.Forms.Label()
        Me.labelSubcategories = New System.Windows.Forms.Label()
        Me.groupFilters = New System.Windows.Forms.GroupBox()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.statusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolStripProgressBar1 = New System.Windows.Forms.ToolStripProgressBar()
        Me.btnGetTotal = New System.Windows.Forms.Button()
        Me.btnExportPR = New System.Windows.Forms.Button()
        Me.dlgSaveFile = New System.Windows.Forms.SaveFileDialog()
        Me.cbSources = New System.Windows.Forms.ComboBox()
        Me.labelSource = New System.Windows.Forms.Label()
        Me.lblCurrPage = New System.Windows.Forms.Label()
        Me.lblNumberOfPages = New System.Windows.Forms.Label()
        Me.shopeeNavPanel = New System.Windows.Forms.GroupBox()
        Me.btnPrevious = New System.Windows.Forms.Button()
        Me.btnNext = New System.Windows.Forms.Button()
        CType(Me.dgvOctopartResults, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBuildPR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupFilters.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.shopeeNavPanel.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvOctopartResults
        '
        Me.dgvOctopartResults.AllowUserToAddRows = False
        Me.dgvOctopartResults.AllowUserToDeleteRows = False
        Me.dgvOctopartResults.AllowUserToResizeColumns = False
        Me.dgvOctopartResults.AllowUserToResizeRows = False
        Me.dgvOctopartResults.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvOctopartResults.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvOctopartResults.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvOctopartResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvOctopartResults.Location = New System.Drawing.Point(280, 40)
        Me.dgvOctopartResults.Name = "dgvOctopartResults"
        Me.dgvOctopartResults.ReadOnly = True
        Me.dgvOctopartResults.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvOctopartResults.Size = New System.Drawing.Size(756, 349)
        Me.dgvOctopartResults.TabIndex = 0
        '
        'dgvBuildPR
        '
        Me.dgvBuildPR.AllowUserToResizeColumns = False
        Me.dgvBuildPR.AllowUserToResizeRows = False
        Me.dgvBuildPR.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvBuildPR.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.dgvBuildPR.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvBuildPR.Location = New System.Drawing.Point(280, 493)
        Me.dgvBuildPR.Name = "dgvBuildPR"
        Me.dgvBuildPR.Size = New System.Drawing.Size(756, 196)
        Me.dgvBuildPR.TabIndex = 1
        '
        'labelOctopartResults
        '
        Me.labelOctopartResults.AutoSize = True
        Me.labelOctopartResults.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelOctopartResults.Location = New System.Drawing.Point(280, 13)
        Me.labelOctopartResults.Name = "labelOctopartResults"
        Me.labelOctopartResults.Size = New System.Drawing.Size(139, 20)
        Me.labelOctopartResults.TabIndex = 2
        Me.labelOctopartResults.Text = "OctoPart results"
        '
        'labelBuildPR
        '
        Me.labelBuildPR.AutoSize = True
        Me.labelBuildPR.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelBuildPR.Location = New System.Drawing.Point(280, 411)
        Me.labelBuildPR.Name = "labelBuildPR"
        Me.labelBuildPR.Size = New System.Drawing.Size(78, 20)
        Me.labelBuildPR.TabIndex = 3
        Me.labelBuildPR.Text = "Build PR"
        '
        'labelKeyword
        '
        Me.labelKeyword.AutoSize = True
        Me.labelKeyword.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelKeyword.Location = New System.Drawing.Point(16, 80)
        Me.labelKeyword.Name = "labelKeyword"
        Me.labelKeyword.Size = New System.Drawing.Size(67, 16)
        Me.labelKeyword.TabIndex = 6
        Me.labelKeyword.Text = "Keyword"
        '
        'tbKeyword
        '
        Me.tbKeyword.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbKeyword.Location = New System.Drawing.Point(16, 104)
        Me.tbKeyword.Name = "tbKeyword"
        Me.tbKeyword.Size = New System.Drawing.Size(232, 22)
        Me.tbKeyword.TabIndex = 7
        '
        'cbCategories
        '
        Me.cbCategories.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cbCategories.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cbCategories.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbCategories.FormattingEnabled = True
        Me.cbCategories.Location = New System.Drawing.Point(16, 48)
        Me.cbCategories.Name = "cbCategories"
        Me.cbCategories.Size = New System.Drawing.Size(232, 24)
        Me.cbCategories.TabIndex = 14
        '
        'cbSubcategories
        '
        Me.cbSubcategories.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cbSubcategories.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cbSubcategories.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSubcategories.FormattingEnabled = True
        Me.cbSubcategories.Location = New System.Drawing.Point(16, 104)
        Me.cbSubcategories.Name = "cbSubcategories"
        Me.cbSubcategories.Size = New System.Drawing.Size(232, 24)
        Me.cbSubcategories.TabIndex = 16
        '
        'labelCategories
        '
        Me.labelCategories.AutoSize = True
        Me.labelCategories.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Bold)
        Me.labelCategories.Location = New System.Drawing.Point(16, 24)
        Me.labelCategories.Name = "labelCategories"
        Me.labelCategories.Size = New System.Drawing.Size(76, 15)
        Me.labelCategories.TabIndex = 15
        Me.labelCategories.Text = "Categories"
        '
        'labelSubcategories
        '
        Me.labelSubcategories.AutoSize = True
        Me.labelSubcategories.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.5!, System.Drawing.FontStyle.Bold)
        Me.labelSubcategories.Location = New System.Drawing.Point(16, 80)
        Me.labelSubcategories.Name = "labelSubcategories"
        Me.labelSubcategories.Size = New System.Drawing.Size(99, 15)
        Me.labelSubcategories.TabIndex = 17
        Me.labelSubcategories.Text = "Subcategories"
        '
        'groupFilters
        '
        Me.groupFilters.Controls.Add(Me.labelSubcategories)
        Me.groupFilters.Controls.Add(Me.labelCategories)
        Me.groupFilters.Controls.Add(Me.cbSubcategories)
        Me.groupFilters.Controls.Add(Me.cbCategories)
        Me.groupFilters.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.groupFilters.Location = New System.Drawing.Point(8, 136)
        Me.groupFilters.Name = "groupFilters"
        Me.groupFilters.Size = New System.Drawing.Size(264, 144)
        Me.groupFilters.TabIndex = 10
        Me.groupFilters.TabStop = False
        Me.groupFilters.Text = "Search Filters"
        '
        'btnSearch
        '
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(72, 144)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(128, 56)
        Me.btnSearch.TabIndex = 18
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusLabel, Me.ToolStripProgressBar1})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 729)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1054, 22)
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'statusLabel
        '
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(119, 17)
        Me.statusLabel.Text = "ToolStripStatusLabel1"
        '
        'ToolStripProgressBar1
        '
        Me.ToolStripProgressBar1.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.ToolStripProgressBar1.Name = "ToolStripProgressBar1"
        Me.ToolStripProgressBar1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ToolStripProgressBar1.Size = New System.Drawing.Size(100, 16)
        Me.ToolStripProgressBar1.Visible = False
        '
        'btnGetTotal
        '
        Me.btnGetTotal.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnGetTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGetTotal.Location = New System.Drawing.Point(312, 694)
        Me.btnGetTotal.Name = "btnGetTotal"
        Me.btnGetTotal.Size = New System.Drawing.Size(336, 32)
        Me.btnGetTotal.TabIndex = 20
        Me.btnGetTotal.Text = "Get Total"
        Me.btnGetTotal.UseVisualStyleBackColor = True
        '
        'btnExportPR
        '
        Me.btnExportPR.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnExportPR.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExportPR.Location = New System.Drawing.Point(660, 694)
        Me.btnExportPR.Name = "btnExportPR"
        Me.btnExportPR.Size = New System.Drawing.Size(336, 32)
        Me.btnExportPR.TabIndex = 21
        Me.btnExportPR.Text = "Export PR"
        Me.btnExportPR.UseVisualStyleBackColor = True
        '
        'cbSources
        '
        Me.cbSources.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Append
        Me.cbSources.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cbSources.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbSources.FormattingEnabled = True
        Me.cbSources.Location = New System.Drawing.Point(16, 48)
        Me.cbSources.Name = "cbSources"
        Me.cbSources.Size = New System.Drawing.Size(232, 24)
        Me.cbSources.TabIndex = 22
        '
        'labelSource
        '
        Me.labelSource.AutoSize = True
        Me.labelSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelSource.Location = New System.Drawing.Point(16, 24)
        Me.labelSource.Name = "labelSource"
        Me.labelSource.Size = New System.Drawing.Size(57, 16)
        Me.labelSource.TabIndex = 23
        Me.labelSource.Text = "Source"
        '
        'lblCurrPage
        '
        Me.lblCurrPage.AutoSize = True
        Me.lblCurrPage.BackColor = System.Drawing.Color.Transparent
        Me.lblCurrPage.ForeColor = System.Drawing.Color.FromArgb(CType(CType(210, Byte), Integer), CType(CType(123, Byte), Integer), CType(CType(87, Byte), Integer))
        Me.lblCurrPage.Location = New System.Drawing.Point(62, 21)
        Me.lblCurrPage.Name = "lblCurrPage"
        Me.lblCurrPage.Size = New System.Drawing.Size(13, 13)
        Me.lblCurrPage.TabIndex = 24
        Me.lblCurrPage.Text = "1"
        Me.lblCurrPage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblNumberOfPages
        '
        Me.lblNumberOfPages.AutoSize = True
        Me.lblNumberOfPages.Location = New System.Drawing.Point(72, 21)
        Me.lblNumberOfPages.Name = "lblNumberOfPages"
        Me.lblNumberOfPages.Size = New System.Drawing.Size(30, 13)
        Me.lblNumberOfPages.TabIndex = 25
        Me.lblNumberOfPages.Text = "/100"
        Me.lblNumberOfPages.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'shopeeNavPanel
        '
        Me.shopeeNavPanel.Controls.Add(Me.btnNext)
        Me.shopeeNavPanel.Controls.Add(Me.lblNumberOfPages)
        Me.shopeeNavPanel.Controls.Add(Me.btnPrevious)
        Me.shopeeNavPanel.Controls.Add(Me.lblCurrPage)
        Me.shopeeNavPanel.Location = New System.Drawing.Point(874, 395)
        Me.shopeeNavPanel.Name = "shopeeNavPanel"
        Me.shopeeNavPanel.Size = New System.Drawing.Size(162, 47)
        Me.shopeeNavPanel.TabIndex = 27
        Me.shopeeNavPanel.TabStop = False
        Me.shopeeNavPanel.Text = "Shopee Page Control"
        Me.shopeeNavPanel.Visible = False
        '
        'btnPrevious
        '
        Me.btnPrevious.Enabled = False
        Me.btnPrevious.Location = New System.Drawing.Point(28, 16)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(28, 23)
        Me.btnPrevious.TabIndex = 28
        Me.btnPrevious.Text = "<"
        Me.btnPrevious.UseVisualStyleBackColor = True
        '
        'btnNext
        '
        Me.btnNext.Location = New System.Drawing.Point(108, 16)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(28, 23)
        Me.btnNext.TabIndex = 28
        Me.btnNext.Text = ">"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'OctoPart_API
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1054, 751)
        Me.Controls.Add(Me.shopeeNavPanel)
        Me.Controls.Add(Me.labelSource)
        Me.Controls.Add(Me.cbSources)
        Me.Controls.Add(Me.btnExportPR)
        Me.Controls.Add(Me.btnGetTotal)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.groupFilters)
        Me.Controls.Add(Me.tbKeyword)
        Me.Controls.Add(Me.labelKeyword)
        Me.Controls.Add(Me.labelBuildPR)
        Me.Controls.Add(Me.labelOctopartResults)
        Me.Controls.Add(Me.dgvBuildPR)
        Me.Controls.Add(Me.dgvOctopartResults)
        Me.Name = "OctoPart_API"
        Me.Text = "OctoPart_API"
        CType(Me.dgvOctopartResults, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvBuildPR, System.ComponentModel.ISupportInitialize).EndInit()
        Me.groupFilters.ResumeLayout(False)
        Me.groupFilters.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.shopeeNavPanel.ResumeLayout(False)
        Me.shopeeNavPanel.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgvOctopartResults As DataGridView
    Friend WithEvents dgvBuildPR As DataGridView
    Friend WithEvents labelOctopartResults As Label
    Friend WithEvents labelBuildPR As Label
    Friend WithEvents labelKeyword As Label
    Friend WithEvents tbKeyword As TextBox
    Friend WithEvents cbCategories As ComboBox
    Friend WithEvents cbSubcategories As ComboBox
    Friend WithEvents labelCategories As Label
    Friend WithEvents labelSubcategories As Label
    Friend WithEvents groupFilters As GroupBox
    Friend WithEvents btnSearch As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents statusLabel As ToolStripStatusLabel
    Friend WithEvents btnGetTotal As Button
    Friend WithEvents btnExportPR As Button
    Friend WithEvents dlgSaveFile As SaveFileDialog
    Friend WithEvents ToolStripProgressBar1 As ToolStripProgressBar
    Friend WithEvents cbSources As ComboBox
    Friend WithEvents labelSource As Label
    Friend WithEvents lblCurrPage As Label
    Friend WithEvents lblNumberOfPages As Label
    Friend WithEvents shopeeNavPanel As GroupBox
    Friend WithEvents btnNext As Button
    Friend WithEvents btnPrevious As Button
End Class
