﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.btnGetTotal = New System.Windows.Forms.Button()
        Me.btnExportPR = New System.Windows.Forms.Button()
        Me.dlgSaveFile = New System.Windows.Forms.SaveFileDialog()
        CType(Me.dgvOctopartResults, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvBuildPR, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.groupFilters.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
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
        Me.dgvOctopartResults.Location = New System.Drawing.Point(272, 40)
        Me.dgvOctopartResults.Name = "dgvOctopartResults"
        Me.dgvOctopartResults.ReadOnly = True
        Me.dgvOctopartResults.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.dgvOctopartResults.Size = New System.Drawing.Size(760, 380)
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
        Me.dgvBuildPR.Location = New System.Drawing.Point(272, 448)
        Me.dgvBuildPR.Name = "dgvBuildPR"
        Me.dgvBuildPR.Size = New System.Drawing.Size(760, 241)
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
        Me.labelBuildPR.Location = New System.Drawing.Point(280, 424)
        Me.labelBuildPR.Name = "labelBuildPR"
        Me.labelBuildPR.Size = New System.Drawing.Size(78, 20)
        Me.labelBuildPR.TabIndex = 3
        Me.labelBuildPR.Text = "Build PR"
        '
        'labelKeyword
        '
        Me.labelKeyword.AutoSize = True
        Me.labelKeyword.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.labelKeyword.Location = New System.Drawing.Point(16, 16)
        Me.labelKeyword.Name = "labelKeyword"
        Me.labelKeyword.Size = New System.Drawing.Size(67, 16)
        Me.labelKeyword.TabIndex = 6
        Me.labelKeyword.Text = "Keyword"
        '
        'tbKeyword
        '
        Me.tbKeyword.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbKeyword.Location = New System.Drawing.Point(16, 40)
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
        Me.groupFilters.Location = New System.Drawing.Point(8, 72)
        Me.groupFilters.Name = "groupFilters"
        Me.groupFilters.Size = New System.Drawing.Size(264, 144)
        Me.groupFilters.TabIndex = 10
        Me.groupFilters.TabStop = False
        Me.groupFilters.Text = "Search Filters"
        '
        'btnSearch
        '
        Me.btnSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearch.Location = New System.Drawing.Point(72, 232)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(128, 56)
        Me.btnSearch.TabIndex = 18
        Me.btnSearch.Text = "Search"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statusLabel})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 729)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1050, 22)
        Me.StatusStrip1.TabIndex = 19
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'statusLabel
        '
        Me.statusLabel.Name = "statusLabel"
        Me.statusLabel.Size = New System.Drawing.Size(119, 17)
        Me.statusLabel.Text = "ToolStripStatusLabel1"
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
        Me.btnExportPR.Location = New System.Drawing.Point(656, 694)
        Me.btnExportPR.Name = "btnExportPR"
        Me.btnExportPR.Size = New System.Drawing.Size(336, 32)
        Me.btnExportPR.TabIndex = 21
        Me.btnExportPR.Text = "Export PR"
        Me.btnExportPR.UseVisualStyleBackColor = True
        '
        'OctoPart_API
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1050, 751)
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
End Class