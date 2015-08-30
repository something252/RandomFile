<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Search
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
        Me.components = New System.ComponentModel.Container()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Computer")
        Me.ListViewMain = New System.Windows.Forms.ListView()
        Me.NameColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.DateModifiedColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TypeColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SizeColumn = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OpenContainingFolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.TreeViewAll = New System.Windows.Forms.TreeView()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.OpenContainingFolderToolStripMenuItem2 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExtensionsCheckBox = New System.Windows.Forms.CheckBox()
        Me.RefreshButton = New System.Windows.Forms.Button()
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer()
        Me.SearchContextMenuStrip = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.BreakButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.SearchTextBox = New System.Windows.Forms.TextBox()
        Me.ForwardButton = New System.Windows.Forms.Button()
        Me.UpButton = New System.Windows.Forms.Button()
        Me.AddressTextBox = New System.Windows.Forms.TextBox()
        Me.BackButton = New System.Windows.Forms.Button()
        Me.SearchUpdateTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SearchChangeTimer = New System.Windows.Forms.Timer(Me.components)
        Me.SearchProgress = New System.Windows.Forms.Timer(Me.components)
        Me.ContextMenuStrip1.SuspendLayout()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.ContextMenuStrip2.SuspendLayout()
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SearchContextMenuStrip.SuspendLayout()
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListViewMain
        '
        Me.ListViewMain.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.NameColumn, Me.DateModifiedColumn, Me.TypeColumn, Me.SizeColumn})
        Me.ListViewMain.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListViewMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListViewMain.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListViewMain.Location = New System.Drawing.Point(0, 0)
        Me.ListViewMain.Name = "ListViewMain"
        Me.ListViewMain.Size = New System.Drawing.Size(1033, 770)
        Me.ListViewMain.TabIndex = 1
        Me.ListViewMain.UseCompatibleStateImageBehavior = False
        Me.ListViewMain.View = System.Windows.Forms.View.Details
        '
        'NameColumn
        '
        Me.NameColumn.Text = "Name"
        Me.NameColumn.Width = 476
        '
        'DateModifiedColumn
        '
        Me.DateModifiedColumn.Text = "Date Modified"
        Me.DateModifiedColumn.Width = 139
        '
        'TypeColumn
        '
        Me.TypeColumn.Text = "Type"
        Me.TypeColumn.Width = 78
        '
        'SizeColumn
        '
        Me.SizeColumn.Text = "Size"
        Me.SizeColumn.Width = 81
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenContainingFolderToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(202, 26)
        '
        'OpenContainingFolderToolStripMenuItem
        '
        Me.OpenContainingFolderToolStripMenuItem.Name = "OpenContainingFolderToolStripMenuItem"
        Me.OpenContainingFolderToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.OpenContainingFolderToolStripMenuItem.Text = "Open Containing Folder"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(272, 770)
        Me.TabControl1.TabIndex = 2
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.TreeViewAll)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage1.Size = New System.Drawing.Size(264, 744)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "All"
        '
        'TreeViewAll
        '
        Me.TreeViewAll.ContextMenuStrip = Me.ContextMenuStrip2
        Me.TreeViewAll.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TreeViewAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TreeViewAll.Location = New System.Drawing.Point(3, 3)
        Me.TreeViewAll.Name = "TreeViewAll"
        TreeNode1.Name = "root"
        TreeNode1.Text = "Computer"
        Me.TreeViewAll.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode1})
        Me.TreeViewAll.Size = New System.Drawing.Size(258, 738)
        Me.TreeViewAll.TabIndex = 3
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.OpenContainingFolderToolStripMenuItem2})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip2"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(202, 26)
        '
        'OpenContainingFolderToolStripMenuItem2
        '
        Me.OpenContainingFolderToolStripMenuItem2.Name = "OpenContainingFolderToolStripMenuItem2"
        Me.OpenContainingFolderToolStripMenuItem2.Size = New System.Drawing.Size(201, 22)
        Me.OpenContainingFolderToolStripMenuItem2.Text = "Open Containing Folder"
        '
        'ExtensionsCheckBox
        '
        Me.ExtensionsCheckBox.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.ExtensionsCheckBox.AutoSize = True
        Me.ExtensionsCheckBox.Location = New System.Drawing.Point(7, 20)
        Me.ExtensionsCheckBox.Name = "ExtensionsCheckBox"
        Me.ExtensionsCheckBox.Size = New System.Drawing.Size(122, 17)
        Me.ExtensionsCheckBox.TabIndex = 3
        Me.ExtensionsCheckBox.Text = "Show file extensions"
        Me.ExtensionsCheckBox.UseVisualStyleBackColor = True
        Me.ExtensionsCheckBox.Visible = False
        '
        'RefreshButton
        '
        Me.RefreshButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RefreshButton.Location = New System.Drawing.Point(7, 18)
        Me.RefreshButton.Name = "RefreshButton"
        Me.RefreshButton.Size = New System.Drawing.Size(66, 21)
        Me.RefreshButton.TabIndex = 4
        Me.RefreshButton.Text = "Refresh All"
        Me.RefreshButton.UseVisualStyleBackColor = True
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.TabControl1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.ListViewMain)
        Me.SplitContainer1.Size = New System.Drawing.Size(1309, 770)
        Me.SplitContainer1.SplitterDistance = 272
        Me.SplitContainer1.TabIndex = 3
        '
        'SearchContextMenuStrip
        '
        Me.SearchContextMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.SearchContextMenuStrip.Name = "ContextMenuStrip1"
        Me.SearchContextMenuStrip.Size = New System.Drawing.Size(202, 26)
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(201, 22)
        Me.ToolStripMenuItem1.Text = "Open Containing Folder"
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        Me.SplitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.ProgressBar1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label3)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BreakButton)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.Label1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TextBox2)
        Me.SplitContainer2.Panel1.Controls.Add(Me.TextBox1)
        Me.SplitContainer2.Panel1.Controls.Add(Me.SearchTextBox)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ForwardButton)
        Me.SplitContainer2.Panel1.Controls.Add(Me.UpButton)
        Me.SplitContainer2.Panel1.Controls.Add(Me.AddressTextBox)
        Me.SplitContainer2.Panel1.Controls.Add(Me.BackButton)
        Me.SplitContainer2.Panel1.Controls.Add(Me.RefreshButton)
        Me.SplitContainer2.Panel1.Controls.Add(Me.ExtensionsCheckBox)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.SplitContainer1)
        Me.SplitContainer2.Size = New System.Drawing.Size(1309, 816)
        Me.SplitContainer2.SplitterDistance = 42
        Me.SplitContainer2.TabIndex = 5
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ProgressBar1.Location = New System.Drawing.Point(276, 0)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(755, 18)
        Me.ProgressBar1.TabIndex = 16
        Me.ProgressBar1.Visible = False
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(1104, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(127, 13)
        Me.Label3.TabIndex = 15
        Me.Label3.Text = "(Case Insensitive Search)"
        '
        'BreakButton
        '
        Me.BreakButton.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.BreakButton.Location = New System.Drawing.Point(1240, -3)
        Me.BreakButton.Name = "BreakButton"
        Me.BreakButton.Size = New System.Drawing.Size(68, 23)
        Me.BreakButton.TabIndex = 14
        Me.BreakButton.Text = "Break"
        Me.BreakButton.UseVisualStyleBackColor = True
        Me.BreakButton.Visible = False
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(11, -7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 13
        Me.Label2.Text = "forward dir"
        Me.Label2.Visible = False
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, -34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(61, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "previous dir"
        Me.Label1.Visible = False
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.TextBox2.Location = New System.Drawing.Point(78, -10)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(495, 20)
        Me.TextBox2.TabIndex = 11
        Me.TextBox2.Visible = False
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.TextBox1.Location = New System.Drawing.Point(78, -37)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(495, 20)
        Me.TextBox1.TabIndex = 10
        Me.TextBox1.Visible = False
        '
        'SearchTextBox
        '
        Me.SearchTextBox.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.SearchTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SearchTextBox.ForeColor = System.Drawing.Color.DarkGray
        Me.SearchTextBox.Location = New System.Drawing.Point(1037, 18)
        Me.SearchTextBox.Name = "SearchTextBox"
        Me.SearchTextBox.Size = New System.Drawing.Size(260, 21)
        Me.SearchTextBox.TabIndex = 9
        Me.SearchTextBox.Text = "  Search"
        '
        'ForwardButton
        '
        Me.ForwardButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ForwardButton.Enabled = False
        Me.ForwardButton.Location = New System.Drawing.Point(171, 18)
        Me.ForwardButton.Name = "ForwardButton"
        Me.ForwardButton.Size = New System.Drawing.Size(59, 21)
        Me.ForwardButton.TabIndex = 8
        Me.ForwardButton.Text = "Forward"
        Me.ForwardButton.UseVisualStyleBackColor = True
        '
        'UpButton
        '
        Me.UpButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.UpButton.Enabled = False
        Me.UpButton.Location = New System.Drawing.Point(235, 18)
        Me.UpButton.Name = "UpButton"
        Me.UpButton.Size = New System.Drawing.Size(37, 21)
        Me.UpButton.TabIndex = 7
        Me.UpButton.Text = "Up"
        Me.UpButton.UseVisualStyleBackColor = True
        '
        'AddressTextBox
        '
        Me.AddressTextBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.AddressTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddressTextBox.Location = New System.Drawing.Point(276, 18)
        Me.AddressTextBox.Name = "AddressTextBox"
        Me.AddressTextBox.Size = New System.Drawing.Size(755, 21)
        Me.AddressTextBox.TabIndex = 6
        Me.AddressTextBox.Text = "Computer"
        '
        'BackButton
        '
        Me.BackButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BackButton.Enabled = False
        Me.BackButton.Location = New System.Drawing.Point(106, 18)
        Me.BackButton.Name = "BackButton"
        Me.BackButton.Size = New System.Drawing.Size(59, 21)
        Me.BackButton.TabIndex = 5
        Me.BackButton.Text = "Back"
        Me.BackButton.UseVisualStyleBackColor = True
        '
        'SearchUpdateTimer
        '
        Me.SearchUpdateTimer.Interval = 5000
        '
        'SearchChangeTimer
        '
        Me.SearchChangeTimer.Interval = 250
        '
        'SearchProgress
        '
        Me.SearchProgress.Interval = 150
        '
        'Search
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1309, 816)
        Me.Controls.Add(Me.SplitContainer2)
        Me.MinimumSize = New System.Drawing.Size(804, 170)
        Me.Name = "Search"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Search"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer1.ResumeLayout(False)
        Me.SearchContextMenuStrip.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.PerformLayout()
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListViewMain As System.Windows.Forms.ListView
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents ExtensionsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents RefreshButton As System.Windows.Forms.Button
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents NameColumn As System.Windows.Forms.ColumnHeader
    Friend WithEvents DateModifiedColumn As System.Windows.Forms.ColumnHeader
    Friend WithEvents TypeColumn As System.Windows.Forms.ColumnHeader
    Friend WithEvents SizeColumn As System.Windows.Forms.ColumnHeader
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents OpenContainingFolderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BackButton As System.Windows.Forms.Button
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents TreeViewAll As System.Windows.Forms.TreeView
    Friend WithEvents AddressTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SearchTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ForwardButton As System.Windows.Forms.Button
    Friend WithEvents UpButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ContextMenuStrip2 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents OpenContainingFolderToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchContextMenuStrip As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BreakButton As System.Windows.Forms.Button
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Public WithEvents SearchUpdateTimer As System.Windows.Forms.Timer
    Friend WithEvents SearchChangeTimer As System.Windows.Forms.Timer
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents SearchProgress As System.Windows.Forms.Timer
End Class
