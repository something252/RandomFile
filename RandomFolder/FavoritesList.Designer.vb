<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FavoritesList
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
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.RemoveFavoriteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenContainingFolderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ListBox1Names = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(22, 24)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(115, 95)
        Me.ListBox1.TabIndex = 0
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.RemoveFavoriteToolStripMenuItem, Me.OpenToolStripMenuItem, Me.OpenContainingFolderToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(202, 70)
        '
        'RemoveFavoriteToolStripMenuItem
        '
        Me.RemoveFavoriteToolStripMenuItem.Name = "RemoveFavoriteToolStripMenuItem"
        Me.RemoveFavoriteToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.RemoveFavoriteToolStripMenuItem.Text = "Unfavorite"
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.OpenToolStripMenuItem.Text = "Open"
        '
        'OpenContainingFolderToolStripMenuItem
        '
        Me.OpenContainingFolderToolStripMenuItem.Name = "OpenContainingFolderToolStripMenuItem"
        Me.OpenContainingFolderToolStripMenuItem.Size = New System.Drawing.Size(201, 22)
        Me.OpenContainingFolderToolStripMenuItem.Text = "Open Containing Folder"
        '
        'ListBox1Names
        '
        Me.ListBox1Names.FormattingEnabled = True
        Me.ListBox1Names.Location = New System.Drawing.Point(156, 24)
        Me.ListBox1Names.Name = "ListBox1Names"
        Me.ListBox1Names.Size = New System.Drawing.Size(120, 95)
        Me.ListBox1Names.TabIndex = 1
        Me.ListBox1Names.Visible = False
        '
        'FavoritesList
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(602, 737)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.ListBox1Names)
        Me.Name = "FavoritesList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Favorites List"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ListBox1Names As System.Windows.Forms.ListBox
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents OpenContainingFolderToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents RemoveFavoriteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
