Public Class TrayMenu


    Private Sub TrayMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ContextMenuStrip1.Show(Cursor.Position)
        Me.Left = ContextMenuStrip1.Left + 1 ' put form behind context menu 
        Me.Top = ContextMenuStrip1.Top + 1 ' put form behind context menu 
    End Sub

    Private Sub TrayMenu_Deactivate(sender As Object, e As EventArgs) Handles MyBase.Deactivate
        Me.Close()
    End Sub

    ' terminate the program
    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        RandomFolder.Close() ' close the application
    End Sub

    Private Sub ShowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowToolStripMenuItem.Click
        RandomFolder.Show()
        RandomFolder.WindowState = FormWindowState.Normal
        Me.Close()
    End Sub

End Class