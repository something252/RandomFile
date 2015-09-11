Imports System.IO

Public Class FavoritesList


    Private Sub FavoritesList_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.icon

        If Not IsNothing(My.Settings.FavoritesListSize) Then
            If My.Settings.FavoritesListSize.Width > 0 AndAlso My.Settings.FavoritesListSize.Height > 0 Then
                Me.Size = My.Settings.FavoritesListSize
            End If
        End If

        'Me.Location = New Point(-4444, -4242)

        If Not IsNothing(My.Settings.FavoritesListPosition) AndAlso
            (My.Settings.FavoritesListPosition.X <> -1 AndAlso My.Settings.FavoritesListPosition.Y <> -1) Then
            If RandomFile.FormVisible(My.Settings.FavoritesListPosition) OrElse _
                RandomFile.FormVisible(New Point(My.Settings.FavoritesListPosition.X + Me.Size.Width, My.Settings.FavoritesListPosition.Y)) Then
                Me.Location = My.Settings.FavoritesListPosition
            Else
                Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width / 2) - (Me.Size.Width / 2), (Screen.PrimaryScreen.WorkingArea.Height / 2) - (Me.Size.Height / 2))
            End If
        End If


        If ListBox1.Items.Count > 0 Then
            ListBox1.Items.Clear()
        End If
        If ListBox1Names.Items.Count > 0 Then
            ListBox1Names.Items.Clear()
        End If

        If Not IsNothing(My.Settings.FavoriteRollsList) Then
            For Each item As String In My.Settings.FavoriteRollsList
                If File.Exists(item) Then
                    ListBox1.Items.Add(Search.RemoveExtension(Search.GetFileName(item))) ' partial path
                Else
                    ListBox1.Items.Add(Search.GetFileName(item)) ' partial path
                End If
                ListBox1Names.Items.Add(item) ' full path
            Next
        End If

        ListBox1.Dock = DockStyle.Fill
        ListBox1.Visible = True
    End Sub

    Private Sub FavoritesList_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

    End Sub

    Private Sub FavoritesList_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Me.WindowState = FormWindowState.Minimized Then ' minimizing makes location -32000 coords
            Me.Opacity = 0.0
            Me.WindowState = FormWindowState.Normal
        End If
        My.Settings.FavoritesListSize = Me.Size
        My.Settings.FavoritesListPosition = Me.Location
    End Sub

    ''' <summary>
    ''' Delete the selected item in the ListBox and user settings.
    ''' </summary>
    Private Sub DeleteToolStripMenuItem1_Click() Handles RemoveFavoriteToolStripMenuItem.Click
        Dim selectedIndex As Integer = ListBox1.SelectedIndex
        If Not IsNothing(selectedIndex) AndAlso selectedIndex >= 0 Then

            Dim Name As String = ListBox1Names.Items.Item(selectedIndex)

            ' delete from settings
            If My.Settings.FavoriteRollsList.Contains(Name) Then
                My.Settings.FavoriteRollsList.Remove(Name)
            End If

            If RandomFile.ResultBox.Text = Name Then
                RandomFile.NonFavoriteRolled()
            End If

            ' delete item in listboxes
            ListBox1.Items.RemoveAt(selectedIndex)
            ListBox1Names.Items.RemoveAt(selectedIndex)

        End If
    End Sub

    ''' <summary>
    ''' Open containing folder.
    ''' </summary>
    Private Sub OpenContainingFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenContainingFolderToolStripMenuItem.Click
        Dim selectedIndex As Integer = ListBox1.SelectedIndex
        If Not IsNothing(selectedIndex) AndAlso selectedIndex >= 0 Then
            If File.Exists(ListBox1Names.Items.Item(selectedIndex)) OrElse Directory.Exists(ListBox1Names.Items.Item(selectedIndex)) Then
                Call Shell("explorer /select," & ListBox1Names.Items.Item(selectedIndex), AppWinStyle.NormalFocus) ' select in containing folder
            Else
                MsgBox("File or folder no longer exists!" & vbNewLine & """" & ListBox1Names.Items.Item(selectedIndex) & """", MsgBoxStyle.Critical, "Error")
            End If
        End If
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        OpenListBoxItem()
    End Sub

    Private Sub OpenToolStripMenuItem_Click() Handles OpenToolStripMenuItem.Click
        OpenListBoxItem()
    End Sub

    Private Sub OpenListBoxItem()
        Dim selectedIndex As Integer = ListBox1.SelectedIndex
        If Not IsNothing(selectedIndex) AndAlso selectedIndex >= 0 Then
            If File.Exists(ListBox1Names.Items.Item(selectedIndex)) OrElse Directory.Exists(ListBox1Names.Items.Item(selectedIndex)) Then
                Try
                    Process.Start(ListBox1Names.Items.Item(selectedIndex))
                Catch
                End Try
            Else
                MsgBox("File or folder no longer exists!" & vbNewLine & """" & ListBox1Names.Items.Item(selectedIndex) & """", MsgBoxStyle.Critical, "Error")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Checks if a full path is a shortcut, or basically if it has the .lnk extension.
    ''' </summary>
    ''' <param name="str">Full path string</param>
    Private Function CheckIfShortcut(ByRef str As String) As Boolean ' checks if a arrow should be put on an icon's image (if shortcut)
        Dim myStr() As String = Split(str, "\")
        Dim myStr2() As String = Split(myStr(myStr.Length - 1), ".")

        If myStr2(myStr2.Length - 1) = "lnk" Then
            Return True ' is a shortcut
        Else
            Return False
        End If
    End Function

    Private Sub ListBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListBox1.KeyDown
        If e.KeyCode = Keys.Delete Then
            DeleteToolStripMenuItem1_Click()
        End If
    End Sub
End Class