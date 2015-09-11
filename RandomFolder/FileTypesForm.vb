Public Class FileTypesForm

    ''' <summary>
    ''' Handles whitelist and blacklist.
    ''' </summary>
    Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click
        Static locked As Boolean = False
        If Not locked Then
            locked = True

            If Not TextChangedCheck.Text = RichTextBox1.Text Then ' changes were made to whitelist
                RandomFile.RecomputeTreesChangesWereMade = True ' recompute file tree(s)
                Dim str = RichTextBox1.Text
                str = str.Replace(" ", "")
                str = str.Replace(".", "")
                str = str.Replace(Chr(10), "") ' replace enter keys if pasted in ("vbLf")
                Dim FileTypesArray() As String = str.Split(","c)

                My.Settings.WhitelistedFileTypes.Clear()
                If Not RichTextBox1.Text = "" Then
                    Dim index As Integer = 0
                    For Each i In FileTypesArray
                        If Not My.Settings.WhitelistedFileTypes.Contains(i) AndAlso Not FileTypesArray(index) = "" Then ' does not contain the file type
                            My.Settings.WhitelistedFileTypes.Add(i)
                        End If
                        index += 1
                    Next
                End If
            End If

            If Not TextChangedCheck2.Text = RichTextBox2.Text Then ' changes were made to blacklist
                RandomFile.RecomputeTreesChangesWereMade = True ' recompute file tree(s)
                Dim str = RichTextBox2.Text
                str = str.Replace(" ", "")
                str = str.Replace(".", "")
                str = str.Replace(Chr(10), "") ' replace enter keys if pasted in ("vbLf")
                Dim FileTypesArray() As String = str.Split(","c)

                My.Settings.BlacklistedFileTypes.Clear()
                If Not RichTextBox2.Text = "" Then
                    Dim index As Integer = 0
                    For Each i In FileTypesArray
                        If Not My.Settings.BlacklistedFileTypes.Contains(i) AndAlso Not FileTypesArray(index) = "" Then ' does not contain the file type
                            My.Settings.BlacklistedFileTypes.Add(i)
                        End If
                        index += 1
                    Next
                End If
            End If

            Me.Close()

            locked = False
        End If
    End Sub

    Private Sub FileTypesForm_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Me.Icon = My.Resources.icon

        If My.Settings.WhitelistedFileTypes.Count <> 0 Then
            For Each element In My.Settings.WhitelistedFileTypes ' contains list of strings representing whitelisted extension types as only ones allowed
                RichTextBox1.AppendText(element & ",")
            Next
            If RichTextBox1.TextLength > 0 Then
                ' delete last comma
                RichTextBox1.SelectionStart = RichTextBox1.TextLength - 1
                RichTextBox1.SelectionLength = 1
                RichTextBox1.SelectedText = ""
            End If

            TextChangedCheck.Text = RichTextBox1.Text ' used to detect if text changes while form is open
        End If
        If My.Settings.BlacklistedFileTypes.Count <> 0 Then
            For Each element In My.Settings.BlacklistedFileTypes ' contains list of strings representing whitelisted extension types as only ones allowed
                RichTextBox2.AppendText(element & ",")
            Next
            If RichTextBox2.TextLength > 0 Then
                ' delete last comma
                RichTextBox2.SelectionStart = RichTextBox2.TextLength - 1
                RichTextBox2.SelectionLength = 1
                RichTextBox2.SelectedText = ""
            End If

            TextChangedCheck2.Text = RichTextBox2.Text ' used to detect if text changes while form is open
        End If

        If My.Settings.WhitelistEnabled = False Then
            WhitelistEnabledToggle("DISABLED")
        Else
            WhitelistEnabledToggle("ENABLED")
        End If
        If My.Settings.BlacklistEnabled = False Then
            BlacklistEnabledToggle("DISABLED")
        Else
            BlacklistEnabledToggle("ENABLED")
        End If
    End Sub

    Private Sub RichTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox1.KeyDown
        If e.KeyCode = 13 Then ' enter key pressed
            e.SuppressKeyPress = True
        End If
    End Sub
    Private Sub RichTextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles RichTextBox2.KeyDown
        If e.KeyCode = 13 Then ' enter key pressed
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub WhitelistEnabledButton_Click(sender As Object, e As EventArgs) Handles WhitelistEnabledButton.Click
        If WhitelistEnabledButton.Text = "ENABLED" Then
            WhitelistEnabledToggle("DISABLED")
            My.Settings.WhitelistEnabled = False
            RandomFile.WhitelistToolStripMenuItem.Checked = False
        Else
            WhitelistEnabledToggle("ENABLED")
            My.Settings.WhitelistEnabled = True
            RandomFile.WhitelistToolStripMenuItem.Checked = True
        End If
        RandomFile.RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub
    Private Sub BlacklistEnabledButton_Click(sender As Object, e As EventArgs) Handles BlacklistEnabledButton.Click
        If BlacklistEnabledButton.Text = "ENABLED" Then
            BlacklistEnabledToggle("DISABLED")
            My.Settings.BlacklistEnabled = False
            RandomFile.BlacklistToolStripMenuItem.Checked = False
        Else
            BlacklistEnabledToggle("ENABLED")
            My.Settings.BlacklistEnabled = True
            RandomFile.BlacklistToolStripMenuItem.Checked = True
        End If
        RandomFile.RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub

    Sub WhitelistEnabledToggle(str As String)
        If str = "ENABLED" Then
            WhitelistEnabledButton.Text = "ENABLED"
            WhitelistEnabledButton.ForeColor = Color.Blue
        ElseIf str = "DISABLED" Then
            WhitelistEnabledButton.Text = "DISABLED"
            WhitelistEnabledButton.ForeColor = Color.Red
        End If
    End Sub
    Sub BlacklistEnabledToggle(str As String)
        If str = "ENABLED" Then
            BlacklistEnabledButton.Text = "ENABLED"
            BlacklistEnabledButton.ForeColor = Color.Blue
        ElseIf str = "DISABLED" Then
            BlacklistEnabledButton.Text = "DISABLED"
            BlacklistEnabledButton.ForeColor = Color.Red
        End If
    End Sub
End Class