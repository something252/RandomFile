Public Class Rules
    Dim MaxValue As Integer = 100

    Private Sub Rules_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ComboBox1.SelectedIndex = 0 ' "EVEN"
        ComboBox5.SelectedIndex = 1 ' "ODD"

        ComboBox2.SelectedIndex = 1 ' "and even,"
        ComboBox3.SelectedIndex = 0 ' "do nothing special."
    End Sub

    Private Sub RadioButton_EditType(sender As Object, e As EventArgs) Handles ChanceRadioButton.CheckedChanged, RangeRadioButton.CheckedChanged
        If ChanceRadioButton.Checked Then
            RangeRadioButton.Checked = False ' only one should be set

            ' Chance NumericUpDowns
            GroupBox6.Enabled = True
            GroupBox3.Enabled = True
            GroupBox8.Enabled = True
            GroupBox11.Enabled = True

            ' Range NumericUpDowns
            GroupBox5.Enabled = False
            GroupBox4.Enabled = False
            GroupBox9.Enabled = False
            GroupBox12.Enabled = False
        ElseIf RangeRadioButton.Checked Then
            ChanceRadioButton.Checked = False ' only one should be set

            ' Chance NumericUpDowns
            GroupBox6.Enabled = False
            GroupBox3.Enabled = False
            GroupBox8.Enabled = False
            GroupBox11.Enabled = False

            ' Range NumericUpDowns
            GroupBox5.Enabled = True
            GroupBox4.Enabled = True
            GroupBox9.Enabled = True
            GroupBox12.Enabled = True
        End If
    End Sub

    Private Sub MaxRollNumericUpDown_Value(sender As Object, e As EventArgs) Handles MaxRollNumericUpDown.ValueChanged, MaxRollNumericUpDown.KeyUp
        MaxValue = MaxRollNumericUpDown.Value ' update max value
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        GroupBox1.Visible = Not GroupBox1.Visible
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        GroupBox2.Visible = Not GroupBox2.Visible
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GroupBox7.Visible = Not GroupBox7.Visible
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        GroupBox10.Visible = Not GroupBox10.Visible
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        GroupBox22.Visible = Not GroupBox22.Visible
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        GroupBox19.Visible = Not GroupBox19.Visible
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        GroupBox16.Visible = Not GroupBox16.Visible
    End Sub

End Class