Public Class ParentFolders
    ''' <summary>
    ''' Each directory in array points to directory entry textboxes in respective order starting from 0
    ''' </summary>
    Dim Directories(29) As Object
    ''' <summary>
    ''' Contains weights of each respective directories' variable index.
    ''' </summary>
    Dim Weights(29) As Object
    ''' <summary>
    ''' Contains weight percent value for use in displaying percentage weight in this form.
    ''' </summary>
    Dim WeightsDisplay(29) As Object
    ''' <summary>
    ''' Local version of RandomFile variable for pre-confirmation holding. (Gets put into main form variable upon confirmation success)
    ''' </summary>
    Dim DisabledLine(29) As Boolean
    ''' <summary>
    ''' Contains group value for each row.
    ''' </summary>
    Dim Groups(29) As Object
    ''' <summary>
    ''' Used to prevent multiple clicks, not necessary without threading.
    ''' </summary>
    Dim locked As Boolean = False
    ''' <summary>
    ''' Flag for when changes were made to an open save.
    ''' </summary>
    Dim ChangesWereMade As Boolean = False
    ''' <summary>
    ''' Points to each saves' respective My.Settings directories for each row.
    ''' </summary>
    Dim DirectoriesSavePtr() As Object = {0, My.Settings.DirectoriesSave1, My.Settings.DirectoriesSave2, My.Settings.DirectoriesSave3, My.Settings.DirectoriesSave4, My.Settings.DirectoriesSave5, _
                                          My.Settings.DirectoriesSave6, My.Settings.DirectoriesSave7, My.Settings.DirectoriesSave8, My.Settings.DirectoriesSave9, My.Settings.DirectoriesSave10}
    ''' <summary>
    ''' Points to each saves' respective My.Settings weights for each row.
    ''' </summary>
    Dim WeightsSavePtr() As Object = {0, My.Settings.WeightsSave1, My.Settings.WeightsSave2, My.Settings.WeightsSave3, My.Settings.WeightsSave4, My.Settings.WeightsSave5, _
                                      My.Settings.WeightsSave6, My.Settings.WeightsSave7, My.Settings.WeightsSave8, My.Settings.WeightsSave9, My.Settings.WeightsSave10}
    ''' <summary>
    ''' Points to each saves' respective My.Settings disabled line state for each row.
    ''' </summary>
    Dim DisabledLineSavePtr() As Object = {0, My.Settings.DisabledLineSave1, My.Settings.DisabledLineSave2, My.Settings.DisabledLineSave3, My.Settings.DisabledLineSave4, My.Settings.DisabledLineSave5, _
                                           My.Settings.DisabledLineSave6, My.Settings.DisabledLineSave7, My.Settings.DisabledLineSave8, My.Settings.DisabledLineSave9, My.Settings.DisabledLineSave10}
    ''' <summary>
    ''' Points to each saves' respective My.Settings groups identifiers for each row.
    ''' </summary>
    Dim GroupsSavePtr() As Object = {0, My.Settings.GroupsSave1, My.Settings.GroupsSave2, My.Settings.GroupsSave3, My.Settings.GroupsSave4, My.Settings.GroupsSave5, _
                                           My.Settings.GroupsSave6, My.Settings.GroupsSave7, My.Settings.GroupsSave8, My.Settings.GroupsSave9, My.Settings.GroupsSave10}
    ''' <summary>
    ''' Handles local current active save number which is then dumped into My.Settings on confirmation button success. (Defined in Shown event)
    ''' </summary>
    Dim CurrentActiveSave As Integer

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        ToggleEnabled(0)
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        ToggleEnabled(1)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        ToggleEnabled(2)
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        ToggleEnabled(3)
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        ToggleEnabled(4)
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        ToggleEnabled(5)
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        ToggleEnabled(6)
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        ToggleEnabled(7)
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        ToggleEnabled(8)
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        ToggleEnabled(9)
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        ToggleEnabled(10)
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        ToggleEnabled(11)
    End Sub

    Private Sub TextBox13_TextChanged(sender As Object, e As EventArgs) Handles TextBox13.TextChanged
        ToggleEnabled(12)
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        ToggleEnabled(13)
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        ToggleEnabled(14)
    End Sub

    Private Sub TextBox16_TextChanged(sender As Object, e As EventArgs) Handles TextBox16.TextChanged
        ToggleEnabled(15)
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        ToggleEnabled(16)
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        ToggleEnabled(17)
    End Sub

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        ToggleEnabled(18)
    End Sub

    Private Sub TextBox20_TextChanged(sender As Object, e As EventArgs) Handles TextBox20.TextChanged
        ToggleEnabled(19)
    End Sub

    Private Sub TextBox21_TextChanged(sender As Object, e As EventArgs) Handles TextBox21.TextChanged
        ToggleEnabled(20)
    End Sub

    Private Sub TextBox22_TextChanged(sender As Object, e As EventArgs) Handles TextBox22.TextChanged
        ToggleEnabled(21)
    End Sub

    Private Sub TextBox23_TextChanged(sender As Object, e As EventArgs) Handles TextBox23.TextChanged
        ToggleEnabled(22)
    End Sub

    Private Sub TextBox24_TextChanged(sender As Object, e As EventArgs) Handles TextBox24.TextChanged
        ToggleEnabled(23)
    End Sub

    Private Sub TextBox25_TextChanged(sender As Object, e As EventArgs) Handles TextBox25.TextChanged
        ToggleEnabled(24)
    End Sub

    Private Sub TextBox26_TextChanged(sender As Object, e As EventArgs) Handles TextBox26.TextChanged
        ToggleEnabled(25)
    End Sub

    Private Sub TextBox27_TextChanged(sender As Object, e As EventArgs) Handles TextBox27.TextChanged
        ToggleEnabled(26)
    End Sub

    Private Sub TextBox28_TextChanged(sender As Object, e As EventArgs) Handles TextBox28.TextChanged
        ToggleEnabled(27)
    End Sub

    Private Sub TextBox29_TextChanged(sender As Object, e As EventArgs) Handles TextBox29.TextChanged
        ToggleEnabled(28)
    End Sub

    Private Sub TextBox30_TextChanged(sender As Object, e As EventArgs) Handles TextBox30.TextChanged
        ToggleEnabled(29)
    End Sub

    ''' <summary>
    ''' Toggle directories' respective weight control.
    ''' </summary>
    ''' <param name="i">Index</param>
    Private Sub ToggleEnabled(ByVal i As Integer)
        If shownComplete = True Then
            If Directories(i).Text = "" Then
                Weights(i).Text = "1"
                Weights(i).Enabled = False ' Toggle
                Groups(i).Enabled = False ' Toggle
                Groups(i).SelectedItem = ""
            Else
                Weights(i).Enabled = True ' Toggle
                Groups(i).Enabled = True ' Toggle
            End If
            UpdateWeightDisplays() ' update displayed weight percentages
            ChangesWereMade = True
        End If
    End Sub

    ''' <summary>
    ''' Check if any of the inputted directories don't exist.
    ''' </summary>
    Private Function CheckDirectories() As Boolean
        Dim Flag As Boolean = False

        For i As Integer = 0 To Directories.Count - 1
            If Not (Directories(i).Text = "" OrElse Me.DisabledLine(i) = True) Then
                If Not System.IO.Directory.Exists(Directories(i).Text) Then
                    Flag = True
                    MsgBox("Row " & i + 1 & ": " & """" & Directories(i).Text & """" & vbNewLine & "Invalid directory.", MsgBoxStyle.Exclamation, "Error")
                End If
            End If
        Next

        If Flag = True Then
            'MsgBox("One or more folder paths provided do not exist.", MsgBoxStyle.Critical, "Error")
            Return False ' directory paths entered do not exist
        Else
            Return True ' directory paths entered exist
        End If
    End Function

    Private Sub ConfirmButton_Click(sender As Object, e As EventArgs) Handles ConfirmButton.Click
        If (Not locked) Then
            locked = True
            Try
                Dim total As Integer = 0
                For i As Integer = 0 To Directories.Count - 1 ' find sum of all weights
                    If Not (Directories(i).Text = "" OrElse Me.DisabledLine(i) = True) Then
                        total += Weights(i).Value
                    End If
                Next
                If total = 0 Then ' check if weights is equal to 0
                    MsgBox("Total of all weights should not equal 0 and all rows should not be empty or deactivated.", MsgBoxStyle.Critical, "Error")
                    locked = False
                    Exit Sub ' terminate any further button executation
                Else
                    RandomFile.TotalWeight = total ' update total weight
                End If

                DisableLineUpdate() ' update lines disabled variable in main form
                GroupsUpdate() ' update lines group values variable in main form

                If CheckDirectories() = True Then
                    ' Everything checks out, now do work
                    If My.Settings.AlreadyRolledListFlag = True Then
                        RandomFile.CreateRollOnceList() ' generate roll once list
                    End If
                    UpdateInfo()
                    If Not IsNothing(RandomFile.GroupWeightInfo) Then
                        RandomFile.GroupWeightInfo.Clear()
                    End If
                    RandomFile.GroupWeightInfo = Me.GetGroupWeights() ' store group weights info for UpdateTrees()
                    RandomFile.UpdateTrees() ' update file and folder trees
                    RandomFile.MultipleFoldersHandled = True ' Successful handling has happened
                    My.Settings.ParentFoldersLastActiveSave = CurrentActiveSave ' set new last active save value
                    RandomFile.RecomputeTreesChangesWereMade = False ' don't double compute tree(s) unnecessarily
                    RandomFile.RollButton.PerformClick()
                    If Search.Visible = True Then
                        Search.RefreshButton_Click()
                    End If
                    Me.Close() ' Close this window
                End If

            Catch err As Exception
                MsgBox(err.Message, MsgBoxStyle.Critical, "Validation Error")
            End Try
            locked = False
        End If
    End Sub

    Private Sub UpdateInfo()
        For y As Integer = 0 To Directories.Count - 1
            If Not (Directories(y).Text = "") Then ' if not empty
                RandomFile.FolderDirectories(y) = Directories(y).Text
                RandomFile.FolderWeights(y) = CInt(Weights(y).Text)
            Else ' empty
                RandomFile.FolderDirectories(y) = "" ' mark as empty
                RandomFile.FolderWeights(y) = 0
            End If
        Next

        SaveActiveSave() ' save the active save to My.Settings
    End Sub

    Private Sub ParentFolders_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.icon

        ' assign directory control objects
        Directories(0) = TextBox1 : Directories(1) = TextBox2 : Directories(2) = TextBox3
        Directories(3) = TextBox4 : Directories(4) = TextBox5 : Directories(5) = TextBox6
        Directories(6) = TextBox7 : Directories(7) = TextBox8 : Directories(8) = TextBox9
        Directories(9) = TextBox10 : Directories(10) = TextBox11 : Directories(11) = TextBox12
        Directories(12) = TextBox13 : Directories(13) = TextBox14 : Directories(14) = TextBox15
        Directories(15) = TextBox16 : Directories(16) = TextBox17 : Directories(17) = TextBox18
        Directories(18) = TextBox19 : Directories(19) = TextBox20 : Directories(20) = TextBox21
        Directories(21) = TextBox22 : Directories(22) = TextBox23 : Directories(23) = TextBox24
        Directories(24) = TextBox25 : Directories(25) = TextBox26 : Directories(26) = TextBox27
        Directories(27) = TextBox28 : Directories(28) = TextBox29 : Directories(29) = TextBox30

        ' assign weight control objects
        Weights(0) = NumericUpDown1 : Weights(1) = NumericUpDown2 : Weights(2) = NumericUpDown3
        Weights(3) = NumericUpDown4 : Weights(4) = NumericUpDown5 : Weights(5) = NumericUpDown6
        Weights(6) = NumericUpDown7 : Weights(7) = NumericUpDown8 : Weights(8) = NumericUpDown9
        Weights(9) = NumericUpDown10 : Weights(10) = NumericUpDown11 : Weights(11) = NumericUpDown12
        Weights(12) = NumericUpDown13 : Weights(13) = NumericUpDown14 : Weights(14) = NumericUpDown15
        Weights(15) = NumericUpDown16 : Weights(16) = NumericUpDown17 : Weights(17) = NumericUpDown18
        Weights(18) = NumericUpDown19 : Weights(19) = NumericUpDown20 : Weights(20) = NumericUpDown21
        Weights(21) = NumericUpDown22 : Weights(22) = NumericUpDown23 : Weights(23) = NumericUpDown24
        Weights(24) = NumericUpDown25 : Weights(25) = NumericUpDown26 : Weights(26) = NumericUpDown27
        Weights(27) = NumericUpDown28 : Weights(28) = NumericUpDown29 : Weights(29) = NumericUpDown30

        ' displaying weight percentage as label text
        WeightsDisplay(0) = Label31 : WeightsDisplay(1) = Label32 : WeightsDisplay(2) = Label33
        WeightsDisplay(3) = Label34 : WeightsDisplay(4) = Label35 : WeightsDisplay(5) = Label36
        WeightsDisplay(6) = Label37 : WeightsDisplay(7) = Label38 : WeightsDisplay(8) = Label39
        WeightsDisplay(9) = Label40 : WeightsDisplay(10) = Label41 : WeightsDisplay(11) = Label42
        WeightsDisplay(12) = Label43 : WeightsDisplay(13) = Label44 : WeightsDisplay(14) = Label45
        WeightsDisplay(15) = Label46 : WeightsDisplay(16) = Label47 : WeightsDisplay(17) = Label48
        WeightsDisplay(18) = Label49 : WeightsDisplay(19) = Label50 : WeightsDisplay(20) = Label51
        WeightsDisplay(21) = Label52 : WeightsDisplay(22) = Label53 : WeightsDisplay(23) = Label54
        WeightsDisplay(24) = Label55 : WeightsDisplay(25) = Label56 : WeightsDisplay(26) = Label57
        WeightsDisplay(27) = Label58 : WeightsDisplay(28) = Label59 : WeightsDisplay(29) = Label60

        ' assign group combobox control objects
        Groups(0) = ComboBox1 : Groups(1) = ComboBox2 : Groups(2) = ComboBox3
        Groups(3) = ComboBox4 : Groups(4) = ComboBox5 : Groups(5) = ComboBox6
        Groups(6) = ComboBox7 : Groups(7) = ComboBox8 : Groups(8) = ComboBox9
        Groups(9) = ComboBox10 : Groups(10) = ComboBox11 : Groups(11) = ComboBox12
        Groups(12) = ComboBox13 : Groups(13) = ComboBox14 : Groups(14) = ComboBox15
        Groups(15) = ComboBox16 : Groups(16) = ComboBox17 : Groups(17) = ComboBox18
        Groups(18) = ComboBox19 : Groups(19) = ComboBox20 : Groups(20) = ComboBox21
        Groups(21) = ComboBox22 : Groups(22) = ComboBox23 : Groups(23) = ComboBox24
        Groups(24) = ComboBox25 : Groups(25) = ComboBox26 : Groups(26) = ComboBox27
        Groups(27) = ComboBox28 : Groups(28) = ComboBox29 : Groups(29) = ComboBox30

        For i As Integer = 0 To 29
            If Not Directories(i).Text = "" Then ' if path field not blank
                Weights(i).Enabled = True ' enable respective weights field
                Groups(i).Enabled = True
            End If
        Next

        With My.Settings
            CurrentActiveSave = .ParentFoldersLastActiveSave

            ' ParentFoldersLastSaveUsed determines save used (default is 1)
            If Not (CurrentActiveSave >= 1 AndAlso CurrentActiveSave <= 10) Then ' if out of range
                MsgBox("Out of range save value." & vbNewLine & "Setting to save 1.", MsgBoxStyle.Critical, "Error")
                .ParentFoldersLastActiveSave = 1 ' default to 1
                CurrentActiveSave = 1 ' default to 1
            End If
        End With

        LoadSave(CurrentActiveSave) ' load respective saved control data

        'DisableLinesStartup() ' disable lines that were previously disabled
        'UpdateWeightDisplays() ' re-compute weight display percentage labels

        If RandomFile.ParentFoldersFormOpenedOnce = False Then
            ' Prevent no textBox data shown when pressing x without doing anything first time opening,
            ' then opening again to find no default data in fields
            UpdateInfo()
        End If

        shownComplete = True
        RandomFile.ParentFoldersFormOpenedOnce = True ' Opened this form once or more flag (used in other form)
    End Sub

    ''' <summary>
    ''' Flag for when shown event has finished.
    ''' </summary>
    Dim shownComplete As Boolean = False
    Private Sub ParentFolders_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown

    End Sub

    Private Sub Browse_Click(sender As Object, e As EventArgs) Handles Browse1.Click
        Browse(TextBox1)
    End Sub

    Private Sub Browse2_Click(sender As Object, e As EventArgs) Handles Browse2.Click
        Browse(TextBox2)
    End Sub

    Private Sub Browse3_Click(sender As Object, e As EventArgs) Handles Browse3.Click
        Browse(TextBox3)
    End Sub

    Private Sub Browse4_Click(sender As Object, e As EventArgs) Handles Browse4.Click
        Browse(TextBox4)
    End Sub

    Private Sub Browse5_Click(sender As Object, e As EventArgs) Handles Browse5.Click
        Browse(TextBox5)
    End Sub

    Private Sub Browse6_Click(sender As Object, e As EventArgs) Handles Browse6.Click
        Browse(TextBox6)
    End Sub

    Private Sub Browse7_Click(sender As Object, e As EventArgs) Handles Browse7.Click
        Browse(TextBox7)
    End Sub

    Private Sub Browse8_Click(sender As Object, e As EventArgs) Handles Browse8.Click
        Browse(TextBox8)
    End Sub

    Private Sub Browse9_Click(sender As Object, e As EventArgs) Handles Browse9.Click
        Browse(TextBox9)
    End Sub

    Private Sub Browse10_Click(sender As Object, e As EventArgs) Handles Browse10.Click
        Browse(TextBox10)
    End Sub

    Private Sub Browse11_Click(sender As Object, e As EventArgs) Handles Browse11.Click
        Browse(TextBox11)
    End Sub

    Private Sub Browse12_Click(sender As Object, e As EventArgs) Handles Browse12.Click
        Browse(TextBox12)
    End Sub

    Private Sub Browse13_Click(sender As Object, e As EventArgs) Handles Browse13.Click
        Browse(TextBox13)
    End Sub

    Private Sub Browse14_Click(sender As Object, e As EventArgs) Handles Browse14.Click
        Browse(TextBox14)
    End Sub

    Private Sub Browse15_Click(sender As Object, e As EventArgs) Handles Browse15.Click
        Browse(TextBox15)
    End Sub

    Private Sub Browse16_Click(sender As Object, e As EventArgs) Handles Browse16.Click
        Browse(TextBox16)
    End Sub

    Private Sub Browse17_Click(sender As Object, e As EventArgs) Handles Browse17.Click
        Browse(TextBox17)
    End Sub

    Private Sub Browse18_Click(sender As Object, e As EventArgs) Handles Browse18.Click
        Browse(TextBox18)
    End Sub

    Private Sub Browse19_Click(sender As Object, e As EventArgs) Handles Browse19.Click
        Browse(TextBox19)
    End Sub

    Private Sub Browse20_Click(sender As Object, e As EventArgs) Handles Browse20.Click
        Browse(TextBox20)
    End Sub

    Private Sub Browse21_Click(sender As Object, e As EventArgs) Handles Browse21.Click
        Browse(TextBox21)
    End Sub

    Private Sub Browse22_Click(sender As Object, e As EventArgs) Handles Browse22.Click
        Browse(TextBox22)
    End Sub

    Private Sub Browse23_Click(sender As Object, e As EventArgs) Handles Browse23.Click
        Browse(TextBox23)
    End Sub

    Private Sub Browse24_Click(sender As Object, e As EventArgs) Handles Browse24.Click
        Browse(TextBox24)
    End Sub

    Private Sub Browse25_Click(sender As Object, e As EventArgs) Handles Browse25.Click
        Browse(TextBox25)
    End Sub

    Private Sub Browse26_Click(sender As Object, e As EventArgs) Handles Browse26.Click
        Browse(TextBox26)
    End Sub

    Private Sub Browse27_Click(sender As Object, e As EventArgs) Handles Browse27.Click
        Browse(TextBox27)
    End Sub

    Private Sub Browse28_Click(sender As Object, e As EventArgs) Handles Browse28.Click
        Browse(TextBox28)
    End Sub

    Private Sub Browse29_Click(sender As Object, e As EventArgs) Handles Browse29.Click
        Browse(TextBox29)
    End Sub

    Private Sub Browse30_Click(sender As Object, e As EventArgs) Handles Browse30.Click
        Browse(TextBox30)
    End Sub

    Private Sub Browse(obj As Object)
        If Not (FolderBrowserDialog1.ShowDialog() = DialogResult.Cancel Or DialogResult.Abort) Then
            If FolderBrowserDialog1.SelectedPath <> "" Then
                obj.Text = FolderBrowserDialog1.SelectedPath
                UpdateWeightDisplays() ' update displayed weight percentages
            End If
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        DisableLine(1, TextBox1, NumericUpDown1)
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        DisableLine(2, TextBox2, NumericUpDown2)
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        DisableLine(3, TextBox3, NumericUpDown3)
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        DisableLine(4, TextBox4, NumericUpDown4)
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        DisableLine(5, TextBox5, NumericUpDown5)
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        DisableLine(6, TextBox6, NumericUpDown6)
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        DisableLine(7, TextBox7, NumericUpDown7)
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        DisableLine(8, TextBox8, NumericUpDown8)
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        DisableLine(9, TextBox9, NumericUpDown9)
    End Sub

    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        DisableLine(10, TextBox10, NumericUpDown10)
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        DisableLine(11, TextBox11, NumericUpDown11)
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        DisableLine(12, TextBox12, NumericUpDown12)
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        DisableLine(13, TextBox13, NumericUpDown13)
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        DisableLine(14, TextBox14, NumericUpDown14)
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        DisableLine(15, TextBox15, NumericUpDown15)
    End Sub

    Private Sub Label16_Click(sender As Object, e As EventArgs) Handles Label16.Click
        DisableLine(16, TextBox16, NumericUpDown16)
    End Sub

    Private Sub Label17_Click(sender As Object, e As EventArgs) Handles Label17.Click
        DisableLine(17, TextBox17, NumericUpDown17)
    End Sub

    Private Sub Label18_Click(sender As Object, e As EventArgs) Handles Label18.Click
        DisableLine(18, TextBox18, NumericUpDown18)
    End Sub

    Private Sub Label19_Click(sender As Object, e As EventArgs) Handles Label19.Click
        DisableLine(19, TextBox19, NumericUpDown19)
    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs) Handles Label20.Click
        DisableLine(20, TextBox20, NumericUpDown20)
    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs) Handles Label21.Click
        DisableLine(21, TextBox21, NumericUpDown21)
    End Sub

    Private Sub Label22_Click(sender As Object, e As EventArgs) Handles Label22.Click
        DisableLine(22, TextBox22, NumericUpDown22)
    End Sub

    Private Sub Label23_Click(sender As Object, e As EventArgs) Handles Label23.Click
        DisableLine(23, TextBox23, NumericUpDown23)
    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
        DisableLine(24, TextBox24, NumericUpDown24)
    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        DisableLine(25, TextBox25, NumericUpDown25)
    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs) Handles Label26.Click
        DisableLine(26, TextBox26, NumericUpDown26)
    End Sub

    Private Sub Label27_Click(sender As Object, e As EventArgs) Handles Label27.Click
        DisableLine(27, TextBox27, NumericUpDown27)
    End Sub

    Private Sub Label28_Click(sender As Object, e As EventArgs) Handles Label28.Click
        DisableLine(28, TextBox28, NumericUpDown28)
    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
        DisableLine(29, TextBox29, NumericUpDown29)
    End Sub

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
        DisableLine(30, TextBox30, NumericUpDown30)
    End Sub

    ''' <summary>
    ''' Disables row when disable button is clicked.
    ''' </summary>
    ''' <param name="line">Number of row.</param>
    ''' <param name="path">TextBox of row.</param>
    ''' <param name="weight">NumericUpDown of row.</param>
    Private Sub DisableLine(line As Integer, path As Object, weight As Object)
        If Not path.Text = "" Then
            Me.DisabledLine(line - 1) = Not Me.DisabledLine(line - 1) ' toggle disabled line flag
            path.enabled = Not path.enabled ' toggle
            weight.enabled = Not weight.enabled ' toggle
            Me.Groups(line - 1).Enabled = Not Me.Groups(line - 1).Enabled

            UpdateWeightDisplays() ' update weight display percentages
            ChangesWereMade = True
        End If
    End Sub

    ''' <summary>
    ''' Used when confirmation button successful. (Updates main form variable)
    ''' </summary>
    Private Sub DisableLineUpdate()
        For i As Integer = 0 To RandomFile.DisabledLine.Count - 1
            RandomFile.DisabledLine(i) = Me.DisabledLine(i)
        Next
    End Sub

    ''' <summary>
    ''' Used when confirmation button successful. (Updates main form group variable)
    ''' </summary>
    Private Sub GroupsUpdate()
        For i As Integer = 0 To RandomFile.FolderGroups.Count - 1
            Dim tmp As String = Me.Groups(i).SelectedItem
            RandomFile.FolderGroups(i) = tmp
        Next
    End Sub

    Private Sub DisableLinesStartup()
        For i As Integer = 0 To Me.Directories.Count - 1
            If DisabledLineSavePtr(CurrentActiveSave).Item(i + 1) = True Then
                Me.Weights(i).Enabled = False
                Me.Directories(i).Enabled = False
                Me.Groups(i).Enabled = False
                Me.DisabledLine(i) = True ' local version for pre-confirmation holding
            Else
                If Me.Directories(i).Text = "" Then
                    Me.Weights(i).Enabled = False
                    Me.Directories(i).Enabled = True
                    Me.Groups(i).Enabled = False
                Else
                    Me.Weights(i).Enabled = True
                    Me.Directories(i).Enabled = True
                    Me.Groups(i).Enabled = True
                End If
                Me.DisabledLine(i) = False ' local version for pre-confirmation holding
            End If
        Next
    End Sub

    Dim InternallyChanging As Boolean = False ' to prevent weight value events to trigger if changed.
    Private Sub WeightNumericUpDown_ValueChanged(sender As Object, e As EventArgs) Handles _
           NumericUpDown1.ValueChanged, NumericUpDown1.KeyUp, NumericUpDown9.ValueChanged, NumericUpDown9.KeyUp, NumericUpDown8.ValueChanged, NumericUpDown8.KeyUp, _
           NumericUpDown7.ValueChanged, NumericUpDown7.KeyUp, NumericUpDown6.ValueChanged, NumericUpDown6.KeyUp, NumericUpDown5.ValueChanged, NumericUpDown5.KeyUp, _
           NumericUpDown4.ValueChanged, NumericUpDown4.KeyUp, NumericUpDown30.ValueChanged, NumericUpDown30.KeyUp, NumericUpDown3.ValueChanged, NumericUpDown3.KeyUp, _
           NumericUpDown29.ValueChanged, NumericUpDown29.KeyUp, NumericUpDown28.ValueChanged, NumericUpDown28.KeyUp, NumericUpDown27.ValueChanged, NumericUpDown27.KeyUp, _
           NumericUpDown26.ValueChanged, NumericUpDown26.KeyUp, NumericUpDown25.ValueChanged, NumericUpDown25.KeyUp, NumericUpDown24.ValueChanged, NumericUpDown24.KeyUp, _
           NumericUpDown23.ValueChanged, NumericUpDown23.KeyUp, NumericUpDown22.ValueChanged, NumericUpDown22.KeyUp, NumericUpDown21.ValueChanged, NumericUpDown21.KeyUp, _
           NumericUpDown20.ValueChanged, NumericUpDown20.KeyUp, NumericUpDown2.ValueChanged, NumericUpDown2.KeyUp, NumericUpDown19.ValueChanged, NumericUpDown19.KeyUp, _
           NumericUpDown18.ValueChanged, NumericUpDown18.KeyUp, NumericUpDown17.ValueChanged, NumericUpDown17.KeyUp, NumericUpDown16.ValueChanged, NumericUpDown16.KeyUp, _
           NumericUpDown15.ValueChanged, NumericUpDown15.KeyUp, NumericUpDown14.ValueChanged, NumericUpDown14.KeyUp, NumericUpDown13.ValueChanged, NumericUpDown13.KeyUp, _
           NumericUpDown12.ValueChanged, NumericUpDown12.KeyUp, NumericUpDown11.ValueChanged, NumericUpDown11.KeyUp, NumericUpDown10.ValueChanged, NumericUpDown10.KeyUp

        If shownComplete = True AndAlso InternallyChanging = False Then
            UpdateWeightDisplays()
            ChangesWereMade = True
        End If
    End Sub

    ''' <summary>
    ''' Update displayed weight text percentages.
    ''' </summary>
    Private Sub UpdateWeightDisplays(Optional GroupBoxChanged As Boolean = False)

        ' determine if groups are being used
        Dim groupsInUse As Boolean = False
        For i As Integer = 0 To Directories.Count - 1
            Dim tmp As String = Groups(i).SelectedItem
            If Not tmp = "" Then ' not the default blank ("") option (index 0)
                groupsInUse = True ' groups are being used
                Exit For
            End If
        Next

        If groupsInUse = False Then
            ' determine sum of all weights
            Dim total As Integer = 0
            For i As Integer = 0 To Me.Directories.Count - 1 ' find sum of all weights
                If Not (Me.Directories(i).Text = "" OrElse Me.DisabledLine(i) = True) Then
                    total += Weights(i).Value
                End If
            Next

            ' update weight text display percentages
            For i As Integer = 0 To WeightsDisplay.Count - 1
                If Not (Weights(i).Value = 0 OrElse Me.DisabledLine(i) = True OrElse total = 0 OrElse Weights(i).Enabled = False) Then
                    WeightsDisplay(i).Text = CStr(Math.Round((Weights(i).Value / total) * 100, 1)) & "%"
                Else
                    WeightsDisplay(i).Text = "0%"
                End If
            Next
        Else ' groups are being used
            InternallyChanging = True ' prevent value change event calls during this code block

            Dim total As Integer = 0 ' holds sum of all corresponding group weights

            Dim tempGroupInfo As New ArrayList ' stores structs GroupWeightDisplay
            Dim skipRows(Me.Directories.Count - 1) As Boolean
            For i As Integer = 0 To Me.Directories.Count - 1
                skipRows(i) = False ' initialize all to false
            Next

            Dim flag1 As Boolean = False
            For y As Integer = 0 To Me.Directories.Count - 1
                Dim tmp1 As String = Groups(y).SelectedItem
                If skipRows(y) = False AndAlso Not tmp1 = "" Then
                    Dim currentGroupTotal As Integer = 0 ' temporary total for current group
                    Dim currentGroup(Me.Directories.Count - 1) As Boolean
                    If Not currentGroup(0) = False Then
                        For i As Integer = 0 To Me.Directories.Count - 1
                            currentGroup(i) = False ' initialize all to false
                        Next
                    End If

                    For i As Integer = 0 To Me.Directories.Count - 1 ' find sum of all current group weights
                        Dim tmp2 As String = Groups(i).SelectedItem
                        If skipRows(i) = False AndAlso tmp2 = tmp1 Then ' tally weights only for certain group
                            skipRows(i) = True
                            If Not Me.Directories(i).Text = "" AndAlso Not Me.DisabledLine(i) = True Then
                                currentGroupTotal += Weights(i).Value
                                total += Weights(i).Value
                                flag1 = True
                            End If
                            currentGroup(i) = True
                        End If
                    Next

                    ' update current group counters (later, but store for now)
                    If (flag1) Then
                        Dim newstruct As New GroupWeightDisplay
                        'newstruct.currentGroup = New ArrayList
                        newstruct.currentGroup = currentGroup
                        newstruct.currentGroupTotal = currentGroupTotal
                        tempGroupInfo.Add(newstruct)
                    End If

                ElseIf skipRows(y) = False Then ' not a group ("" value)
                    If Not Me.Directories(y).Text = "" AndAlso Not Me.DisabledLine(y) = True Then
                        total += Weights(y).Value
                    End If
                End If
            Next

            ' groups are in use so up those only right now
            If groupsInUse = True AndAlso Not IsNothing(tempGroupInfo) Then
                For k As Integer = 0 To tempGroupInfo.Count - 1
                    For i As Integer = 0 To Me.Directories.Count - 1
                        Dim temp As Object = tempGroupInfo(k).currentGroup
                        If temp(i) = True Then ' change current group weights only
                            'temp(i) = False
                            If Me.DisabledLine(i) = False AndAlso Not Me.Directories(i).Text = "" AndAlso Not tempGroupInfo(k).currentGroupTotal = 0 AndAlso _
                                Not tempGroupInfo(k).currentGroupTotal = 0 AndAlso Weights(i).Enabled = True Then
                                WeightsDisplay(i).Text = CStr(Math.Round((CDec(tempGroupInfo(k).currentGroupTotal) / total) * 100, 1)) & "%"
                            Else
                                WeightsDisplay(i).Text = "0%"
                            End If
                        End If
                    Next
                Next
            End If

            ' update weight text display percentages
            For i As Integer = 0 To WeightsDisplay.Count - 1
                Dim tmp As String = Groups(i).SelectedItem
                If tmp = "" Then
                    If Not Weights(i).Value = 0 AndAlso Me.DisabledLine(i) = False AndAlso Not total = 0 AndAlso Weights(i).Enabled = True Then
                        WeightsDisplay(i).Text = CStr(Math.Round((Weights(i).Value / total) * 100, 1)) & "%"
                    Else
                        WeightsDisplay(i).Text = "0%"
                    End If
                End If
            Next

            For i As Integer = 0 To Me.Directories.Count - 1
                If Me.DisabledLine(i) = True Then
                    WeightsDisplay(i).Text = "0%"
                End If
            Next

            InternallyChanging = False ' unset
            End If
    End Sub

    Structure GroupWeightDisplay
        Dim currentGroup As Object ' boolean array of row size
        Dim currentGroupTotal As Integer
    End Structure

    ''' <summary>
    ''' Determine the weight total for each group in use and return an ArrayList of GroupWeightDisplay structs (of each unique group)
    ''' </summary>
    Private Function GetGroupWeights() As ArrayList
        Dim total As Integer = 0 ' holds sum of all corresponding group weights

        Dim tempGroupInfo As New ArrayList ' stores structs GroupWeightDisplay
        Dim skipRows(Me.Directories.Count - 1) As Boolean
        For i As Integer = 0 To Me.Directories.Count - 1
            skipRows(i) = False ' initialize all to false
        Next

        Dim flag1 As Boolean = False
        For y As Integer = 0 To Me.Directories.Count - 1
            Dim tmp1 As String = Groups(y).SelectedItem
            If skipRows(y) = False AndAlso Not tmp1 = "" Then
                ' initialization
                Dim currentGroupTotal As Integer = 0 ' temporary total for current group
                Dim currentGroup(Me.Directories.Count - 1) As Boolean
                If Not currentGroup(0) = False Then
                    For i As Integer = 0 To Me.Directories.Count - 1
                        currentGroup(i) = False ' initialize all to false
                    Next
                End If

                ' find sum of all current group weights
                For i As Integer = 0 To Me.Directories.Count - 1
                    Dim tmp2 As String = Groups(i).SelectedItem
                    If skipRows(i) = False AndAlso tmp2 = tmp1 Then ' tally weights only for certain group
                        skipRows(i) = True
                        If Not Me.Directories(i).Text = "" AndAlso Not Me.DisabledLine(i) = True Then
                            currentGroupTotal += Weights(i).Value
                            total += Weights(i).Value
                            flag1 = True
                        End If
                        currentGroup(i) = True
                    End If
                Next

                ' store current group weight info
                If (flag1) Then
                    Dim newstruct As New GroupWeightDisplay
                    'newstruct.currentGroup = New ArrayList
                    newstruct.currentGroup = currentGroup
                    newstruct.currentGroupTotal = currentGroupTotal
                    tempGroupInfo.Add(newstruct)
                End If

                'ElseIf skipRows(y) = False Then ' not a group ("" value)
                '    If Not Me.Directories(y).Text = "" AndAlso Not Me.DisabledLine(y) = True Then
                '       total += Weights(y).Value
                '    End If
            End If
        Next

        Return tempGroupInfo
    End Function

    ''' <summary>
    ''' Load the save from My.Settings and change the visible controls accordingly. (Saves are numbered from 1 and above)
    ''' </summary>
    ''' <param name="saveNum">Save number to be loaded.</param>
    Private Sub LoadSave(saveNum As Integer)
        For i As Integer = 0 To 29
            Me.Directories(i).Text = DirectoriesSavePtr(saveNum).Item(i + 1)
            Me.Weights(i).Text = WeightsSavePtr(saveNum).Item(i + 1)
            Me.DisabledLine(i) = CBool(DisabledLineSavePtr(saveNum).Item(i + 1))
            If Me.Groups(i).Items.Contains(GroupsSavePtr(saveNum).Item(i + 1)) Then
                Me.Groups(i).SelectedItem = GroupsSavePtr(saveNum).Item(i + 1)
                If DisabledLine(i) = False AndAlso Not Me.Directories(i).Text = "" Then
                    Me.Groups(i).Enabled = True
                End If
            Else
                Me.Groups(i).SelectedItem = "" ' default is no group
            End If
        Next
        DisableLinesStartup() ' disable the proper lines
        UpdateWeightDisplays()
        ActiveSaveLabel.Text = CStr(CurrentActiveSave)
    End Sub

    ''' <summary>
    ''' Save the active save button to My.Settings. (Numbered from 1 and above)
    ''' </summary>
    Private Sub SaveActiveSave()
        Dim x = CurrentActiveSave
        For i As Integer = 0 To 29
            DirectoriesSavePtr(x).Item(i + 1) = Directories(i).Text
            WeightsSavePtr(x).Item(i + 1) = Weights(i).Text
            If CStr(DisabledLine(i)) = "False" Then
                DisabledLineSavePtr(x).Item(i + 1) = "0"
            ElseIf CStr(DisabledLine(i)) = "True" Then
                DisabledLineSavePtr(x).Item(i + 1) = "1"
            End If
            GroupsSavePtr(x).item(i + 1) = CStr(Groups(i).SelectedItem)
        Next
    End Sub

    Private Sub SaveButtonHandler(ClickedSave As Integer)
        If Not CurrentActiveSave = ClickedSave Then
            If ChangesWereMade = True Then
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Exclamation)
                Dim result As MsgBoxResult = MsgBox("Would you like to save any changes before loading a new save?", MsgBoxStyle.YesNoCancel, "Warning")
                If result = MsgBoxResult.Yes Then
                    SaveActiveSave() ' save currently visible settings to proper settings

                    CurrentActiveSave = ClickedSave ' change to new save value
                    LoadSave(ClickedSave) ' load the newly selected save
                    ChangesWereMade = False ' reset
                ElseIf result = MsgBoxResult.No Then ' switch save without saving
SaveChangesNoOption:
                    CurrentActiveSave = ClickedSave ' change to new save value
                    LoadSave(ClickedSave) ' load the newly selected save
                    ChangesWereMade = False ' reset
                End If
            Else
                GoTo SaveChangesNoOption
            End If
        Else ' clicked on same save button as currently active one
            If ChangesWereMade = True Then
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Exclamation)
                Dim result As MsgBoxResult = MsgBox("Overwrite current save?", MsgBoxStyle.YesNo, "Warning")
                If result = MsgBoxResult.Yes Then
                    SaveActiveSave() ' save currently visible settings to proper settings
                    ChangesWereMade = False ' reset
                End If
            End If
        End If
    End Sub

    Private Sub SaveButton1_Click(sender As Object, e As EventArgs) Handles SaveButton1.Click
        SaveButtonHandler(1)
    End Sub

    Private Sub SaveButton2_Click(sender As Object, e As EventArgs) Handles SaveButton2.Click
        SaveButtonHandler(2)
    End Sub

    Private Sub SaveButton3_Click(sender As Object, e As EventArgs) Handles SaveButton3.Click
        SaveButtonHandler(3)
    End Sub

    Private Sub SaveButton4_Click(sender As Object, e As EventArgs) Handles SaveButton4.Click
        SaveButtonHandler(4)
    End Sub

    Private Sub SaveButton5_Click(sender As Object, e As EventArgs) Handles SaveButton5.Click
        SaveButtonHandler(5)
    End Sub

    Private Sub SaveButton6_Click(sender As Object, e As EventArgs) Handles SaveButton6.Click
        SaveButtonHandler(6)
    End Sub

    Private Sub SaveButton7_Click(sender As Object, e As EventArgs) Handles SaveButton7.Click
        SaveButtonHandler(7)
    End Sub

    Private Sub SaveButton8_Click(sender As Object, e As EventArgs) Handles SaveButton8.Click
        SaveButtonHandler(8)
    End Sub

    Private Sub SaveButton9_Click(sender As Object, e As EventArgs) Handles SaveButton9.Click
        SaveButtonHandler(9)
    End Sub

    Private Sub SaveButton10_Click(sender As Object, e As EventArgs) Handles SaveButton10.Click
        SaveButtonHandler(10)
    End Sub

    Private Sub DirectoryGroupComboBoxValueChanged(sender As Object, e As EventArgs) Handles _
           ComboBox1.SelectedIndexChanged, ComboBox9.SelectedIndexChanged, ComboBox8.SelectedIndexChanged, ComboBox7.SelectedIndexChanged, _
           ComboBox6.SelectedIndexChanged, ComboBox5.SelectedIndexChanged, ComboBox4.SelectedIndexChanged, ComboBox30.SelectedIndexChanged, _
           ComboBox3.SelectedIndexChanged, ComboBox29.SelectedIndexChanged, ComboBox28.SelectedIndexChanged, ComboBox27.SelectedIndexChanged, _
           ComboBox26.SelectedIndexChanged, ComboBox25.SelectedIndexChanged, ComboBox24.SelectedIndexChanged, ComboBox23.SelectedIndexChanged, _
           ComboBox22.SelectedIndexChanged, ComboBox21.SelectedIndexChanged, ComboBox20.SelectedIndexChanged, ComboBox2.SelectedIndexChanged, _
           ComboBox19.SelectedIndexChanged, ComboBox18.SelectedIndexChanged, ComboBox17.SelectedIndexChanged, ComboBox16.SelectedIndexChanged, _
           ComboBox15.SelectedIndexChanged, ComboBox14.SelectedIndexChanged, ComboBox13.SelectedIndexChanged, ComboBox12.SelectedIndexChanged, _
           ComboBox11.SelectedIndexChanged, ComboBox10.SelectedIndexChanged

        If shownComplete = True Then
            UpdateWeightDisplays(True)
            ChangesWereMade = True
        End If
    End Sub

    ''' <summary>
    ''' Drag and drop folders onto the directory textboxes to replace the current text with the folder path(s). Works with multiple dragged folders.
    ''' </summary>
    Private Sub DirectoryTextBox_DragDrop(sender As Object, e As DragEventArgs) Handles _
        TextBox1.DragDrop, TextBox9.DragDrop, TextBox8.DragDrop, TextBox7.DragDrop, TextBox6.DragDrop, TextBox5.DragDrop, _
        TextBox4.DragDrop, TextBox30.DragDrop, TextBox3.DragDrop, TextBox29.DragDrop, TextBox28.DragDrop, TextBox27.DragDrop, _
        TextBox26.DragDrop, TextBox25.DragDrop, TextBox24.DragDrop, TextBox23.DragDrop, TextBox22.DragDrop, TextBox21.DragDrop, _
        TextBox20.DragDrop, TextBox2.DragDrop, TextBox19.DragDrop, TextBox18.DragDrop, TextBox17.DragDrop, TextBox16.DragDrop, _
        TextBox15.DragDrop, TextBox14.DragDrop, TextBox13.DragDrop, TextBox12.DragDrop, TextBox11.DragDrop, TextBox10.DragDrop

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim directoryStrings As String() = e.Data.GetData(DataFormats.FileDrop)
            Dim index As Integer, Stringsindex As Integer = 0

            For i As Integer = 0 To Directories.Count - 1
                If sender.Equals(Directories(i)) Then
                    index = i
                End If
            Next

            For DirectoryIndex As Integer = index To Directories.Count - 1
                If DisabledLine(DirectoryIndex) = False Then
22352352356:
                    If Stringsindex <= directoryStrings.Length - 1 Then
                        If IO.Directory.Exists(directoryStrings(Stringsindex)) Then
                            Directories(DirectoryIndex).Text = directoryStrings(Stringsindex)
                        Else
                            Stringsindex += 1
                            GoTo 22352352356
                        End If
                        Stringsindex += 1
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Related to dragging of folders into directoy textboxes.
    ''' </summary>
    Private Sub DirectoryTextBox_DragEnter(sender As Object, e As DragEventArgs) Handles _
        TextBox1.DragEnter, TextBox9.DragEnter, TextBox8.DragEnter, TextBox7.DragEnter, TextBox6.DragEnter, TextBox5.DragEnter, _
        TextBox4.DragEnter, TextBox30.DragEnter, TextBox3.DragEnter, TextBox29.DragEnter, TextBox28.DragEnter, TextBox27.DragEnter, _
        TextBox26.DragEnter, TextBox25.DragEnter, TextBox24.DragEnter, TextBox23.DragEnter, TextBox22.DragEnter, TextBox21.DragEnter, _
        TextBox20.DragEnter, TextBox2.DragEnter, TextBox19.DragEnter, TextBox18.DragEnter, TextBox17.DragEnter, TextBox16.DragEnter, _
        TextBox15.DragEnter, TextBox14.DragEnter, TextBox13.DragEnter, TextBox12.DragEnter, TextBox11.DragEnter, TextBox10.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub
End Class