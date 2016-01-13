Imports System.IO

Public Class RandomFile
    ''' <summary>
    ''' Used to prevent multiple clicks, not necessary without threading.
    ''' </summary>
    Dim locked As Boolean = False
    ''' <summary>
    ''' Flag determining if shortcut output folder is specified. (For when shortcut output is enabled)
    ''' </summary>
    Public SpecifiedShortcutFolder As Boolean = False
    ''' <summary>
    ''' Flag for when multiple folders option in use.
    ''' </summary>
    Public MultipleFoldersFlag As Boolean = True
    ''' <summary>
    ''' One time handling has or has not been performed flag.
    ''' </summary>
    Public MultipleFoldersHandled As Boolean = False
    ''' <summary>
    ''' Flag determining whether or not shortcuts will be created on roll.
    ''' </summary>
    Public shortcutCreationAllowed As Boolean = True
    ''' <summary>
    ''' Used to move controls when a legendary roll has occured, but only once. (Consecutive leg. roll won't move further)
    ''' </summary>
    Dim LegendaryMoved As Boolean = False
    ''' <summary>
    ''' Directory for shortcuts to be placed in if enabled.
    ''' </summary>
    Dim creationDir As String
    ''' <summary>
    ''' Contains total weight, used in multiple folders option.
    ''' </summary>
    Public TotalWeight As Integer = 0
    ''' <summary>
    ''' Flag for updating parent folders form once on the first opening.
    ''' </summary>
    Public ParentFoldersFormOpenedOnce As Boolean = False
    ''' <summary>
    ''' Contains path name of parent folders.
    ''' </summary>
    Public FolderDirectories(29) As String
    ''' <summary>
    '''Contains weights for respective parent folders.
    ''' </summary>
    Public FolderWeights(29) As Integer
    ''' <summary>
    ''' Contains flag for whether or not a parent folder is disabled.
    ''' </summary>
    Public DisabledLine(29) As Boolean
    ''' <summary>
    ''' Contains group value of each parent folder
    ''' </summary>
    Public FolderGroups(29) As String
    ''' <summary>
    ''' Default times to repeat main button click is once per click. (Used in automation)
    ''' </summary>
    Public Duration As Integer = 1
    ''' <summary>
    ''' Flag allowing folders to be searched/rolled for in directories.
    ''' </summary>
    Public FolderFlag As Boolean = False
    ''' <summary>
    ''' Flag allowing Files to be searched/rolled for in directories.
    ''' </summary>
    Public FileFlag As Boolean = True
    ''' <summary>
    ''' ArrayList of ArrayLists of all files corresponding to parent folders. (When using multiple folders option)
    ''' </summary>
    Public DirectoryFileTree(29) As ArrayList
    ''' <summary>
    ''' ArrayList of ArrayList of all folders corresponding to parent folders. (When using multiple folders option)
    ''' </summary>
    Public DirectoryFolderTree(29) As ArrayList
    ''' <summary>
    ''' Contains weights for each group in use as structs.
    ''' </summary>
    Public GroupWeightInfo As ArrayList
    'Structure GroupWeightDisplay ' structure layout
    '   Dim currentGroup As Object ' boolean array of row size
    '   Dim currentGroupTotal As Integer
    'End Structure
    ''' <summary>
    ''' Stores file results for single parent folder option.
    ''' </summary>
    Public TempFileTree As ArrayList
    ''' <summary>
    ''' Stores folder results for single parent folder option.
    ''' </summary>
    Public TempFolderTree As ArrayList
    ''' <summary>
    ''' Flag that roll once list is in use.
    ''' </summary>
    Public RollOnceListFlag As Boolean = False
    ''' <summary>
    ''' List used for preventing a file/folder from being rolled more than once when it has been already rolled.
    ''' </summary>
    Public RollOnceList As ArrayList
    ''' <summary>
    ''' Flag used for re-computing trees.
    ''' </summary>
    Public RecomputeTreesChangesWereMade As Boolean = False

    Public Sub RollButton_Click(sender As Object, e As EventArgs) Handles RollButton.Click
        If (Not locked) Then
            locked = True ' Lock

            If (SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = False Then
                MainWork()
            ElseIf (SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = True Then
                MainWork()
            Else
                If (SpecifiedShortcutFolder = False AndAlso shortcutCreationAllowed = True) Then
                    OutputFolderConfig() ' attempt to specify output folder
                    If SpecifiedShortcutFolder = True AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = False Then
                        GoTo Setting2
                    End If
                ElseIf (SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = False Then
Setting2:
                    FoldersExpandButton.PerformClick() ' MultipleFoldersHandled is set to True only when confirm button is successful in other form just opened
                End If
            End If

            locked = False ' Unlock
        End If
    End Sub

    ''' <summary>
    ''' Handles main work related to rolling.
    ''' </summary>
    Private Sub MainWork()
        If RecomputeTreesChangesWereMade = True Then
            If ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = False) Then
                ParentFolderSingleFolderChanged = True ' recompute for single folder option
            ElseIf ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = True) Then
                UpdateTrees()
            End If
            If Search.Visible = True Then
                Search.RefreshButton_Click()
            End If
            RecomputeTreesChangesWereMade = False ' done recomputing
        End If

        For i As Integer = 1 To Duration

            If MultipleFoldersFlag = True Then ' multiple folders option activated
                DetermineParentFolder()
            Else ' single folder option/mode
                If Not System.IO.Directory.Exists(TextBox1.Text) Then ' error in path given
                    BrowseFolderButton.PerformClick()
                    If Not Directory.Exists(TextBox1.Text) Then
                        MsgBox("Please enter a valid directory and try again.", MsgBoxStyle.Exclamation, "Error")
                        Exit Sub
                    Else
                        ParentFolderSingleFolderChanged = True
                        GoTo SingleFolderOption1
                    End If
                Else
SingleFolderOption1:
                    My.Settings.SingleParentFolder = TextBox1.Text ' update setting

                    If ParentFolderSingleFolderChanged = True Then ' re-compute file/folder tree
                        Dim tmp As New ArrayList
                        tmp = FileTree.CreateFileList(TextBox1.Text)
                        TempFileTree = tmp.Clone()
                        tmp = FileTree.CreateFolderList(TextBox1.Text)
                        TempFolderTree = tmp.Clone()
                        ParentFolderSingleFolderChanged = False

                        If Search.Visible = True Then
                            Search.RefreshButton_Click()
                        End If
                    End If
                End If
            End If

            Dim ResultFolder As String
            If ParentFolderChosen = True Then
                ResultFolder = GetRandomSubFolder() ' DetermineParentFolder() gave index for this method via currentParentFolderIndex
            Else ' parent folder was not chosen
                ResultFolder = ""
            End If
            ResultBox.Text = ResultFolder

            Static nothingCount As Integer = 1
            If ResultFolder = "" OrElse ParentFolderChosen = False Then ' path does not exist or parent folder not determined
                If FileFlag = True AndAlso FolderFlag = False Then
                    'MsgBox("No file was found.", MsgBoxStyle.Information, "")
                ElseIf FileFlag = False AndAlso FolderFlag = True Then
                    'MsgBox("No folder was found.", MsgBoxStyle.Information, "")
                ElseIf FileFlag = True AndAlso FolderFlag = True Then
                    'MsgBox("No file or folder was found.", MsgBoxStyle.Information, "")
                End If

                If RollOnceListFlag = True Then
                    IgnoreRollButton.Visible = True
                    ClearRollOnceListButton.Visible = True
                End If
                IgnoreRollButton.Enabled = False

                ResultBox.Text = "Nothing Found (" & nothingCount & " times)"
                nothingCount += 1 ' increment count
                Continue For
            Else
                If Not IsNothing(My.Settings.FavoriteRollsList) AndAlso My.Settings.FavoriteRollsList.Contains(ResultBox.Text) Then ' if rolled item is a favorite
                    FavoriteRolled()
                    IgnoreRollButton.Enabled = False
                    ' favorite rolls are excluded from roll once list
                Else
                    NonFavoriteRolled()

                    ' nonfavorite so add to roll once list if enabled
                    If RollOnceListFlag = True Then
                        RollOnceListHandling("Add") ' add to exlusion list and remove from trees
                    End If
                End If
                nothingCount = 1 ' reset count
            End If

            Dim myStr() As String = Split(ResultFolder, "\")
            Dim shortcutName As String = myStr(myStr.Count - 1) ' Put last element into variable

            RollsRichTextBox.Clear()
            TimesBox.Text = Roll_Dice() ' Execute roll dice function
            TopRollsList() ' Update top rolls list

            DetermineShortcuts(shortcutName, creationDir) ' Create corresponding amount of shortcuts

            If LegendaryFlag = True Then ' Show or hide legendary label text
                LegendaryLabel.Visible = True
                LegendaryFlag = False
                If LegendaryMoved = False Then
                    Label3.Location = New Point(Label3.Location.X, Label3.Location.Y + 53) ' move to make room for legendary text
                    LegendaryMoved = True
                End If
                If LegendaryDelay = True Then
                    RollButton.Enabled = False
                    SleepTimer.Start()
                End If
            ElseIf LegendaryLabel.Visible = True Then
                LegendaryLabel.Visible = False
                If LegendaryMoved = True Then
                    Label3.Location = New Point(Label3.Location.X, Label3.Location.Y - 53) ' move back to normal
                    LegendaryMoved = False
                End If
            End If
        Next

        If AutomateFlag = True Then
            AutomateFlag = False
            If RollOnceAutomateFlag = True Then
                RollOnceAutomateFlag = False
                RollOnceListFlag = True ' set back to true again now that automation is over
            End If
            Duration = original ' Set duration back to normal
            My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
            MsgBox("Automation Completed Successfully", MsgBoxStyle.OkOnly, "")
        Else
            If YesToolStripMenuItem3.Checked = True Then ' roll and execute checked
                OpenButton.PerformClick() ' rolled, now execute file
            End If
        End If
    End Sub

    ''' <summary>
    ''' Configures the output folder for shortcut creation.
    ''' </summary>
    Public Sub OutputFolderConfig() ' (public because of admin)
        If Not My.Settings.CreateShortcuts = False Then
            creationDir = InputBox("Enter directory for shortcuts to be placed in.", "Enter directory", My.Settings.OutputFolder)
            If Not creationDir = "" Then ' not aborted or canceled
                If Not System.IO.Directory.Exists(creationDir) Then
                    My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
                    Dim result As MsgBoxResult = MsgBox("Directory entered does not exist." & vbNewLine & "Attempt to create directory?", MsgBoxStyle.OkCancel, "Error")
                    If result = MsgBoxResult.Ok Then
                        Try
                            My.Computer.FileSystem.CreateDirectory(creationDir)
                            MsgBox("Directory created successfully.", MsgBoxStyle.Information)
                            GoTo OutputFolderSuccess
                        Catch ex As Exception
                            MsgBox("Error creating directory." & vbNewLine & ex.Message, MsgBoxStyle.Critical, "Error")
                        End Try
                    End If
                Else
OutputFolderSuccess:
                    My.Settings.OutputFolder = creationDir
                    SpecifiedShortcutFolder = True ' Folder was just specified
                    OpenOutputFolderButton.Enabled = True
                    ClearOutputFolderButton.Enabled = True
                End If
            End If
        Else ' shortcut creation is disabled
            'SpecifiedShortcutFolder = True ' technically not specified so be careful
        End If
    End Sub

    Private Sub OpenButton_Click(sender As Object, e As EventArgs) Handles OpenButton.Click
        OpenFile(ResultBox.Text)
    End Sub

    Public Sub OpenFile(fullPath As String)
        If fullPath <> "" OrElse fullPath <> "Nothing Found" Then
            Try
                If System.IO.File.Exists(fullPath) Then
                    Process.Start(fullPath)
                ElseIf System.IO.Directory.Exists(fullPath) Then
                    Process.Start(fullPath)
                End If
            Catch ex As Exception
                If ex.Message = "Application not found" Then
                    MsgBox("No program associated with this filetype", MsgBoxStyle.Exclamation, "Error")
                End If
            End Try
        End If
    End Sub

    Private Sub OpenContainingFolderButton_Click(sender As Object, e As EventArgs) Handles OpenContainingFolderButton.Click
        If Not ParentFolderChosen = False AndAlso (File.Exists(ResultBox.Text) OrElse Directory.Exists(ResultBox.Text)) Then
            'Call Shell("explorer /select," & """" & ResultBox.Text & """", AppWinStyle.NormalFocus) ' select in containing folder
            'Process.Start("explorer", "/select," & """" & ResultBox.Text & """") ' select in containing folder
            ContainingButton(ResultBox.Text)
        End If
    End Sub

    ''' <summary>
    ''' Get the containing folder of a given path as string.
    ''' </summary>
    Public Function GetContainingFolder(path As String) As String
        Dim myStr() As String = Split(path, "\")
        Dim containingFolder As String = ""
        For i As Integer = 0 To myStr.Length - 2
            containingFolder = containingFolder & myStr(i) & "\" ' contruct containing folder
        Next
        Return containingFolder
    End Function

    Private Sub OpenOutputFolderButton_Click(sender As Object, e As EventArgs) Handles OpenOutputFolderButton.Click
        If My.Settings.CreateShortcuts = False Then
            MsgBox("Shortcut creation is disabled.", MsgBoxStyle.Information, "Error")
        End If
        If SpecifiedShortcutFolder = True Then
            If System.IO.Directory.Exists(creationDir) Then
                Process.Start("explorer.exe", """" & creationDir & """") ' Open output folder
            Else
                MsgBox("Output folder does not exist!", MsgBoxStyle.Critical, "Error")
                OutputFolderConfig()
            End If
        Else
            OutputFolderConfig() ' specify output folder
        End If
    End Sub

    ''' <summary>
    ''' Deletes ALL shortcut files (*.lnk) in output folder. (Shortcut option related)
    ''' </summary>
    Private Sub ClearOutputFolderButton_Click(sender As Object, e As EventArgs) Handles ClearOutputFolderButton.Click
        If Directory.Exists(creationDir) Then
            ' determine if shortcuts (.lnk) files exist in output folder
            Dim shortcutsFoundFlag As Boolean = False ' flag for at least one shortcut found
            Dim list = System.IO.Directory.GetFiles(creationDir)
            For i As Integer = 0 To list.Count - 1
                Dim myStr() As String = Split(list(i), "\")
                Dim fileName() As String = Split(myStr(myStr.Count - 1), ".") ' Put last element into variable
                If fileName(fileName.Count - 1) = "lnk" Then ' extension type
                    shortcutsFoundFlag = True ' flag for at least one shortcut found
                    Exit For
                End If
            Next

            If Not shortcutsFoundFlag = False Then ' shortcuts found in output folder
                My.Computer.Audio.PlaySystemSound(Media.SystemSounds.Beep)
                Dim result As MsgBoxResult = MsgBox("Are you sure you want to delete ALL shortcut files in the output folder? ", MsgBoxStyle.OkCancel, "Warning")
                If Not (result = DialogResult.Cancel Or DialogResult.Abort) Then
                    For Each foundFile As String In My.Computer.FileSystem.GetFiles(creationDir, _
                        Microsoft.VisualBasic.FileIO.SearchOption.SearchTopLevelOnly, "*.lnk")

                        My.Computer.FileSystem.DeleteFile(foundFile, _
                            Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs, _
                            Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin)
                    Next
                End If
            Else
                MsgBox("No shortcuts exist in the output directory.", MsgBoxStyle.Information, "Error")
            End If
        Else
            MsgBox("Output folder does not exist!", MsgBoxStyle.Critical, "Error")
            OutputFolderConfig()
        End If
    End Sub

    ''' <summary>
    ''' Contains index of current parent folder that had an item rolled from it.
    ''' </summary>
    Dim currentParentFolderIndex As Integer
    ''' <summary>
    ''' Gets random subfolder OR file. (Depending on what is enabled)
    ''' </summary>
    Public Function GetRandomSubFolder() As String ' given path of a parent folder
        ' Static create a Random object so that we do not create a new one each time
        Static R As New Random()

        Dim index As Integer = currentParentFolderIndex

        If MultipleFoldersFlag = True Then
            If (FolderFlag = True AndAlso FileFlag = False) Then ' Folders only
                ' Sanity check
                If DirectoryFolderTree(index).Count = 0 Then
                    If Not FolderGroups(index) = "" Then : TextBox1.Text = "Group " & FolderGroups(index) : Else : TextBox1.Text = FolderDirectories(index) : End If
                    Return FolderDirectories(index) ' return parent folder itself instead
                End If
                ' Get a random number. The second parameter is exclusive so (0,4) will always return 3 as a maximum
                Dim RandomIndex As Integer = R.Next(0, DirectoryFolderTree(index).Count)
                ' Return the path at that index
                If Not FolderGroups(index) = "" Then : TextBox1.Text = "Group " & FolderGroups(index) : Else : TextBox1.Text = FolderDirectories(index) : End If
                Return DirectoryFolderTree(index).Item(RandomIndex)
            ElseIf (FileFlag = True AndAlso FolderFlag = False) Then ' Files only
                If Not (DirectoryFileTree(index).Count = 0) Then ' files exist (not empty)
                    Dim RandomIndex As Integer = R.Next(0, DirectoryFileTree(index).Count)
                    If Not FolderGroups(index) = "" Then : TextBox1.Text = "Group " & FolderGroups(index) : Else : TextBox1.Text = FolderDirectories(index) : End If
                    Return DirectoryFileTree(index).Item(RandomIndex)
                Else ' nothing found at all
                    Return ""
                End If
            ElseIf (FileFlag = True AndAlso FolderFlag = True) Then ' Files and Folders
                Dim RandomIndex As Integer = R.Next(0, DirectoryFileTree(index).Count + DirectoryFolderTree(index).Count)
                If RandomIndex >= DirectoryFileTree(index).Count Then ' return a folder path
                    If Not FolderGroups(index) = "" Then : TextBox1.Text = "Group " & FolderGroups(index) : Else : TextBox1.Text = FolderDirectories(index) : End If
                    Return DirectoryFolderTree(index).Item(RandomIndex - DirectoryFileTree(index).Count)
                Else ' return a file path
                    If Not FolderGroups(index) = "" Then : TextBox1.Text = "Group " & FolderGroups(index) : Else : TextBox1.Text = FolderDirectories(index) : End If
                    Return DirectoryFileTree(index).Item(RandomIndex)
                End If
            End If
        ElseIf MultipleFoldersFlag = False Then
            If IsNothing(TempFileTree) Then ' not initiated
                TempFileTree = New ArrayList
            ElseIf IsNothing(TempFolderTree) Then
                TempFolderTree = New ArrayList
            End If

            If (FolderFlag = True AndAlso FileFlag = False) Then ' Folders only
                If Not TempFolderTree.Count = 0 Then
                    ' Get a random number. The second parameter is exclusive so (0,4) will always return 3 as a maximum
                    Dim RandomIndex As Integer = R.Next(0, TempFolderTree.Count)
                    ' Return the path at that index
                    Return TempFolderTree.Item(RandomIndex)
                Else ' nothing found
                    Return ""
                End If
            ElseIf (FileFlag = True AndAlso FolderFlag = False) Then ' Files only
                If Not TempFileTree.Count = 0 Then ' files exist (not empty)
                    Dim RandomIndex As Integer = R.Next(0, TempFileTree.Count)
                    Return TempFileTree.Item(RandomIndex)
                Else ' nothing found at all
                    Return ""
                End If
            ElseIf (FileFlag = True AndAlso FolderFlag = True) Then ' Files and Folders
                If Not TempFileTree.Count = 0 OrElse Not TempFolderTree.Count = 0 Then ' files/folders exist (not empty)
                    Dim RandomIndex As Integer = R.Next(0, TempFileTree.Count + TempFolderTree.Count)
                    If RandomIndex >= TempFileTree.Count Then ' return a folder path
                        Return TempFolderTree.Item(RandomIndex - TempFileTree.Count)
                    Else ' return a file path
                        Return TempFileTree.Item(RandomIndex)
                    End If
                Else ' nothing found at all
                    Return ""
                End If
            End If
        End If

        MsgBox("Error Code: 421", MsgBoxStyle.Critical)
        Return "Null"
    End Function

    ''' <summary>
    ''' File Or folder no longer exists error message.
    ''' </summary>
    Private Sub ErrorMessage124(str As String)
        MsgBox("File or Folder no longer exists:" & vbNewLine & """" & str & """", MsgBoxStyle.Critical)
    End Sub


    ''' <summary>
    ''' Used if multiple parent folders option is in use.
    ''' </summary>
    Dim ParentFolderChosen As Boolean = True
    Private Sub DetermineParentFolder()
        Static R As New Random()

        Dim RollValue As Integer
        If TotalWeight + 1 > 0 Then
            RollValue = R.Next(1, TotalWeight + 1) ' use total weight as max value for roll
            'TextBox2.Text = CStr(RollValue)
        Else
            'MsgBox("Weights total value is not greater than 1, Program may have severe issues after now.", MsgBoxStyle.Critical, "CRITICAL ERROR")
            Exit Sub
        End If

        Dim sum As Integer = 0
        For i As Integer = 0 To FolderWeights.Count() - 1 ' test for which weight wins
            If FolderWeights(i) <> 0 AndAlso DisabledLine(i) = False AndAlso DirectoryWeightAlreadySubtractedFlag(i) = False Then
                sum += FolderWeights(i)

                If sum >= RollValue Then
                    ParentFolderChosen = True ' parent folder chosen
                    currentParentFolderIndex = i ' for external usage
                    TextBox3.Text = currentParentFolderIndex
                    Exit Sub
                    'Return FolderDirectories(i)
                End If
            End If
        Next

        'MsgBox("Catastrophic error has occured! You should not be seeing this message.", MsgBoxStyle.Critical)
        'End
        ParentFolderChosen = False ' no parent folder chosen
        'Return ""
    End Sub

    ''' <summary>
    ''' Legendary has been rolled.
    ''' </summary>
    Dim LegendaryFlag As Boolean = False
    ''' <summary>
    ''' Handles the roll value determination for a given file/folder.
    ''' </summary>
    Private Function Roll_Dice() As Integer
        If RollRulesEnabled = False Then
            Return 1 ' rules are disabled
        Else
            Static R As New Random() ' Static create a Random object so that we do not create a new one each time
            Dim RollValue As Integer
            Dim FirstTime As Boolean = True
            Dim Total_Times As Integer = 0 ' Tally for amount of times to add to list
            For Rolls As Integer = 1 To 1 Step -1 ' For range is inclusive to max value (so 1 to 2 is 1 AND 2 then stop)

                ' Get a random number. The second parameter is exclusive so (0,4) will always return 3 as a maximum
                RollValue = R.Next(1, 101) ' 1 - 100
                RollsRichTextBox.AppendText(RollValue & ", ")
                'If FirstTime = True Then
                '    RollValue = 1
                'End If

                If RollValue >= 36 Then ' Range = 36-100    (COMMON)
                    If CheckParity(RollValue) = True Then ' Even roll
                        Rolls += 1 ' Reroll 1 more time
                        Total_Times += 1
                    ElseIf FirstTime = True Then ' First time and odd roll
                        ' Odd roll; No additional rolls added
                        Total_Times += 1
                    ElseIf LegendaryFlag = True Then ' Odd reroll, but legendary was rolled (so add time anyway!)
                        ' Odd roll; No additional rolls added
                        Total_Times += 1
                    End If
                    ' Odd rerolls are not added. (see below)****
                ElseIf RollValue >= 19 Then ' Range = 19-35 (RARE!)
                    ' If reroll into this range and odd, add them then stop
                    If CheckParity(RollValue) = True Then ' Even roll
                        Rolls += 1 ' Reroll 1 more time
                        Total_Times += 5
                    ElseIf FirstTime = True Then ' First time and odd roll
                        ' Odd roll; No additional rolls added
                        Total_Times += 5
                    ElseIf LegendaryFlag = True Then ' Odd reroll, but legendary was rolled (so add time anyway!)
                        ' Odd roll; No additional rolls added
                        Total_Times += 5
                    End If
                    ' Odd rerolls are not added. (see below)****
                ElseIf RollValue >= 2 Then ' Range = 3-18   (UNIQUE!!)
                    ' (**If roll this as first roll or repeat roll then add and roll again) if reroll into this range and odd, add them then stop
GoTo1:
                    If CheckParity(RollValue) = True Then ' Even roll
                        Rolls += 1 ' Reroll 1 more time
                        Total_Times += 15
                    ElseIf FirstTime = True Then ' First time and odd roll
                        ' Odd roll; No additional rolls added
                        Total_Times += 15
                    ElseIf LegendaryFlag = True Then ' Odd reroll, but legendary was rolled (so add time!)
                        ' Odd roll; No additional rolls added
                        Total_Times += 15
                    End If
                    ' Odd rerolls are not added. (see below)****
                ElseIf RollValue >= 1 Then ' Range = 1-1    (LEGENDARY!!!)
                    ' (**If roll this as repeat roll or first roll then add and roll again) if reroll into this range or again, ignore and treat as usual
                    ' ****(If reroll value is odd value then no rolls are added but respective range's times value is added to Total_Times)
                    If LegendaryRollsEnabled Then
                        Total_Times += 50
                        Rolls += 10 ' Reroll 10 more times! (Stacks with additional rerolls including itself if it hits this range again!)
                        LegendaryFlag = True
                    Else ' Legendary rolls disabled
                        GoTo GoTo1 ' GoTo previous roll range instead
                    End If
                End If

                FirstTime = False
            Next
            RollsRichTextBoxBuffer.AppendText(RollsRichTextBox.Text)
            RollsRichTextBox.Clear()
            RollsRichTextBox.AppendText(RollsRichTextBoxBuffer.Text.Remove(RollsRichTextBoxBuffer.Text.Count - 2)) ' remove last ", " chars
            RollsRichTextBoxBuffer.Clear()
            Return Total_Times ' Return total times to add to list
        End If
    End Function

    ''' <summary>
    ''' Check the parity of a given integer and return true if even or false if odd.
    ''' </summary>
    Private Function CheckParity(ByVal Value As Integer) As Boolean
        If Value Mod 2 = 0 Then
            Return True ' Even number
        Else
            Return False ' Odd number
        End If
    End Function

    ''' <summary>
    ''' Used to temporarily make the roll button unclickable if legendary delay option is enabled.
    ''' </summary>
    Private Sub Sleeptimer_Tick(sender As Object, e As EventArgs) Handles SleepTimer.Tick
        RollButton.Enabled = True
        SleepTimer.Stop()
    End Sub

    Dim ListRolls(9) As Object
    Dim ListNames(9) As Object
    ''' <summary>
    ''' Display version of textbox with sanitized name. (Doesn't always show full path as name)
    ''' </summary>
    Dim ListNamesSan(9) As Object
    Dim HistoryList As New Dictionary(Of String, Integer) ' (Name, Roll)
    Dim MinListRoll As Integer = -9999
    ''' <summary>
    ''' Updates top rolls list.
    ''' </summary>
    Private Sub TopRollsList()
        Dim Roll As Integer = CInt(TimesBox.Text), Name As String = ResultBox.Text

        'RollNumTextBox.Text = CStr(CInt(RollNumTextBox.Text) + 1) ' debug related

        ' History dictionary work
        If HistoryList.ContainsKey(Name) Then
            Dim OldRoll As Integer = HistoryList.Item(Name)
            HistoryList.Item(Name) = Roll + OldRoll ' Set new value
            'ReAddTextBox.Text = CStr(CInt(ReAddTextBox.Text) + 1) ' debug related
            Roll += OldRoll ' Add history roll value to current roll
        Else
            HistoryList.Add(Name, Roll)
        End If

        If Roll >= MinListRoll Then ' If >= lowest roll on list
            ' Check if name already on list
            For i As Integer = 0 To ListNames.Count - 1
                If Name = ListNames(i).Text Then
                    ' set name/roll to blank if found on list
                    ListRolls(i).Text = ""
                    ListNames(i).Text = ""
                    ListNamesSan(i).text = ""
                    Exit For
                End If
            Next

            ' Determine position on list
            For i As Integer = 0 To ListRolls.Count - 1
                If Not ListRolls(i).Text = "" Then
                    If Roll >= CInt(ListRolls(i).Text) Then
                        ListInsert(Roll, i)
                        Exit For
                    End If
                Else
                    ListInsert(Roll, i)
                    Exit For
                End If
            Next

            ' Update minimum value
            If Not ListRolls(ListRolls.Count - 1).Text = "" Then ' Last list position
                MinListRoll = ListRolls(ListRolls.Count - 1).Text ' Set new minimum
            End If
        End If

        ClearListButton.Enabled = True
    End Sub

    ''' <summary>
    ''' Check if on visible textbox list (not the dictionary history list)
    ''' </summary>
    ''' <param name="Name">Name to be checked.</param>
    Private Function CheckIfOnList(ByVal Name As String) As Boolean
        For i As Integer = 0 To ListNames.Count - 1
            If Name = ListNames(i).Text Then ' On visible Textbox list (full path ones)
                Return True
            End If
        Next
        Return False
    End Function

    Private Sub ListInsert(roll As Integer, pos As Integer)
        Dim RollTmp As Integer = -9999, NameTmp As String = "Error Code: 420", NameSanTmp As String = "Error Code: 420"
        Dim RollTmp2 As Integer, NameTmp2 As String = "Error Code: 420", NameSanTmp2 As String = "Error Code: 420"

        If Not ListRolls(pos).Text = "" Then
            RollTmp = CInt(ListRolls(pos).Text)
            NameTmp = ListNames(pos).Text
            NameSanTmp = ListNamesSan(pos).Text
        End If
        ListRolls(pos).Text = CStr(roll)
        ListNames(pos).Text = ResultBox.Text
        ListNamesSan(pos).Text = ListRollDisplayName(ResultBox.Text)
        If RollTmp = -9999 Then
            Exit Sub ' Only replacing one
        End If
        Dim toggle1 As Boolean = True, EndFlag As Boolean = False
        For i As Integer = pos + 1 To ListRolls.Count - 1 ' Perform name and roll swapping
            If EndFlag = True Then ' Loop termination flag
                Exit For
            End If
            If toggle1 Then
                If Not ListRolls(i).Text = "" Then
                    RollTmp2 = ListRolls(i).Text
                    NameTmp2 = ListNames(i).Text
                    NameSanTmp2 = ListNamesSan(i).Text
                Else
                    EndFlag = True ' Flag for loop to end
                End If
                ListRolls(i).Text = RollTmp
                ListNames(i).Text = NameTmp
                ListNamesSan(i).Text = NameSanTmp
                toggle1 = Not toggle1
            Else
                If Not ListRolls(i).Text = "" Then
                    RollTmp = ListRolls(i).Text
                    NameTmp = ListNames(i).Text
                    NameSanTmp = ListNamesSan(i).Text
                Else
                    EndFlag = True ' Flag for loop to end
                End If
                ListRolls(i).Text = RollTmp2
                ListNames(i).Text = NameTmp2
                ListNamesSan(i).Text = NameSanTmp2
                toggle1 = Not toggle1
            End If
        Next
    End Sub

    ''' <summary>
    ''' Flag for whether each directory weight has already been subtracted when updating the trees.
    ''' </summary>
    Dim DirectoryWeightAlreadySubtractedFlag(29) As Boolean
    ''' <summary>
    ''' Update all file and folder trees. (Used for multiple folders option)
    ''' </summary>
    Public Sub UpdateTrees()

        ' clear and reset all first
        For i As Integer = 0 To 29
            If Not IsNothing(Me.DirectoryFileTree(i)) Then
                Me.DirectoryFileTree(i).Clear()
            End If
            If Not IsNothing(Me.DirectoryFolderTree(i)) Then
                Me.DirectoryFolderTree(i).Clear()
            End If

            DirectoryWeightAlreadySubtractedFlag(i) = False
        Next

        Dim total As Integer = 0
        For i As Integer = 0 To FolderDirectories.Count - 1 ' find sum of all weights
            If Not (FolderDirectories(i) = "" OrElse Me.DisabledLine(i) = True) Then
                total += FolderWeights(i)
            End If
        Next
        If total = 0 AndAlso MultipleFoldersFlag = True Then ' check if weights is equal to 0 (MultipleFoldersFlag = True is so that single directory mode doesn't mess up)
            MsgBox("Total of all weights equal 0 and all rows may be empty or deactivated.", MsgBoxStyle.Critical, "Error")
            Exit Sub
        Else
            Me.TotalWeight = total ' update total weight
        End If

        Dim didOnce As Boolean = False ' flag for doing all the group directory rows all at once (so skip once it's done from then on)

        For i As Integer = 0 To 29
            Dim grp1 As String = Me.FolderGroups(i)
            If (Not Me.FolderDirectories(i) = "" AndAlso (Not Me.FolderWeights(i) = 0 OrElse Not grp1 = "") AndAlso Not Me.DisabledLine(i) = True) Then
                If grp1 = "" Then

                    Dim tmp As New ArrayList, tmp2 As New ArrayList
                    tmp = FileTree.CreateFileList(Me.FolderDirectories(i))
                    Me.DirectoryFileTree(i) = tmp.Clone()
                    tmp2 = FileTree.CreateFolderList(Me.FolderDirectories(i))
                    Me.DirectoryFolderTree(i) = tmp2.Clone()
                    If tmp.Count = 0 Then ' no files found
                        ' if a parent folder has no files in it it will be 0 length, 
                        ' this may occur if there are no whitelisted file types in a folder, prevents pointless rolls
                        If (tmp.Count = 0 AndAlso FileFlag = True AndAlso FolderFlag = False) OrElse _
                           (tmp2.Count = 0 AndAlso FileFlag = False AndAlso FolderFlag = True) OrElse _
                           (tmp.Count = 0 AndAlso tmp2.Count = 0 AndAlso FileFlag = True AndAlso FolderFlag = True) Then ' no files found
                            If DirectoryWeightAlreadySubtractedFlag(i) = False Then
                                TotalWeight = TotalWeight - FolderWeights(i) ' subtract because it wont be used
                                DirectoryWeightAlreadySubtractedFlag(i) = True
                            End If
                        Else ' file(s)/folder(s) found
                            If DirectoryWeightAlreadySubtractedFlag(i) = True Then
                                TotalWeight = TotalWeight + FolderWeights(i) ' add because it will be used
                                DirectoryWeightAlreadySubtractedFlag(i) = False
                            End If
                        End If
                    End If

                Else ' is part of a group
                    If Not didOnce Then
                        didOnce = True

                        For Each struct In GroupWeightInfo
                            Dim createdListOnce As Boolean = False
                            ' GroupWeightInfo Structure:
                            ' currentGroup As Object ' boolean array of row size
                            ' currentGroupTotal As Integer
                            Dim tmp As New ArrayList, tmp2 As New ArrayList
                            For y As Integer = 0 To struct.currentGroup.Length - 1
                                If struct.currentGroup(y) = True Then
                                    If Not struct.currentGroupTotal = 0 Then
                                        If createdListOnce = False Then
                                            If Me.DisabledLine(y) = False Then
                                                tmp = FileTree.CreateFileList(Me.FolderDirectories(y), Me.FolderGroups(y)) ' send group value also so all directories of same group are put into variable
                                                tmp2 = FileTree.CreateFolderList(Me.FolderDirectories(y), Me.FolderGroups(y)) ' send group value also so all directories of same group are put into variable
                                                createdListOnce = True
                                                GoTo Goto112512
                                            End If
                                        Else
Goto112512:
                                            ' set all group parent directories to the exact same list
                                            Me.DirectoryFileTree(y) = tmp.Clone()
                                            Me.DirectoryFolderTree(y) = tmp2.Clone()

                                            If (tmp.Count = 0 AndAlso FileFlag = True AndAlso FolderFlag = False) OrElse _
                                               (tmp2.Count = 0 AndAlso FileFlag = False AndAlso FolderFlag = True) OrElse _
                                               (tmp.Count = 0 AndAlso tmp2.Count = 0 AndAlso FileFlag = True AndAlso FolderFlag = True) Then ' no files found
                                                If DirectoryWeightAlreadySubtractedFlag(y) = False Then
                                                    TotalWeight = TotalWeight - FolderWeights(y) ' subtract because it wont be used
                                                    DirectoryWeightAlreadySubtractedFlag(y) = True
                                                End If
                                            Else ' file(s)/folder(s) found
                                                If DirectoryWeightAlreadySubtractedFlag(y) = True Then
                                                    TotalWeight = TotalWeight + FolderWeights(y) ' add because it will be used
                                                    DirectoryWeightAlreadySubtractedFlag(y) = False
                                                End If
                                            End If
                                        End If

                                    Else ' disabled group
                                        If Not IsNothing(Me.DirectoryFileTree(y)) Then
                                            Me.DirectoryFileTree(y).Clear()
                                        End If
                                        If Not IsNothing(Me.DirectoryFolderTree(y)) Then
                                            Me.DirectoryFolderTree(y).Clear()
                                        End If
                                    End If
                                End If
                            Next
                        Next

                    End If
                End If
            ElseIf Me.DisabledLine(i) = True OrElse (Me.FolderWeights(i) = 0 AndAlso grp1 = "") Then

                If Not IsNothing(Me.DirectoryFileTree(i)) Then
                    Me.DirectoryFileTree(i).Clear()
                End If
                If Not IsNothing(Me.DirectoryFolderTree(i)) Then
                    Me.DirectoryFolderTree(i).Clear()
                End If

            End If
        Next
        TextBoxTotalWeight.Text = TotalWeight
    End Sub

    Private Sub RandomFile_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ListRolls(0) = ListRoll1 : ListRolls(1) = ListRoll2 : ListRolls(2) = ListRoll3 : ListRolls(3) = ListRoll4
        ListRolls(4) = ListRoll5 : ListRolls(5) = ListRoll6 : ListRolls(6) = ListRoll7 : ListRolls(7) = ListRoll8
        ListRolls(8) = ListRoll9 : ListRolls(9) = ListRoll10

        ListNames(0) = ListName1 : ListNames(1) = ListName2 : ListNames(2) = ListName3 : ListNames(3) = ListName4
        ListNames(4) = ListName5 : ListNames(5) = ListName6 : ListNames(6) = ListName7 : ListNames(7) = ListName8
        ListNames(8) = ListName9 : ListNames(9) = ListName10

        ListNamesSan(0) = ListNameSanitized1 : ListNamesSan(1) = ListNameSanitized2 : ListNamesSan(2) = ListNameSanitized3 : ListNamesSan(3) = ListNameSanitized4
        ListNamesSan(4) = ListNameSanitized5 : ListNamesSan(5) = ListNameSanitized6 : ListNamesSan(6) = ListNameSanitized7 : ListNamesSan(7) = ListNameSanitized8
        ListNamesSan(8) = ListNameSanitized9 : ListNamesSan(9) = ListNameSanitized10

        ListNameSanitized1.Location = ListName1.Location : ListNameSanitized2.Location = ListName2.Location : ListNameSanitized3.Location = ListName3.Location
        ListNameSanitized4.Location = ListName4.Location : ListNameSanitized5.Location = ListName5.Location : ListNameSanitized6.Location = ListName6.Location
        ListNameSanitized7.Location = ListName7.Location : ListNameSanitized8.Location = ListName8.Location : ListNameSanitized9.Location = ListName9.Location
        ListNameSanitized10.Location = ListName10.Location

        TopRollsExpandButton_original = New Point(Me.Size.Width, Me.Size.Height)
        If My.Settings.TopRollsExpandButton = True Then ' expand top rolls button related
            TopRollsExpandButton_toggle = True
            TopRollsExpandButton_Click()
        Else
            TopRollsExpandButton_toggle = False
            TopRollsExpandButton_Click()
        End If

        If MultipleFoldersFlag = True Then ' enable/disable drag and drop 
            TextBox1.AllowDrop = False
        Else
            TextBox1.AllowDrop = True
        End If

        If MultipleFoldersFlag = False Then
            RecomputeTreesChangesWereMade = True ' flag for tree refresh if search is opened at startup with single folder not rolled yet
        End If
    End Sub

    Function ListRollDisplayName(ByRef str As String) As String
        If str.Length > 1 Then
            Dim slashes As Integer = 0
            For pos As Integer = str.Length To 1 Step -1
                If Mid(str, pos, 1) = "\" Then
                    slashes += 1
                    If slashes = 3 Then
                        If pos - 1 = 1 AndAlso Mid(str, pos - 1, 1) = ":" Then ' check if next 2 are not drive letter and colon
                            Return str
                        Else
                            Return "..." & Mid(str, pos, (str.Length) - pos + 1)
                        End If
                    End If
                End If
            Next
        End If
        Return str
    End Function

    Sub ListNameMouseEnter(pos As Integer)
        ListNamesSan(pos - 1).Visible = False
        ListNames(pos - 1).Visible = True
    End Sub
    Sub ListNameMouseLeave(pos As Integer)
        If Not ListNames(pos - 1).Focused = True Then
            ListNamesSan(pos - 1).Visible = True
            ListNames(pos - 1).Visible = False
        End If
    End Sub

    Private Sub ListNameSanitized1_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized1.MouseEnter
        ListNameMouseEnter(1)
    End Sub
    Private Sub ListNameSanitized2_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized2.MouseEnter
        ListNameMouseEnter(2)
    End Sub
    Private Sub ListNameSanitized3_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized3.MouseEnter
        ListNameMouseEnter(3)
    End Sub
    Private Sub ListNameSanitized4_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized4.MouseEnter
        ListNameMouseEnter(4)
    End Sub
    Private Sub ListNameSanitized5_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized5.MouseEnter
        ListNameMouseEnter(5)
    End Sub
    Private Sub ListNameSanitized6_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized6.MouseEnter
        ListNameMouseEnter(6)
    End Sub
    Private Sub ListNameSanitized7_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized7.MouseEnter
        ListNameMouseEnter(7)
    End Sub
    Private Sub ListNameSanitized8_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized8.MouseEnter
        ListNameMouseEnter(8)
    End Sub
    Private Sub ListNameSanitized9_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized9.MouseEnter
        ListNameMouseEnter(9)
    End Sub
    Private Sub ListNameSanitized10_MouseEnter(sender As Object, e As EventArgs) Handles ListNameSanitized10.MouseEnter
        ListNameMouseEnter(10)
    End Sub

    Private Sub ListName1_MouseLeave(sender As Object, e As EventArgs) Handles ListName1.MouseLeave
        ListNameMouseLeave(1)
    End Sub
    Private Sub ListName2_MouseLeave(sender As Object, e As EventArgs) Handles ListName2.MouseLeave
        ListNameMouseLeave(2)
    End Sub
    Private Sub ListName3_MouseLeave(sender As Object, e As EventArgs) Handles ListName3.MouseLeave
        ListNameMouseLeave(3)
    End Sub
    Private Sub ListName4_MouseLeave(sender As Object, e As EventArgs) Handles ListName4.MouseLeave
        ListNameMouseLeave(4)
    End Sub
    Private Sub ListName5_MouseLeave(sender As Object, e As EventArgs) Handles ListName5.MouseLeave
        ListNameMouseLeave(5)
    End Sub
    Private Sub ListName6_MouseLeave(sender As Object, e As EventArgs) Handles ListName6.MouseLeave
        ListNameMouseLeave(6)
    End Sub
    Private Sub ListName7_MouseLeave(sender As Object, e As EventArgs) Handles ListName7.MouseLeave
        ListNameMouseLeave(7)
    End Sub
    Private Sub ListName8_MouseLeave(sender As Object, e As EventArgs) Handles ListName8.MouseLeave
        ListNameMouseLeave(8)
    End Sub
    Private Sub ListName9_MouseLeave(sender As Object, e As EventArgs) Handles ListName9.MouseLeave
        ListNameMouseLeave(9)
    End Sub
    Private Sub ListName10_MouseLeave(sender As Object, e As EventArgs) Handles ListName10.MouseLeave
        ListNameMouseLeave(10)
    End Sub

    Private Sub ListName1_Leave(sender As Object, e As EventArgs) Handles ListName1.Leave
        ListNameMouseLeave(1)
    End Sub
    Private Sub ListName2_Leave(sender As Object, e As EventArgs) Handles ListName2.Leave
        ListNameMouseLeave(2)
    End Sub
    Private Sub ListName3_Leave(sender As Object, e As EventArgs) Handles ListName3.Leave
        ListNameMouseLeave(3)
    End Sub
    Private Sub ListName4_Leave(sender As Object, e As EventArgs) Handles ListName4.Leave
        ListNameMouseLeave(4)
    End Sub
    Private Sub ListName5_Leave(sender As Object, e As EventArgs) Handles ListName5.Leave
        ListNameMouseLeave(5)
    End Sub
    Private Sub ListName6_Leave(sender As Object, e As EventArgs) Handles ListName6.Leave
        ListNameMouseLeave(6)
    End Sub
    Private Sub ListName7_Leave(sender As Object, e As EventArgs) Handles ListName7.Leave
        ListNameMouseLeave(7)
    End Sub
    Private Sub ListName8_Leave(sender As Object, e As EventArgs) Handles ListName8.Leave
        ListNameMouseLeave(8)
    End Sub
    Private Sub ListName9_Leave(sender As Object, e As EventArgs) Handles ListName9.Leave
        ListNameMouseLeave(9)
    End Sub
    Private Sub ListName10_Leave(sender As Object, e As EventArgs) Handles ListName10.Leave
        ListNameMouseLeave(10)
    End Sub

    ''' <summary>
    ''' Legendary rolls enabled flag.
    ''' </summary>
    Dim LegendaryRollsEnabled As Boolean = True
    ''' <summary>
    ''' Legendary rolls option.
    ''' </summary>
    Private Sub YesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem.Click
        My.Settings.LegendaryRolls = True
        YesToolStripMenuItem.Checked = True
        NoToolStripMenuItem.Checked = False
        LegendaryRollsEnabled = True
    End Sub

    ''' <summary>
    ''' Legendary rolls option.
    ''' </summary>
    Private Sub NoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem.Click
        My.Settings.LegendaryRolls = False
        YesToolStripMenuItem.Checked = False
        NoToolStripMenuItem.Checked = True
        LegendaryRollsEnabled = False
    End Sub

    ''' <summary>
    ''' Multiple parent folders option.
    ''' </summary>
    Private Sub YesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem1.Click
        My.Settings.MultipleFolders = True

        Dim flag1 As Boolean = False
        If Not YesToolStripMenuItem1.Checked = True Then ' if not true already
            flag1 = True
        End If
        YesToolStripMenuItem1.Checked = True
        NoToolStripMenuItem1.Checked = False
        BrowseFolderButton.Enabled = False
        'TextBox2.Visible = True
        'Label6.Visible = True
        TextBox1.ReadOnly = True
        TextBox1.AllowDrop = False
        TextBox1.Text = "[Multiple Folders]"
        'Label6.Visible = True
        'TextBox2.Visible = True
        FoldersExpandButton.Enabled = True
        MultipleFoldersFlag = True

        If flag1 Then
            If Search.Visible = True Then
                Search.RefreshButton_Click()
            End If
        End If
    End Sub

    ''' <summary>
    ''' Multiple parent folders option.
    ''' </summary>
    Private Sub NoToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem1.Click
        My.Settings.MultipleFolders = False

        Dim flag1 As Boolean = False
        If Not NoToolStripMenuItem1.Checked = True Then ' if not true already
            flag1 = True
        End If
        YesToolStripMenuItem1.Checked = False
        NoToolStripMenuItem1.Checked = True
        BrowseFolderButton.Enabled = True
        'TextBox2.Visible = False
        'Label6.Visible = False
        TextBox1.ReadOnly = False
        TextBox1.AllowDrop = True
        TextBox1.Text = My.Settings.SingleParentFolder
        ParentFolderSingleFolderChanged = True ' make sure tree(s) of text are computed via this flag
        'Label6.Visible = False
        'TextBox2.Visible = False
        FoldersExpandButton.Enabled = False
        MultipleFoldersFlag = False

        If flag1 Then
            If Search.Visible = True Then
                Search.RefreshButton_Click()
            End If
        End If
    End Sub

    Dim LegendaryDelay As Boolean = False
    ''' <summary>
    ''' Legendary delays option.
    ''' </summary>
    Private Sub YesToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem2.Click
        My.Settings.LegendaryDelays = True
        YesToolStripMenuItem2.Checked = True
        NoToolStripMenuItem2.Checked = False
        LegendaryDelay = True
    End Sub

    ''' <summary>
    ''' Legendary delays option.
    ''' </summary>
    Private Sub NoToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem2.Click
        My.Settings.LegendaryDelays = False
        YesToolStripMenuItem2.Checked = False
        NoToolStripMenuItem2.Checked = True
        LegendaryDelay = False
    End Sub

    ''' <summary>
    ''' Folders rollable option.
    ''' </summary>
    Private Sub FoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FoldersToolStripMenuItem.Click
        If FoldersToolStripMenuItem.Checked = True AndAlso FilesToolStripMenuItem.Checked = False Then
            My.Settings.RollCandidates = "Files"
            FoldersToolStripMenuItem.Checked = False
            FolderFlag = False
            FilesToolStripMenuItem.Checked = True
            FileFlag = True
        ElseIf FoldersToolStripMenuItem.Checked = True AndAlso FilesToolStripMenuItem.Checked = True Then
            My.Settings.RollCandidates = "Files"
            FoldersToolStripMenuItem.Checked = False
            FolderFlag = False
        ElseIf FoldersToolStripMenuItem.Checked = False AndAlso FilesToolStripMenuItem.Checked = True Then
            My.Settings.RollCandidates = "Both"
            FoldersToolStripMenuItem.Checked = True
            FolderFlag = True
        End If
    End Sub

    ''' <summary>
    ''' Files rollable option.
    ''' </summary>
    Private Sub FilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilesToolStripMenuItem.Click
        If FoldersToolStripMenuItem.Checked = True AndAlso FilesToolStripMenuItem.Checked = False Then
            My.Settings.RollCandidates = "Both"
            FilesToolStripMenuItem.Checked = True
            FileFlag = True
        ElseIf FoldersToolStripMenuItem.Checked = True AndAlso FilesToolStripMenuItem.Checked = True Then
            My.Settings.RollCandidates = "Folders"
            FilesToolStripMenuItem.Checked = False
            FileFlag = False
        ElseIf FoldersToolStripMenuItem.Checked = False AndAlso FilesToolStripMenuItem.Checked = True Then
            My.Settings.RollCandidates = "Folders"
            FoldersToolStripMenuItem.Checked = True
            FolderFlag = True
            FilesToolStripMenuItem.Checked = False
            FileFlag = False
        End If
    End Sub

    ''' <summary>
    ''' Roll and execute allowed setting.
    ''' </summary>
    Private Sub YesToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem3.Click
        My.Settings.RollAndExecute = True
        NoToolStripMenuItem3.Checked = False
        YesToolStripMenuItem3.Checked = True
        RollButton.Text = "Roll && Execute"
    End Sub

    ''' <summary>
    ''' Roll and execute disallowed setting.
    ''' </summary>
    Private Sub NoToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem3.Click
        My.Settings.RollAndExecute = False
        NoToolStripMenuItem3.Checked = True
        YesToolStripMenuItem3.Checked = False
        RollButton.Text = "Roll"
    End Sub

    ''' <summary>
    ''' Shortcut creation allowed setting. (In output folder)
    ''' </summary>
    Private Sub YesToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem4.Click
        If My.Settings.CreateShortcuts = False Then
            SpecifiedShortcutFolder = False ' make sure output folder will be specified
        End If
        My.Settings.CreateShortcuts = True
        shortcutCreationAllowed = True
        YesToolStripMenuItem4.Checked = True
        NoToolStripMenuItem4.Checked = False
    End Sub

    ''' <summary>
    ''' Shortcut creation disabled setting. (In output folder)
    ''' </summary>
    Private Sub NoToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem4.Click
        My.Settings.CreateShortcuts = False
        shortcutCreationAllowed = False
        YesToolStripMenuItem4.Checked = False
        NoToolStripMenuItem4.Checked = True
    End Sub

    Private Sub FoldersExpandButton_Click(sender As Object, e As EventArgs) Handles FoldersExpandButton.Click
        ParentFolders.Show()
    End Sub


    Dim TopRollsExpandButton_original As New Size
    Dim TopRollsExpandButton_toggle As Boolean = True
    ''' <summary>
    ''' Expands main screen view to show top rolls.
    ''' </summary>
    Private Sub TopRollsExpandButton_Click() Handles TopRollsExpandButton.Click
        If TopRollsExpandButton_toggle = True Then
            My.Settings.TopRollsExpandButton = True
            TopRollsGroupBox.Visible = True
            ' original variable is defined in shown event instead of right here
            TopRollsGroupBox.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            Me.MinimumSize = New Size(1200, Me.Size.Height)
            Me.MaximumSize = New Size(1600, Me.Size.Height)
            Me.Size = New Size(1200, Me.Size.Height)
            TopRollsGroupBox.Anchor = AnchorStyles.Left Or AnchorStyles.Right
            If Not IsNothing(My.Settings.TopRollsExpandSize) AndAlso Not (My.Settings.TopRollsExpandSize.Width = 0 AndAlso My.Settings.TopRollsExpandSize.Height = 0) Then
                Me.Size = My.Settings.TopRollsExpandSize
            End If
            TopRollsExpandButton.Text = "<"
            TopRollsExpandButton_toggle = False
        Else
            My.Settings.TopRollsExpandButton = False
            TopRollsGroupBox.Visible = False
            If Me.Size.Width > 685 Then
                My.Settings.TopRollsExpandSize = Me.Size
            End If
            Me.Size = TopRollsExpandButton_original
            Me.MinimumSize = New Size(685, Me.Size.Height)
            Me.MaximumSize = New Size(685, Me.Size.Height)
            TopRollsGroupBox.Anchor = AnchorStyles.Left Or AnchorStyles.Top
            TopRollsExpandButton.Text = ">"
            TopRollsExpandButton_toggle = True
        End If
    End Sub

    ''' <summary>
    ''' holds original of Duration variable
    ''' </summary>
    Dim original As Integer
    ''' <summary>
    ''' Flag that automation is occuring.
    ''' </summary>
    Dim AutomateFlag As Boolean
    ''' <summary>
    ''' Flag for when roll once list is enabled. (Currently just disables rolls from being added to the roll once list when using automate)
    ''' </summary>
    Dim RollOnceAutomateFlag As Boolean = False
    ''' <summary>
    ''' Automates rolls for specified amount
    ''' </summary>
    Private Sub AutomateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutomateToolStripMenuItem.Click
        Dim tmp As String = InputBox("Enter number of times to roll", "", CStr(My.Settings.AutomateNum))
        If Not tmp = "" Then ' not aborted
            If Not IsNumeric(tmp) Then
                MsgBox("Value given was not a valid number.", MsgBoxStyle.Exclamation, "Error")
            ElseIf Not tmp = "" Then
                Dim input As Integer = CInt(tmp)
                If IsNumeric(input) Then
                    My.Settings.AutomateNum = input
                    original = Duration
                    Duration = input ' Set duration (number of rolls) to input
                    AutomateFlag = True
                    If RollOnceListFlag = True Then
                        RollOnceAutomateFlag = True
                        RollOnceListFlag = False
                    End If
                    RollButton.PerformClick()

                    ' Button is disabled via locked boolean variable
                    ' Duration set back to normal in button code
                Else
                    MsgBox("Input was not a valid number!", MsgBoxStyle.Exclamation, "Error")
                End If
            End If
        End If
    End Sub

    ' breaks the program (debug related)
    Private Sub BreakButton_Click(sender As Object, e As EventArgs) Handles BreakButton.Click
        Debugger.Break()
    End Sub

    Private Sub BrowseButton1_Click(sender As Object, e As EventArgs) Handles BrowseButton1.Click
        BrowseButton(ListName1.Text)
    End Sub

    Private Sub BrowseButton2_Click(sender As Object, e As EventArgs) Handles BrowseButton2.Click
        BrowseButton(ListName2.Text)
    End Sub

    Private Sub BrowseButton3_Click(sender As Object, e As EventArgs) Handles BrowseButton3.Click
        BrowseButton(ListName3.Text)
    End Sub

    Private Sub BrowseButton4_Click(sender As Object, e As EventArgs) Handles BrowseButton4.Click
        BrowseButton(ListName4.Text)
    End Sub

    Private Sub BrowseButton5_Click(sender As Object, e As EventArgs) Handles BrowseButton5.Click
        BrowseButton(ListName5.Text)
    End Sub

    Private Sub BrowseButton6_Click(sender As Object, e As EventArgs) Handles BrowseButton6.Click
        BrowseButton(ListName6.Text)
    End Sub

    Private Sub BrowseButton7_Click(sender As Object, e As EventArgs) Handles BrowseButton7.Click
        BrowseButton(ListName7.Text)
    End Sub

    Private Sub BrowseButton8_Click(sender As Object, e As EventArgs) Handles BrowseButton8.Click
        BrowseButton(ListName8.Text)
    End Sub

    Private Sub BrowseButton9_Click(sender As Object, e As EventArgs) Handles BrowseButton9.Click
        BrowseButton(ListName9.Text)
    End Sub

    Private Sub BrowseButton10_Click(sender As Object, e As EventArgs) Handles BrowseButton10.Click
        BrowseButton(ListName10.Text)
    End Sub

    Private Sub BrowseButton(path As String)
        If path <> "" Then
            Try
                If System.IO.File.Exists(path) Then
                    Process.Start(path)
                ElseIf System.IO.Directory.Exists(path) Then
                    Process.Start(path)
                End If
            Catch ex As Exception
                If ex.Message = "Application not found" Then
                    MsgBox("No program associated with this filetype", MsgBoxStyle.Exclamation, "Error")
                End If
            End Try
        End If
    End Sub

    Public Sub ContainingButton(path As String)
        If path <> "" Then
            Call Shell("explorer /select," & """" & path & """", AppWinStyle.NormalFocus) ' select in containing folder
        End If
    End Sub

    Private Sub ContainingButton1_Click(sender As Object, e As EventArgs) Handles ContainingButton1.Click
        ContainingButton(ListName1.Text)
    End Sub

    Private Sub ContainingButton2_Click(sender As Object, e As EventArgs) Handles ContainingButton2.Click
        ContainingButton(ListName2.Text)
    End Sub

    Private Sub ContainingButton3_Click(sender As Object, e As EventArgs) Handles ContainingButton3.Click
        ContainingButton(ListName3.Text)
    End Sub

    Private Sub ContainingButton4_Click(sender As Object, e As EventArgs) Handles ContainingButton4.Click
        ContainingButton(ListName4.Text)
    End Sub

    Private Sub ContainingButton5_Click(sender As Object, e As EventArgs) Handles ContainingButton5.Click
        ContainingButton(ListName5.Text)
    End Sub

    Private Sub ContainingButton6_Click(sender As Object, e As EventArgs) Handles ContainingButton6.Click
        ContainingButton(ListName6.Text)
    End Sub

    Private Sub ContainingButton7_Click(sender As Object, e As EventArgs) Handles ContainingButton7.Click
        ContainingButton(ListName7.Text)
    End Sub

    Private Sub ContainingButton8_Click(sender As Object, e As EventArgs) Handles ContainingButton8.Click
        ContainingButton(ListName8.Text)
    End Sub

    Private Sub ContainingButton9_Click(sender As Object, e As EventArgs) Handles ContainingButton9.Click
        ContainingButton(ListName9.Text)
    End Sub

    Private Sub ContainingButton10_Click(sender As Object, e As EventArgs) Handles ContainingButton10.Click
        ContainingButton(ListName10.Text)
    End Sub

    ' Used to specify parent folder if not using multiple parent folders
    Private Sub BrowseFolderButton_Click(sender As Object, e As EventArgs) Handles BrowseFolderButton.Click
        If Not (FolderBrowserDialog1.ShowDialog() = DialogResult.Cancel Or DialogResult.Abort) Then
            TextBox1.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub

    Private Sub SearchButton_Click(sender As Object, e As EventArgs) Handles SearchButton.Click
        If LookupTextBox.Text <> "" Then
            If System.IO.Directory.Exists(LookupTextBox.Text) OrElse System.IO.File.Exists(LookupTextBox.Text) Then
                If HistoryList.ContainsKey(LookupTextBox.Text) Then
                    MsgBox("Roll value is currently " & HistoryList.Item(LookupTextBox.Text), MsgBoxStyle.Information, "File or folder was found!")
                Else
                    MsgBox("File or folder has NOT been rolled!" & vbNewLine & _
                           "Please check to make sure that the file or folder actually exists or was not entered wrong.", MsgBoxStyle.Exclamation, "Error")
                End If
            Else
                MsgBox("Please enter the full path of an existing file or folder and try again." & vbNewLine & _
                       "If looking up a file please include the extension.", MsgBoxStyle.Critical, "Error")
            End If
        End If
    End Sub

    ''' <summary>
    ''' Select a file for lookup in history list dictionary.
    ''' </summary>
    Private Sub SearchFileButton_Click(sender As Object, e As EventArgs) Handles SearchFileButton.Click
        If Not (OpenFileDialog1.ShowDialog() = DialogResult.Cancel Or DialogResult.Abort) Then
            LookupTextBox.Text = OpenFileDialog1.FileName
            SearchButton.PerformClick()
        End If
    End Sub

    ''' <summary>
    ''' Select a folder for lookup in history list dictionary.
    ''' </summary>
    Private Sub SearchFolderButton_Click(sender As Object, e As EventArgs) Handles SearchFolderButton.Click
        If Not (FolderBrowserDialog1.ShowDialog() = DialogResult.Cancel Or DialogResult.Abort) Then
            LookupTextBox.Text = FolderBrowserDialog1.SelectedPath
            SearchButton.PerformClick()
        End If
    End Sub

    Private Sub FileTypesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FileTypesToolStripMenuItem.Click
        FileTypesForm.Show()
        FileTypesForm.Focus()
    End Sub

    ''' <summary>
    ''' Clear top rolls list of all history and textbox values.
    ''' </summary>
    Private Sub ClearListButton_Click(sender As Object, e As EventArgs) Handles ClearListButton.Click
        ClearListButton.Enabled = False

        For i As Integer = 0 To 9 ' clear textboxes
            ListRolls(i).Text = ""
            ListNames(i).Text = ""
            ListNamesSan(i).Text = ""
        Next
        HistoryList.Clear() ' clear full history

        MinListRoll = 0 ' reset minimum value
    End Sub

    ''' <summary>
    ''' Re-compute single parent folder file/folder trees flag.
    ''' </summary>
    Public ParentFolderSingleFolderChanged As Boolean = False
    ''' <summary>
    ''' Single folder option textbox changed event.
    ''' </summary>
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If MultipleFoldersFlag = False Then ' single folder mode
            ' mark as changed
            ParentFolderSingleFolderChanged = True
        End If
    End Sub

    Private Sub RandomFile_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.NotifyIcon1.Icon = My.Resources.icon

        If My.Settings.RollCandidates = "Both" Then ' roll candidates setting
            FoldersToolStripMenuItem.Checked = True
            FolderFlag = True
            FilesToolStripMenuItem.Checked = True
            FileFlag = True
        ElseIf My.Settings.RollCandidates = "Folders" Then
            FoldersToolStripMenuItem.Checked = True
            FolderFlag = True
            FilesToolStripMenuItem.Checked = False
            FileFlag = False
        Else ' Files (default)
            FoldersToolStripMenuItem.Checked = False
            FolderFlag = False
            FilesToolStripMenuItem.Checked = True
            FileFlag = True
        End If
        If My.Settings.LegendaryRolls = True Then ' legendary rolls setting
            YesToolStripMenuItem.Checked = True
            NoToolStripMenuItem.Checked = False
            LegendaryRollsEnabled = True
        Else
            YesToolStripMenuItem.Checked = False
            NoToolStripMenuItem.Checked = True
            LegendaryRollsEnabled = False
        End If
        If My.Settings.LegendaryDelays = False Then ' legendary roll delays setting
            YesToolStripMenuItem2.Checked = False
            NoToolStripMenuItem2.Checked = True
            LegendaryDelay = False
        Else
            YesToolStripMenuItem2.Checked = True
            NoToolStripMenuItem2.Checked = False
            LegendaryDelay = True
        End If
        If My.Settings.MultipleFolders = True Then ' multiple folders allow/disallowed setting
            YesToolStripMenuItem1.Checked = True
            NoToolStripMenuItem1.Checked = False
            BrowseFolderButton.Enabled = False
            'TextBox2.Visible = True
            'Label6.Visible = True
            TextBox1.ReadOnly = True
            TextBox1.Text = "[Multiple Folders]"
            'Label6.Visible = True
            'TextBox2.Visible = True
            FoldersExpandButton.Enabled = True
            MultipleFoldersFlag = True
        Else
            YesToolStripMenuItem1.Checked = False
            NoToolStripMenuItem1.Checked = True
            BrowseFolderButton.Enabled = True
            'TextBox2.Visible = False
            'Label6.Visible = False
            TextBox1.ReadOnly = False
            TextBox1.Text = My.Settings.SingleParentFolder
            ParentFolderSingleFolderChanged = True ' make sure tree(s) of text are computed via this flag
            'Label6.Visible = False
            'TextBox2.Visible = False
            FoldersExpandButton.Enabled = False
            MultipleFoldersFlag = False
        End If
        If My.Settings.RollAndExecute = False Then ' roll and execute button change setting
            RollButton.Text = "Roll"
            YesToolStripMenuItem3.Checked = False
            NoToolStripMenuItem3.Checked = True
        Else
            RollButton.Text = "Roll && Execute"
            YesToolStripMenuItem3.Checked = True
            NoToolStripMenuItem3.Checked = False
        End If
        If My.Settings.CreateShortcuts = False Then ' create shortcuts per roll setting
            shortcutCreationAllowed = False
            YesToolStripMenuItem4.Checked = False
            NoToolStripMenuItem4.Checked = True
        Else
            shortcutCreationAllowed = True
            YesToolStripMenuItem4.Checked = True
            NoToolStripMenuItem4.Checked = False
        End If
        If My.Settings.AlreadyRolledListFlag = True Then
            RollOnceListFlag = True
            YesToolStripMenuItem5.Checked = True
            NoToolStripMenuItem5.Checked = False
            CreateRollOnceList()
            IgnoreRollButton.Visible = True
            ClearRollOnceListButton.Visible = True
        Else
            RollOnceListFlag = False
            YesToolStripMenuItem5.Checked = False
            NoToolStripMenuItem5.Checked = True
            IgnoreRollButton.Visible = False
            ClearRollOnceListButton.Visible = False
        End If
        If My.Settings.RollRulesEnabled = True Then
            RollRulesEnabled = True
            NoToolStripMenuItem6.Checked = False
            YesToolStripMenuItem6.Checked = True

            RollsRichTextBox.Enabled = True
            Label4.Enabled = True
            TimesBox.Enabled = True
            Label3.Enabled = True
        Else
            RollRulesEnabled = False
            NoToolStripMenuItem6.Checked = True
            YesToolStripMenuItem6.Checked = False

            RollsRichTextBox.Enabled = False
            Label4.Enabled = False
            TimesBox.Enabled = False
            Label3.Enabled = False
        End If

        ' set and check saved main form starting location if it's within screen(s) bounds and whether to reset or not
        If Not IsNothing(My.Settings.MainFormLocation) AndAlso _
             (My.Settings.MainFormLocation.X <> -1 AndAlso My.Settings.MainFormLocation.Y <> -1) Then
            If FormVisible(My.Settings.MainFormLocation) OrElse _
                FormVisible(New Point(My.Settings.MainFormLocation.X + Me.Size.Width, My.Settings.MainFormLocation.Y)) Then
                Me.Location = My.Settings.MainFormLocation
            Else
                Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width / 2) - (Me.Size.Width / 2), (Screen.PrimaryScreen.WorkingArea.Height / 2) - (Me.Size.Height / 2))
            End If
        End If
    End Sub

    ''' <summary>
    ''' Main form minimizes to tray flag.
    ''' </summary>
    Dim MinimizeButtonFlag As Boolean = False
    ''' <summary>
    ''' Button used for minimizing the program to the system tray.
    ''' </summary>
    Private Sub MinimizeToTrayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MinimizeToTrayToolStripMenuItem.Click
        MinimizeButtonFlag = True ' button was just clicked
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub MainForm_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If MinimizeButtonFlag = True Then ' only when "minimize to tray" button is used and not standard minimizing
            MinimizeButtonFlag = False
            NotifyIcon1.Visible = True
            Me.Hide()
            'NotifyIcon1.BalloonTipText = "Minimized To Tray"
            'NotifyIcon1.ShowBalloonTip(250)
        ElseIf Me.WindowState = FormWindowState.Normal Then
            NotifyIcon1.Visible = False
        End If
    End Sub

    ''' <summary>
    ''' Test if form location is out of bounds (usually only possible if was saved on a now not active monitor or settings file edited)
    ''' </summary>
    Public Function FormVisible(ByRef Loc As Point) As Boolean
        Try
            For Each screen1 In Screen.AllScreens
                If (screen1.Bounds.Contains(Loc)) Then
                    Return True
                End If
            Next
        Catch ex As Exception
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Tray icon mouse clicking related.
    ''' </summary>
    Private Sub NotifyIcon1_MouseClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Right Then ' right click
            TrayMenu.Show()
            TrayMenu.Activate()
            TrayMenu.Width = 1 'it will be set behind the menu, so it's 1x1 pixels
            TrayMenu.Height = 1
        ElseIf e.Button = Windows.Forms.MouseButtons.Left Then ' left click
            TrayMenu.Hide()
            Me.Show()
            Me.WindowState = FormWindowState.Normal
        End If
    End Sub

    ''' <summary>
    ''' Handles manual specification of output folder.
    ''' </summary>
    Private Sub OutputFolderLabel_Click(sender As Object, e As EventArgs) Handles OutputFolderLabel.Click
        If My.Settings.CreateShortcuts = False Then
            MsgBox("Shortcut creation is disabled.", MsgBoxStyle.Information, "Error")
        End If
        OutputFolderConfig() ' specify output folder
    End Sub

    Private Sub SaveShortcutButton_Click(sender As Object, e As EventArgs) Handles SaveShortcutButton.Click
        If IO.File.Exists(ResultBox.Text) OrElse IO.Directory.Exists(ResultBox.Text) Then
            Dim SaveLocation = InputBox("Specify location to save one shortcut", "Specify directory", My.Computer.FileSystem.SpecialDirectories.Desktop)
            If Not SaveLocation = "" Then ' not aborted or canceled
                Dim myStr() As String = Split(ResultBox.Text, "\")
                Dim shortcutName As String = myStr(myStr.Count - 1) ' put last element into variable
                Dim myStr2() As String = Split(shortcutName, ".")
                shortcutName = myStr2(0)
                For i As Integer = 1 To myStr2.Count - 2 ' include everything but last element (extension)
                    shortcutName &= "." & myStr2(i)
                Next

                Dim tmp As String = Me.TimesBox.Text
                Me.TimesBox.Text = "1" ' set to 1 temporarily to create only one shortcut
                If Me.shortcutCreationAllowed = False Then
                    Me.shortcutCreationAllowed = True ' temporarily set to true to make sure that shortcut creation is performed
                    DetermineShortcuts(shortcutName, SaveLocation) ' create shortcut
                    Me.shortcutCreationAllowed = False
                Else
                    DetermineShortcuts(shortcutName, SaveLocation) ' create shortcut
                End If
                Me.TimesBox.Text = tmp
            End If
        Else
            MsgBox("Not a valid roll.", MsgBoxStyle.Information, "Error")
        End If
    End Sub

    Public Sub RandomFile_FormClosing() Handles MyBase.FormClosing
        If Me.WindowState = FormWindowState.Minimized Then ' minimizing makes location -32000 coords
            If Me.Visible = False Then
                Me.Visible = True
            End If
            Me.Opacity = 0.0
            Me.WindowState = FormWindowState.Normal
        End If
        My.Settings.MainFormLocation = Me.Location

        If TopRollsExpandButton_toggle = False Then
            My.Settings.TopRollsExpandSize = Me.Size
        End If

RestartLoop1:  ' restart loop because form has been removed from iteration
        For Each myForm As Form In My.Application.OpenForms
            ' perform form closing events to save any settings
            If myForm.Name = Search.Name Then
                Search.Close()
                GoTo RestartLoop1
            ElseIf myForm.Name = FavoritesList.Name Then
                FavoritesList.Close()
                GoTo RestartLoop1
            End If
        Next
    End Sub

    Private Sub YesToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem5.Click
        RollOnceListFlag = True
        My.Settings.AlreadyRolledListFlag = True
        CreateRollOnceList()
        IgnoreRollButton.Visible = True
        IgnoreRollButton.Enabled = True
        ClearRollOnceListButton.Visible = True
        YesToolStripMenuItem5.Checked = True
        NoToolStripMenuItem5.Checked = False
        RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub

    Private Sub NoToolStripMenuItem5_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem5.Click
        RollOnceListFlag = False
        My.Settings.AlreadyRolledListFlag = False
        ParentFolderSingleFolderChanged = True ' recompute for single folder option
        IgnoreRollButton.Visible = False
        ClearRollOnceListButton.Visible = False
        YesToolStripMenuItem5.Checked = False
        NoToolStripMenuItem5.Checked = True
        RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub

    Sub CreateRollOnceList()
        If IsNothing(RollOnceList) Then
            RollOnceList = New ArrayList
        End If
        If Not IsNothing(My.Settings.AlreadyRolledList) AndAlso My.Settings.AlreadyRolledList.Count > 0 Then
            For Each item In My.Settings.AlreadyRolledList
                If Not RollOnceList.Contains(CStr(item)) AndAlso (Not IsNothing(My.Settings.FavoriteRollsList) AndAlso Not My.Settings.FavoriteRollsList.Contains(item)) OrElse _
                    IsNothing(My.Settings.FavoriteRollsList) Then
                    RollOnceList.Add(CStr(item))
                End If
            Next
        End If
    End Sub

    Private Sub FavoriteButton_Click(sender As Object, e As EventArgs) Handles FavoriteButton.Click
        If IsNothing(My.Settings.FavoriteRollsList) Then
            My.Settings.FavoriteRollsList = New Specialized.StringCollection
        End If

        If IO.File.Exists(ResultBox.Text) OrElse IO.Directory.Exists(ResultBox.Text) Then
            If FavoriteButton.Text = "Favorite" Then ' current roll is not favorite
                If Not My.Settings.FavoriteRollsList.Contains(ResultBox.Text) Then
                    My.Settings.FavoriteRollsList.Add(ResultBox.Text)
                    FavoriteRolled()

                    If RollOnceListFlag = True Then
                        RollOnceListHandling("Remove") ' remove from roll once list if it is on it
                    Else
                        ' favoriting still removes from exlusion list whether the setting is in use or not
                        If Not IsNothing(My.Settings.AlreadyRolledList) AndAlso My.Settings.AlreadyRolledList.Contains(ResultBox.Text) Then
                            My.Settings.AlreadyRolledList.Remove(ResultBox.Text)
                        End If
                    End If

                    For Each myForm As Form In My.Application.OpenForms
                        If myForm.Name = FavoritesList.Name Then
                            With FavoritesList
                                If File.Exists(ResultBox.Text) Then
                                    .ListBox1.Items.Add(Search.RemoveExtension(Search.GetFileName(ResultBox.Text))) ' partial path
                                Else
                                    .ListBox1.Items.Add(Search.GetFileName(ResultBox.Text)) ' partial path
                                End If
                                .ListBox1Names.Items.Add(ResultBox.Text) ' full path
                            End With
                        End If
                    Next
                End If
            Else ' current roll is favorite

                If My.Settings.FavoriteRollsList.Contains(ResultBox.Text) Then
                    My.Settings.FavoriteRollsList.Remove(ResultBox.Text)
                    NonFavoriteRolled() ' now is unfavorited

                    If RollOnceListFlag = True Then
                        RollOnceListHandling("Add")
                    End If

                    For Each myForm As Form In My.Application.OpenForms
                        If myForm.Name = FavoritesList.Name Then
                            With FavoritesList
                                If File.Exists(ResultBox.Text) Then
                                    .ListBox1.Items.Remove(Search.RemoveExtension(Search.GetFileName(ResultBox.Text))) ' partial path
                                Else
                                    .ListBox1.Items.Remove(Search.GetFileName(ResultBox.Text)) ' partial path
                                End If
                                .ListBox1Names.Items.Remove(ResultBox.Text) ' full path
                            End With
                        End If
                    Next
                End If
            End If
        End If
    End Sub

    Sub FavoriteRolled() ' shows when currently rolled item is a favorite
        FavoriteButton.Text = "Unfavorite"
        FavoriteButton.BackColor = Color.Gold
        Me.BackColor = Color.Gold
        FavoritePictureBox.Image = My.Resources.favorite_85px
    End Sub

    Sub NonFavoriteRolled() ' shows when currently rolled item is not a favorite
        FavoriteButton.Text = "Favorite"
        FavoriteButton.BackColor = SystemColors.Control
        Me.BackColor = SystemColors.Control
        FavoritePictureBox.Image = Nothing
    End Sub

    Private Sub IgnoreRollButton_Click(sender As Object, e As EventArgs) Handles IgnoreRollButton.Click
        If IO.File.Exists(ResultBox.Text) OrElse IO.Directory.Exists(ResultBox.Text) Then
            RollOnceListHandling("Remove")
        End If
    End Sub

    ''' <summary>
    ''' Clear the roll once list.
    ''' </summary>
    Private Sub ClearRollOnceListButton_Click(sender As Object, e As EventArgs) Handles ClearRollOnceListButton.Click
        If MsgBox("Are you sure you wish to erase your roll once list completely?", MsgBoxStyle.YesNo, "WARNING") = DialogResult.Yes Then
            If MsgBox("Are you really sure?", MsgBoxStyle.YesNo, "WARNING") = DialogResult.Yes Then
                If Not IsNothing(RollOnceList) AndAlso Not IsNothing(My.Settings.AlreadyRolledList) Then
                    RollOnceList.Clear()
                    My.Settings.AlreadyRolledList.Clear()

                    RecomputeTreesChangesWereMade = True ' recompute file tree(s)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Used for adding or removing an item from the roll once list
    ''' </summary>
    ''' <param name="Instructions">"Remove" or "Add" strings accepted.</param>
    Private Sub RollOnceListHandling(Instructions As String)
        If ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = False) OrElse _
            ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = True) Then
            If IsNothing(RollOnceList) Then
                RollOnceList = New ArrayList
            End If
            If IsNothing(My.Settings.AlreadyRolledList) Then
                My.Settings.AlreadyRolledList = New Specialized.StringCollection
            End If

            Dim flag1 As Boolean = False
            If Instructions = "Remove" Then ' remove from exclusion list
                If IO.File.Exists(ResultBox.Text) OrElse IO.Directory.Exists(ResultBox.Text) Then

                    If RollOnceList.Contains(ResultBox.Text) Then
                        RollOnceList.Remove(ResultBox.Text)
                    End If
                    If My.Settings.AlreadyRolledList.Contains(ResultBox.Text) Then
                        My.Settings.AlreadyRolledList.Remove(ResultBox.Text)
                    End If

                    ' add back to respective parent folder(s)
                    ' can't know which rows contained the file/folder so just regenerate after removing from exclusion list so it's re-added that way
                    If (FileFlag = True AndAlso IO.File.Exists(ResultBox.Text)) OrElse _
                        (FolderFlag = True AndAlso IO.Directory.Exists(ResultBox.Text)) Then ' in file only rolls, shortcuted folders sometimes appear as folders in results...
                        flag1 = True
                    End If

                    ' single folder option
                    If Not IsNothing(TempFileTree) AndAlso IO.File.Exists(ResultBox.Text) Then
                        If Not TempFileTree.Contains(ResultBox.Text) Then
                            TempFileTree.Add(ResultBox.Text)
                        End If
                    End If
                    If Not IsNothing(TempFolderTree) AndAlso IO.Directory.Exists(ResultBox.Text) Then
                        If Not TempFolderTree.Contains(ResultBox.Text) Then
                            TempFolderTree.Add(ResultBox.Text)
                        End If
                    End If

                    IgnoreRollButton.Enabled = False
                    IgnoreRollButton.Visible = True
                    ClearRollOnceListButton.Visible = True
                End If
            ElseIf Instructions = "Add" Then ' add to exclusion list
                If IO.File.Exists(ResultBox.Text) OrElse IO.Directory.Exists(ResultBox.Text) Then

                    If Not RollOnceList.Contains(ResultBox.Text) Then
                        RollOnceList.Add(ResultBox.Text)
                    End If
                    If Not My.Settings.AlreadyRolledList.Contains(ResultBox.Text) Then
                        My.Settings.AlreadyRolledList.Add(ResultBox.Text)
                    End If

                    ' remove from parent folder(s)
                    If FileFlag = True AndAlso IO.File.Exists(ResultBox.Text) Then
                        For k As Integer = 0 To 29
                            If Not IsNothing(DirectoryFileTree(k)) AndAlso DirectoryFileTree(k).Contains(ResultBox.Text) Then
                                DirectoryFileTree(k).Remove(ResultBox.Text)
                                If DirectoryFileTree(k).Count = 0 Then
                                    'DisabledLine(k) = True
                                    flag1 = True
                                End If
                            End If
                        Next
                    End If
                    If FolderFlag = True AndAlso IO.Directory.Exists(ResultBox.Text) Then
                        For k As Integer = 0 To 29
                            If Not IsNothing(DirectoryFolderTree(k)) AndAlso DirectoryFolderTree(k).Contains(ResultBox.Text) Then
                                DirectoryFolderTree(k).Remove(ResultBox.Text)
                                If DirectoryFolderTree(k).Count = 0 Then
                                    'DisabledLine(k) = True
                                    flag1 = True
                                End If
                            End If
                        Next
                    End If

                    ' single folder option
                    If Not IsNothing(TempFileTree) AndAlso Not TempFileTree.Count = 0 AndAlso IO.File.Exists(ResultBox.Text) Then
                        If TempFileTree.Contains(ResultBox.Text) Then
                            TempFileTree.Remove(ResultBox.Text)
                        End If
                    End If
                    If Not IsNothing(TempFolderTree) AndAlso Not TempFolderTree.Count = 0 AndAlso IO.Directory.Exists(ResultBox.Text) Then
                        If TempFolderTree.Contains(ResultBox.Text) Then
                            TempFolderTree.Remove(ResultBox.Text)
                        End If
                    End If

                    IgnoreRollButton.Enabled = True
                    IgnoreRollButton.Visible = True
                    ClearRollOnceListButton.Visible = True
                End If
            End If

            ' updates trees when a group or single parent row has lost all its items or has gone from 0 to 1 just now and so must be refreshed
            ' because the disable line mechanic is used to deny roll of the group/row, instead of changing the weight around
            If flag1 = True Then
                RecomputeTreesChangesWereMade = True ' recompute file tree(s)
            End If
        End If
    End Sub

    Private Sub RollButton_MouseDown(sender As Object, e As MouseEventArgs) Handles RollButton.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then ' right mouse button pressed
            If NoToolStripMenuItem3.Checked = True Then
                YesToolStripMenuItem3.PerformClick()
            ElseIf YesToolStripMenuItem3.Checked = True Then
                NoToolStripMenuItem3.PerformClick()
            End If
        End If
    End Sub

    Private Sub WhitelistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WhitelistToolStripMenuItem.Click
        If WhitelistToolStripMenuItem.Checked = False Then
            WhitelistToolStripMenuItem.Checked = True
            My.Settings.WhitelistEnabled = True
        ElseIf WhitelistToolStripMenuItem.Checked = True Then
            WhitelistToolStripMenuItem.Checked = False
            My.Settings.WhitelistEnabled = False
        End If

        RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub

    Private Sub BlacklistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BlacklistToolStripMenuItem.Click
        If BlacklistToolStripMenuItem.Checked = False Then
            BlacklistToolStripMenuItem.Checked = True
            My.Settings.BlacklistEnabled = True
        ElseIf BlacklistToolStripMenuItem.Checked = True Then
            BlacklistToolStripMenuItem.Checked = False
            My.Settings.BlacklistEnabled = False
        End If

        RecomputeTreesChangesWereMade = True ' recompute file tree(s)
    End Sub

    ' drag and drop folder(s) into single textbox (tests multiple but will insert first valid folder and break)
    Private Sub TextBox1_DragDrop(sender As Object, e As DragEventArgs) Handles TextBox1.DragDrop
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            Dim directories As String() = e.Data.GetData(DataFormats.FileDrop)

            For i As Integer = 0 To directories.Length - 1
                If Directory.Exists(directories(i)) Then
                    TextBox1.Text = directories(i)
                    My.Settings.SingleParentFolder = directories(i)
                    Exit For
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Related to the dragging of a folder for single folder option TextBox.
    ''' </summary>
    Private Sub TextBox1_DragEnter(sender As Object, e As DragEventArgs) Handles TextBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        If ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = False) OrElse _
          ((SpecifiedShortcutFolder = True OrElse shortcutCreationAllowed = False) AndAlso MultipleFoldersFlag = True AndAlso MultipleFoldersHandled = True) Then
            Search.Show()
        Else
            MsgBox("No files or folders are rollable!", MsgBoxStyle.Exclamation, "Warning")
        End If
    End Sub

    Private Sub FoldersExpandButton_MouseDown(sender As Object, e As MouseEventArgs) Handles FoldersExpandButton.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then ' right mouse button pressed
            If NoToolStripMenuItem1.Checked = True Then
                YesToolStripMenuItem1.PerformClick()
            ElseIf YesToolStripMenuItem1.Checked = True Then
                NoToolStripMenuItem1.PerformClick()
            End If
        End If
    End Sub

    Private Sub RandomFile_MouseDown(sender As Object, e As MouseEventArgs) Handles MyBase.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then ' right mouse button pressed
            Dim mousePos As Point = Me.PointToClient(Windows.Forms.Cursor.Position)
            If mousePos.X >= FoldersExpandButton.Location.X AndAlso mousePos.X <= FoldersExpandButton.Size.Width + FoldersExpandButton.Location.X AndAlso _
               mousePos.Y >= FoldersExpandButton.Location.Y AndAlso mousePos.Y <= FoldersExpandButton.Size.Height + FoldersExpandButton.Location.Y Then
                FoldersExpandButton_MouseDown(sender, e) ' toggle this button (enabled or disabled)
            End If
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        About.Show()
    End Sub

    ''' <summary>
    ''' Flag for roll rules being enabled or not. Disabled means all rolls are worth 1 always.
    ''' </summary>
    Dim RollRulesEnabled As Boolean = False
    ''' <summary>
    ''' Roll rules enabled.
    ''' </summary>
    Private Sub YesToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles YesToolStripMenuItem6.Click
        If Not YesToolStripMenuItem6.Checked = True Then
            RollRulesEnabled = True
            My.Settings.RollRulesEnabled = True
            NoToolStripMenuItem6.Checked = False
            YesToolStripMenuItem6.Checked = True

            RollsRichTextBox.Enabled = True
            Label4.Enabled = True
            TimesBox.Enabled = True
            Label3.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' Roll rules disabled.
    ''' </summary>
    Private Sub NoToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles NoToolStripMenuItem6.Click
        If Not NoToolStripMenuItem6.Checked = True Then
            RollRulesEnabled = False
            My.Settings.RollRulesEnabled = False
            NoToolStripMenuItem6.Checked = True
            YesToolStripMenuItem6.Checked = False

            RollsRichTextBox.Enabled = False
            Label4.Enabled = False
            TimesBox.Enabled = False
            Label3.Enabled = False
        End If
    End Sub

    ''' <summary>
    ''' Shows list of favorites.
    ''' </summary>
    Private Sub ShowFavoritesList() Handles FavoritesListButton.Click
        FavoritesList.Show()
    End Sub
End Class
