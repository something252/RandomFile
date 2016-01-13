Imports System.IO
Imports System.Windows.Forms.ListView
Imports System.ComponentModel
Imports System.Threading

Public Class Search
    ''' <summary>
    ''' Contains the currently displayed TreeNode, excluding searches.
    ''' </summary>
    Dim CurrentNodeDisplayed As TreeNode
    ''' <summary>
    ''' Contains the top most node of the TreeView.
    ''' </summary>
    Dim TopNodeVar As TreeNode
    ''' <summary>
    ''' Contains the current ListView visible directory.
    ''' </summary>
    Dim CurrentDirectory As String = ""
    ''' <summary>
    ''' Contains the previously visited directories.
    ''' </summary>
    Dim PreviousDirectory As New ArrayList
    ''' <summary>
    ''' Contains previously visited directories.
    ''' </summary>
    Dim ForwardDirectory As New ArrayList
    ''' <summary>
    ''' Holds the currently displayed search view.
    ''' </summary>
    Dim CurrentSearch As SearchResult
    ''' <summary>
    ''' Points to the currently visible ListView.
    ''' </summary>
    Dim CurrentVisibleListView As ListView
    ''' <summary>
    ''' Flag for whether or not search view is currently being displayed.
    ''' </summary>
    Dim SearchIsCurrentlyDisplayed As Boolean = False
    ''' <summary>
    ''' Points to this open form. (Search)
    ''' </summary>
    Shared ThisForm As Search
    ''' <summary>
    ''' Used to search with a user given string
    ''' </summary>
    Dim SearchThread As System.Threading.Thread
    ''' <summary>
    ''' Flag used to abort the Search Thread when true.
    ''' </summary>
    Dim SearchThreadAbort As Boolean = False
    ''' <summary>
    ''' State of thread being active or not.
    ''' </summary>
    Dim SearchThreadAlive As Boolean = False
    ''' <summary>
    ''' Used as SyncLock object for Search Thread.
    ''' </summary>
    Dim SearchLock As New Object
    ''' <summary>
    ''' Stores the icons of the currently display non-search ListView.
    ''' </summary>
    Dim ListViewMainIcons As New ImageList
    ''' <summary>
    ''' Used for adding to search ListView items.
    ''' </summary>
    Shared SearchCollectionBuffer As New ArrayList
    ''' <summary>
    ''' Used as SyncLock object for SearchCollectionBuffer.
    ''' </summary>
    Shared BufferLock As New Object
    ''' <summary>
    ''' SyncLock for updating ProgressBar.
    ''' </summary>
    Shared ProgressLock As New Object

    Public Sub RefreshButton_Click() Handles RefreshButton.Click
        SearchThreadAbort = True ' abort the search thread if active
        totalNodes = 0 ' reset total nodes count
        Me.SearchChangeTimer.Stop()
        Me.SearchUpdateTimer.Stop()
        Me.SearchProgress.Stop()
        SyncLock SearchLock
            SearchThreadAbort = False
        End SyncLock
        If Not IsNothing(CurrentSearch) Then
            CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
            CurrentSearch = Nothing ' stop using search results because they may be wrong after refreshing
        End If
        Me.ActiveControl = Nothing
        If Not IsNothing(PreviousDirectory) Then
            PreviousDirectory.Clear()
        End If
        BackButton.Enabled = False
        If Not IsNothing(ForwardDirectory) Then
            ForwardDirectory.Clear()
        End If
        ForwardButton.Enabled = False
        UpButton.Enabled = False
        If RandomFile.RecomputeTreesChangesWereMade = True Then
            RefreshTrees()
        End If
        TreeViewAll.Nodes.Clear()
        Dim node As TreeNode = TreeViewAll.Nodes.Add("Computer")
        node.Name = "root"
        TopNodeVar = node
        CurrentNodeDisplayed = TopNodeVar
        If startupFlag = False Then
            SaveSettings()
        End If
        GenerateTreeViewMain()
        DisplayTopLevel()
        ProgressBar1.Maximum = totalNodes
        ResetSearchText() ' refresh search box text
    End Sub

    Dim startupFlag As Boolean = True
    Private Sub Search_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        ListViewMain.SmallImageList = ListViewMainIcons
        ListViewMain.Sorting = SortOrder.Ascending
        RefreshButton_Click()
        CurrentVisibleListView = Me.ListViewMain
        ThisForm = Me

        startupFlag = False ' startup work is finished now
    End Sub

    Private Sub Search_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = My.Resources.icon
        With My.Settings
            If Not IsNothing(.SearchShowExtensions) AndAlso .SearchShowExtensions = True Then
                ExtensionsCheckBox.Checked = True
            Else
                ExtensionsCheckBox.Checked = False
            End If
            If Not IsNothing(.SearchFormSize) AndAlso _
                .SearchFormSize.Width >= Me.MinimumSize.Width AndAlso .SearchFormSize.Height > Me.MinimumSize.Height Then
                Me.Size = .SearchFormSize
            End If
            If Not IsNothing(.SearchFormPosition) AndAlso _
                (My.Settings.SearchFormPosition.X <> -1 AndAlso My.Settings.SearchFormPosition.Y <> -1) Then
                If RandomFile.FormVisible(.SearchFormPosition) OrElse _
                    RandomFile.FormVisible(New Point(.SearchFormPosition.X + Me.Size.Width, .SearchFormPosition.Y)) Then
                    Me.Location = .SearchFormPosition
                Else
                    Me.Location = New Point((Screen.PrimaryScreen.WorkingArea.Width / 2) - (Me.Size.Width / 2), (Screen.PrimaryScreen.WorkingArea.Height / 2) - (Me.Size.Height / 2))
                End If
            End If
        End With
    End Sub

    Private Sub RefreshTrees()
        With RandomFile
            If .RecomputeTreesChangesWereMade = True Then
                If ((.SpecifiedShortcutFolder = True OrElse .shortcutCreationAllowed = False) AndAlso .MultipleFoldersFlag = False) Then
                    .ParentFolderSingleFolderChanged = True ' recompute for single folder option
                    If Not System.IO.Directory.Exists(.TextBox1.Text) Then ' error in path given
                        'MsgBox("Please enter a valid directory and try again.", MsgBoxStyle.Exclamation, "Error")
                        Exit Sub
                    Else
                        My.Settings.SingleParentFolder = .TextBox1.Text ' update setting

                        If .ParentFolderSingleFolderChanged = True Then ' re-compute file/folder tree
                            Dim tmp As New ArrayList
                            tmp = FileTree.CreateFileList(.TextBox1.Text)
                            .TempFileTree = tmp.Clone()
                            tmp = FileTree.CreateFolderList(.TextBox1.Text)
                            .TempFolderTree = tmp.Clone()
                            .ParentFolderSingleFolderChanged = False
                        End If
                    End If
                ElseIf ((.SpecifiedShortcutFolder = True OrElse .shortcutCreationAllowed = False) AndAlso .MultipleFoldersFlag = True AndAlso .MultipleFoldersHandled = True) Then
                    .UpdateTrees()
                End If
                .RecomputeTreesChangesWereMade = False ' done recomputing
            End If
        End With
    End Sub

    Private Sub GenerateTreeViewMain()
        With RandomFile
            If ((.SpecifiedShortcutFolder = True OrElse .shortcutCreationAllowed = False) AndAlso .MultipleFoldersFlag = False) OrElse _
                ((.SpecifiedShortcutFolder = True OrElse .shortcutCreationAllowed = False) AndAlso .MultipleFoldersFlag = True AndAlso .MultipleFoldersHandled = True) Then
                'If IsNothing(.RollOnceList) Then
                '    .RollOnceList = New ArrayList
                'End If
                'If IsNothing(My.Settings.AlreadyRolledList) Then
                '    My.Settings.AlreadyRolledList = New Specialized.StringCollection
                'End If

                Dim DirectoryTree As Object

                Dim FileAndFolderFlag As Boolean = False
                If .MultipleFoldersFlag = True Then ' multiple folders option
                    If .DirectoryFileTree.Length > 0 AndAlso (.FileFlag = True AndAlso .FolderFlag = False) Then ' files only option
                        DirectoryTree = .DirectoryFileTree
                    ElseIf .DirectoryFolderTree.Length > 0 AndAlso (.FileFlag = False AndAlso .FolderFlag = True) Then ' folders only option
GoToFolderOption1:
                        DirectoryTree = .DirectoryFolderTree
                    ElseIf (.DirectoryFileTree.Length > 0 AndAlso .DirectoryFolderTree.Length > 0) AndAlso (.FileFlag = True AndAlso .FolderFlag = True) Then ' both
                        DirectoryTree = .DirectoryFileTree
                        FileAndFolderFlag = True
                    Else ' error
                        GoTo GoToSearchNotPossible
                    End If
                Else ' single folder option
                    If (.FileFlag = True AndAlso .FolderFlag = False) Then ' file only
                        DirectoryTree = .TempFileTree
                    ElseIf (.FileFlag = False AndAlso .FolderFlag = True) Then ' folder only
GoToFolderOption2:
                        DirectoryTree = .TempFolderTree
                    ElseIf (.FileFlag = True AndAlso .FolderFlag = True) Then ' both
                        DirectoryTree = .TempFileTree
                        FileAndFolderFlag = True
                    Else ' error
                        GoTo GoToSearchNotPossible
                    End If
                End If

                GenerateWork(DirectoryTree)

                If FileAndFolderFlag = True Then ' file and folder (both) option related
                    FileAndFolderFlag = False
                    If .MultipleFoldersFlag = True Then
                        GoTo GoToFolderOption1
                    Else
                        GoTo GoToFolderOption2
                    End If
                End If
            Else
GoToSearchNotPossible:
                MsgBox("No files or folders are rollable!", MsgBoxStyle.Exclamation, "Warning")
                Me.Close()
            End If
        End With
    End Sub

    ''' <summary>
    ''' Used to generate the TreeView based on ArrayList input containing full path strings.
    ''' </summary>
    Private Sub GenerateWork(ByRef DirectoryTree As Object)
        With RandomFile

            Dim PerformedCheck As ArrayList = New ArrayList ' holds groups (grouped rows) that have already populated the TreeView

            Dim listTemp As Object
            If RandomFile.MultipleFoldersFlag = True Then
                Dim tmp() As ArrayList = DirectoryTree
                listTemp = tmp
            Else
                Dim tmp(0) As ArrayList
                tmp(0) = DirectoryTree
                listTemp = tmp
            End If

            For k As Integer = 0 To listTemp.Length - 1
                If Not IsNothing(listTemp(k)) Then
                    If (RandomFile.MultipleFoldersFlag = False) OrElse (Not PerformedCheck.Contains(.FolderGroups(k)) OrElse .FolderGroups(k) = "") Then

                        Dim RootNode As TreeNode = TopNodeVar
                        For Each obj1 As String In listTemp(k)
                            Dim path() As String = Split(obj1, "\")
                            Dim CurrentPath As String = path(0) ' contains current directory path
                            Dim CurrNode As TreeNode
                            CurrNode = FindNode(RootNode, CurrentPath) ' find drive letter root node
                            If IsNothing(CurrNode) Then ' not found
                                CurrNode = RootNode.Nodes.Add(CurrentPath, path(0)) ' add node and set as current node
                                totalNodes += 1
                            End If

                            Dim c As Integer
                            For c = 1 To path.Length - 2 ' perform all except last node (which may be file or folder)
                                CurrentPath &= "\" & path(c)
                                Dim tmp As TreeNode = FindNode(CurrNode, CurrentPath) ' find if child node already exists
                                If IsNothing(tmp) Then ' not found
                                    CurrNode = CurrNode.Nodes.Add(CurrentPath, path(c)) ' add node and set as current node
                                    totalNodes += 1
                                Else
                                    CurrNode = tmp ' set as current node
                                End If
                            Next

                            CurrentPath &= "\" & path(c) ' last element is now added (file or folder)
                            Dim tmp2 As TreeNode = FindNode(CurrNode, CurrentPath) ' find if child node already exists
                            If IsNothing(tmp2) Then ' not found (doesn't exist already)
                                If (.FileFlag = True AndAlso .FolderFlag = False) Then ' file tree only is being used
                                    CurrNode.Nodes.Add(CurrentPath, RemoveExtension(path(c))) ' add last element which is a file without the extension in name
                                    totalNodes += 1
                                Else ' folder only or file and folder options
                                    If File.Exists(obj1) Then
                                        If ExtensionsCheckBox.Checked = False Then
                                            CurrNode.Nodes.Add(CurrentPath, RemoveExtension(path(c))) ' add last element which is a file without the extension in name
                                        Else
                                            CurrNode.Nodes.Add(CurrentPath, path(c)) ' add last element which is a file with the extension in name
                                        End If
                                        totalNodes += 1
                                    Else
                                        CurrNode.Nodes.Add(CurrentPath, path(c)) ' add last folder
                                        totalNodes += 1
                                    End If
                                End If
                            End If
                        Next

                        If RandomFile.MultipleFoldersFlag = True AndAlso (Not .FolderGroups(k) = "" AndAlso Not PerformedCheck.Contains(.FolderGroups(k))) Then
                            PerformedCheck.Add(.FolderGroups(k))
                        End If
                    End If
                End If
            Next
        End With
    End Sub

    ''' <summary>
    ''' Search for a node with the corresponding key value in the nodes of the given node, without searching sub-nodes.
    ''' </summary>
    ''' <param name="GivenNode">Nodes of tree node to be searched, except for the child nodes of those tree nodes.</param>
    ''' <param name="KeyValue">Given string to find first occurance of in the .Name value of the nodes.</param>
    Private Function FindNode(ByRef GivenNode As TreeNode, ByRef KeyValue As String) As TreeNode
        If Not IsNothing(KeyValue) AndAlso Not IsNothing(GivenNode) Then
            Dim LastKeyValueIndex As Integer = KeyValue.Length - 1
            Dim LastTestStringIndex As Integer
            Dim i As Integer
            For Each node As TreeNode In GivenNode.Nodes
                If node.Name.Length = KeyValue.Length Then
                    LastTestStringIndex = node.Name.Length - 1
                    For i = LastKeyValueIndex To 0 Step -1
                        If Not node.Name(LastTestStringIndex) = KeyValue(i) Then
                            GoTo continueForFindNode1 ' char did not match
                        End If
                        LastTestStringIndex -= 1
                    Next
                    Return node ' all chars matched / matching node found
                End If
continueForFindNode1:  ' continue for
            Next
        End If
        Return Nothing ' no matches found
    End Function

    ''' <summary>
    ''' Remove any colons from a given string
    ''' </summary>
    Private Function RemoveColons(ByRef Str As String) As String
        If Str.Length > 0 Then
            Dim newStr As String = ""
            For i As Integer = 0 To Str.Length - 1
                If Not Str(i) = ":" Then
                    newStr &= Str(i)
                End If
            Next
            Return newStr
        Else
            Return Str
        End If
    End Function

    ''' <summary>
    ''' Removes the extension from a string. (That was constructed by split on backslash)
    ''' </summary>
    Public Function RemoveExtension(ByRef Str As String) As String
        If ExtensionsCheckBox.Checked = False Then
            If Str.Length > 0 Then
                Dim index As Integer = -999 ' store index of last period
                For i As Integer = (Str.Length - 1) To 0 Step -1
                    If Str(i) = "." Then
                        index = i
                        Exit For ' extension found
                    ElseIf Str(i) = "\" Then
                        Exit For ' failed to find extension
                    End If
                Next
                If Not index = -999 Then
                    Dim newStr As String = ""
                    For i As Integer = 0 To index - 1
                        newStr &= Str(i)
                    Next
                    Return newStr
                End If
            End If
        End If
        Return Str
    End Function

    Private Sub SaveSettings()
        If ExtensionsCheckBox.Checked = True Then
            My.Settings.SearchShowExtensions = True
        Else
            My.Settings.SearchShowExtensions = False
        End If

        If Me.WindowState = FormWindowState.Minimized Then ' minimizing makes location -32000 coords
            Me.Opacity = 0.0
            Me.WindowState = FormWindowState.Normal
        End If
        My.Settings.SearchFormPosition = Me.Location
        My.Settings.SearchFormSize = Me.Size
    End Sub

    Private Sub Search_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        SaveSettings()

        SearchThreadAbort = True
        SyncLock SearchLock
            SearchThreadAbort = False
        End SyncLock
    End Sub

    ''' <summary>
    ''' Display the top level of the TreeView in the ListView.
    ''' </summary>
    Private Sub DisplayTopLevel()
        If Not IsNothing(TopNodeVar) Then
            UpButton.Enabled = False
            CurrentDirectory = "" 'TreeViewAll.TopNode.Name & "\"
            AddressTextBox.Text = "Computer"
            ListViewMain.BeginUpdate() ' pause drawing of control
            ListViewMain.Items.Clear()
            CurrentNodeDisplayed = TopNodeVar
            'ListViewMainIcons.Images.Clear()
            For Each node As TreeNode In TopNodeVar.Nodes
                Dim item = ListViewMain.Items.Add(node.Text)
                item.Name = node.Name
                GetIcon(node, item, ListViewMainIcons) ' get and assign icon
                PopulateItemColumns(item, item.Name)
            Next
            ListViewMain.EndUpdate() ' resume drawing of control
        End If
    End Sub

    Private Overloads Sub DisplayDirectory(ByRef node1 As TreeNode)
        If Not IsNothing(node1) Then
            Dim currentNode As TreeNode = node1
            If Not IsNothing(currentNode) Then
                If currentNode.Name = TopNodeVar.Name Then
                    UpButton.Enabled = False
                Else
                    UpButton.Enabled = True
                End If
                ListViewMain.BeginUpdate() ' pause drawing of control
                ListViewMain.Items.Clear()
                CurrentNodeDisplayed = currentNode
                'ListViewMainIcons.Images.Clear()
                For Each node As TreeNode In currentNode.Nodes
                    Dim item = ListViewMain.Items.Add(node.Text)
                    item.Name = node.Name
                    GetIcon(node, item, ListViewMainIcons) ' get and assign icon
                    PopulateItemColumns(item, item.Name)
                Next
                ListViewMain.EndUpdate() ' resume drawing of control
            End If
        End If
    End Sub

    Private Overloads Sub DisplayDirectory(ByRef node1() As TreeNode)
        If node1.Length > 0 Then
            Dim currentNode As TreeNode = node1(0)
            If Not IsNothing(currentNode) Then
                If currentNode.Name = TopNodeVar.Name Then
                    UpButton.Enabled = False
                Else
                    UpButton.Enabled = True
                End If
                ListViewMain.BeginUpdate() ' pause drawing of control
                ListViewMain.Items.Clear()
                CurrentNodeDisplayed = currentNode
                'ListViewMainIcons.Images.Clear()
                For Each node As TreeNode In currentNode.Nodes
                    Dim item = ListViewMain.Items.Add(node.Text)
                    item.Name = node.Name
                    GetIcon(node, item, ListViewMainIcons) ' get and assign icon
                    PopulateItemColumns(item, item.Name)
                Next
                ListViewMain.EndUpdate() ' resume drawing of control
            End If
        End If
    End Sub

    Private Overloads Sub DisplayDirectory(ByRef path As String)
        Dim directories() As String = Split(path, "\")
        Dim CurrNode As TreeNode
        If directories.Length > 0 Then
            If Not directories(0) = "" Then
                Dim CurrentPath As String = directories(0)
                CurrNode = FindNode(TopNodeVar, CurrentPath) ' find drive letter root node
                If Not IsNothing(CurrNode) Then
                    For i As Integer = 1 To directories.Length - 1
                        If Not directories(i) = "" Then
                            CurrentPath &= "\" & directories(i)
                            CurrNode = FindNode(CurrNode, CurrentPath)
                        End If
                    Next
                    ListViewMain.BeginUpdate() ' pause drawing of control
                    ListViewMain.Items.Clear()
                    CurrentNodeDisplayed = CurrNode
                    'ListViewMainIcons.Images.Clear()
                    For Each node As TreeNode In CurrNode.Nodes
                        Dim item = ListViewMain.Items.Add(node.Text)
                        item.Name = node.Name
                        GetIcon(node, item, ListViewMainIcons) ' get and assign icon
                        PopulateItemColumns(item, item.Name)
                    Next
                    ListViewMain.EndUpdate() ' resume drawing of control
                Else
                    MsgBox("Node error: 2947437", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
            ElseIf directories(0) = "" Then
                ListViewMain.BeginUpdate() ' pause drawing of control
                ListViewMain.Items.Clear()
                CurrentNodeDisplayed = TopNodeVar
                'ListViewMainIcons.Images.Clear()
                For Each node As TreeNode In TopNodeVar.Nodes
                    Dim item = ListViewMain.Items.Add(node.Text)
                    item.Name = node.Name
                    GetIcon(node, item, ListViewMainIcons) ' get and assign icon
                    PopulateItemColumns(item, item.Name)
                Next
                ListViewMain.EndUpdate() ' resume drawing of control
            End If
            If CurrentNodeDisplayed.Name = TopNodeVar.Name Then
                UpButton.Enabled = False
            Else
                UpButton.Enabled = True
            End If
        End If
    End Sub

    ''' <summary>
    ''' Get and assign the icon for a corresponding ListViewItem.
    ''' </summary>
    ''' <param name="node">Node corresponding to ListViewItem.</param>
    ''' <param name="item">ListViewItem getting assigned the icon.</param>
    ''' <param name="ImageList1">ImageList of item's ListView.</param>
    Private Sub GetIcon(ByRef node As TreeNode, ByRef item As ListViewItem, ByRef ImageList1 As ImageList)
        Try
            If node.Nodes.Count > 0 OrElse Directory.Exists(node.Name) Then ' is a folder (directory check is for empty folders)
                If Not ImageList1.Images.ContainsKey("*F*") Then
                    'ImageList1.Images.Add("*F*", IconReader.GetFolderIcon(IconReader.IconSize.Small, IconReader.FolderType.Open, shfi))
                    ImageList1.Images.Add("*F*", My.Resources.FolderIcon)
                End If
                item.ImageKey = "*F*"
            Else ' is a file
                Dim extension As String = GetFileType(node.Name)
                If Not extension = "lnk" Then
                    If Not ImageList1.Images.ContainsKey(extension) Then
                        ImageList1.Images.Add(extension, Drawing.Icon.ExtractAssociatedIcon(node.Name))
                        'ImageList1.Images.Add(extension, IconReader.ExtractIconFromFileEx(node.Name, IconReader.IconSize.Small, CheckIfShortcut(node.Name), shfi))
                    End If
                    item.ImageKey = extension
                Else
                    If Not ImageList1.Images.ContainsKey(node.Name) Then
                        ImageList1.Images.Add(node.Name, Drawing.Icon.ExtractAssociatedIcon(node.Name))
                        'ImageList1.Images.Add(extension, IconReader.ExtractIconFromFileEx(node.Name, IconReader.IconSize.Small, CheckIfShortcut(node.Name), shfi))
                    End If
                    item.ImageKey = node.Name
                End If
            End If
        Catch ex As Exception
            ' Dim ico As Icon = Drawing.Icon.ExtractAssociatedIcon(fileName)
            'imageListName.Images.Add(ext, ico)
        End Try
    End Sub

    ''' <summary>
    ''' Populate the columns of an item.
    ''' </summary>
    ''' <param name="item">Item of a ListView</param>
    ''' <param name="FullPath">Full path that the item represents.</param>
    Private Sub PopulateItemColumns(ByRef item As ListViewItem, ByRef FullPath As String)
        Dim isFile As Boolean = False
        If File.Exists(FullPath) Then
            isFile = True
        ElseIf Not Directory.Exists(FullPath) Then
            Exit Sub
        End If
        Dim infoReader As System.IO.FileInfo
        infoReader = My.Computer.FileSystem.GetFileInfo(FullPath)

        ' add date modified column
        item.SubItems.Add(CStr(infoReader.LastWriteTime))

        ' add extension/type column
        If isFile = False Then
            item.SubItems.Add("Directory")
        Else
            If infoReader.Extension.Length > 0 Then
                item.SubItems.Add(infoReader.Extension.Remove(0, 1)) ' remove period
            Else
                item.SubItems.Add(infoReader.Extension)
            End If
        End If

        ' add size column
        If isFile = True Then
            If Not IsNothing(infoReader.Length) AndAlso infoReader.Length > 0L Then
                item.SubItems.Add(CStr(Math.Round((infoReader.Length / 1024.0), 0)) & " KB")
            Else
                item.SubItems.Add("")
            End If
        Else
            item.SubItems.Add("")
        End If

    End Sub

    Private Sub ListViewMain_ItemActivate(sender As Object, e As EventArgs) Handles ListViewMain.ItemActivate
        If ListViewMain.SelectedItems.Count > 0 Then
            Dim CurrentPath As String = ListViewMain.SelectedItems.Item(0).Name
            Dim currentNode(0) As TreeNode
            currentNode(0) = TopNodeVar
            Try
                currentNode = currentNode(0).Nodes.Find(CurrentPath, True)
            Catch
                MsgBox("Error 135135623", MsgBoxStyle.Critical, "Error")
                Exit Sub
            End Try
            If Not IsNothing(currentNode) Then
                If Directory.Exists(CurrentPath) Then ' there can't be a file of the same full name as a folder in windows
                    If CurrentPath.Length > 0 AndAlso Not CurrentPath(CurrentPath.Length - 1) = "\" Then
                        CurrentPath &= "\"
                    End If
                    If SearchIsCurrentlyDisplayed = True Then
                        CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                        NewPreviousDirectory(CurrentSearch)
                    Else ' non-search view currently displayed
                        NewPreviousDirectory(CurrentDirectory)
                    End If
                    ForwardDirectory.Clear()
                    ForwardButton.Enabled = False
                    CurrentDirectory = CurrentPath
                    TreeViewAll.SelectedNode = currentNode(0)
                    AddressTextBox.Text = CurrentDirectory

                    ResetSearchText() ' refresh search box text
                    DisplayDirectory(currentNode)
                    If currentNode(0).Name = TopNodeVar.Name Then
                        UpButton.Enabled = False
                    Else
                        UpButton.Enabled = True
                    End If
                ElseIf File.Exists(CurrentPath) Then
                    RandomFile.OpenFile(CurrentPath)
                End If
            Else
                MsgBox("File or folder no longer exists!", MsgBoxStyle.Critical, "Warning")
            End If
        End If
    End Sub

    Private Sub TreeViewAll_NodeMouseDoubleClick(sender As Object, e As TreeNodeMouseClickEventArgs) Handles TreeViewAll.NodeMouseDoubleClick
        If Not IsNothing(TreeViewAll.SelectedNode) Then
            If IsNothing(TreeViewAll.SelectedNode.Parent) Then ' root double clicked
                If Not CurrentNodeDisplayed.Name = TreeViewAll.SelectedNode.Name Then ' if not already the same directory 
                    If SearchIsCurrentlyDisplayed = True Then
                        CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                        NewPreviousDirectory(CurrentSearch)
                    Else ' non-search view currently displayed
                        NewPreviousDirectory(CurrentDirectory)
                    End If
                    ForwardDirectory.Clear()
                    ForwardButton.Enabled = False
                    DisplayTopLevel() ' UpButton.Enabled = False
                    ResetSearchText() ' refresh search box text
                End If
            Else
                Dim ItemFullPath As String = TreeViewAll.SelectedNode.Name
                If Directory.Exists(ItemFullPath) Then
                    Dim CurrentDirectoryTemp As String = CurrentDirectory
                    If CurrentDirectoryTemp.Length > 0 AndAlso CurrentDirectoryTemp(CurrentDirectoryTemp.Length - 1) = "\" Then
                        CurrentDirectoryTemp = CurrentDirectoryTemp.Remove(CurrentDirectoryTemp.Length - 1, 1)
                    End If
                    Dim ItemFullPathTemp As String = ItemFullPath
                    If ItemFullPathTemp.Length > 0 AndAlso ItemFullPathTemp(ItemFullPathTemp.Length - 1) = "\" Then
                        ItemFullPathTemp = ItemFullPathTemp.Remove(ItemFullPathTemp.Length - 1, 1)
                    End If

                    If Not LCase(CurrentDirectoryTemp) = LCase(ItemFullPathTemp) Then ' if not already the same directory 
                        DisplayDirectory(TreeViewAll.SelectedNode)
                        If SearchIsCurrentlyDisplayed = True Then
                            CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                            NewPreviousDirectory(CurrentSearch)
                        Else ' non-search view currently displayed
                            NewPreviousDirectory(CurrentDirectory)
                        End If
                        ForwardDirectory.Clear()
                        ForwardButton.Enabled = False
                        CurrentDirectory = ItemFullPath
                        If CurrentDirectory.Length > 0 AndAlso Not CurrentDirectory(CurrentDirectory.Length - 1) = "\" Then
                            CurrentDirectory &= "\"
                        End If
                    End If

                    AddressTextBox.Text = CurrentDirectory
                    UpButton.Enabled = True
                    ResetSearchText() ' refresh search box text

                ElseIf File.Exists(ItemFullPath) Then
                    RandomFile.OpenFile(ItemFullPath)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Checks if the given path string exists as a node in the TreeView.
    ''' </summary>
    ''' <param name="path">Full path of a directory.</param>
    ''' <returns>True if exists and false otherwise.</returns>
    Private Function CheckDirectoryExists(ByRef path As String) As Boolean
        Dim directories() As String = Split(path, "\")
        Dim CurrentPath As String
        Dim CurrNode As TreeNode

        If directories.Length > 0 Then
            If Not directories(0) = "" Then
                CurrNode = FindNode(TopNodeVar, directories(0)) ' find drive letter root node
                CurrentPath = directories(0)
                If Not IsNothing(CurrNode) Then
                    For i As Integer = 1 To directories.Length - 1
                        If Not directories(i) = "" Then
                            If Not IsNothing(CurrNode) Then
                                CurrentPath &= "\" & directories(i)
                                CurrNode = FindNode(CurrNode, CurrentPath)
                            Else
                                Return False
                            End If
                        End If
                    Next
                    If Not IsNothing(CurrNode) Then
                        Return True
                    Else
                        Return False
                    End If
                Else
                    Return False
                End If
            Else
                Return True ' top node
            End If
        Else
            Return False
        End If

        Return False
    End Function

    ' ListView related
    Private Sub OpenContainingFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenContainingFolderToolStripMenuItem.Click
        With RandomFile
            If Not IsNothing(ListViewMain.SelectedItems) AndAlso ListViewMain.SelectedItems.Count > 0 Then
                .ContainingButton(ListViewMain.SelectedItems.Item(0).Name)
            End If
        End With
    End Sub

    ' TreeView related
    Private Sub OpenContainingFolderToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles OpenContainingFolderToolStripMenuItem2.Click
        If Not IsNothing(TreeViewAll.SelectedNode) Then
            Dim Str As ArrayList = New ArrayList
            Dim CurrentNode As TreeNode = TreeViewAll.SelectedNode

            If CurrentNode.Name = "root" Then
                Try
                    Call Shell("explorer /select," & """" & Environment.GetFolderPath(Environment.SpecialFolder.MyComputer) & """", AppWinStyle.NormalFocus) ' select in containing folder
                Catch
                End Try
            ElseIf Not IsNothing(CurrentNode) Then
                If Directory.Exists(CurrentNode.Name) OrElse File.Exists(CurrentNode.Name) Then
                    With RandomFile
                        .ContainingButton(CurrentNode.Name)
                    End With
                End If
            End If
        End If
    End Sub

    Private Sub NewPreviousDirectory(Dir As Object, Optional forwardButtonClick As Boolean = False) ' Dir type = string or SearchResult
        If Not IsNothing(Dir) Then
            If forwardButtonClick = False Then ' directory was navigated to by user in listview doubleclick or treeview doubleclick (not forward button click)

                PreviousDirectory.Add(Dir)

            ElseIf forwardButtonClick = True Then ' forward button was clicked
                If ForwardDirectory.Count > 0 Then

                    PreviousDirectory.Add(Dir) ' pop start
                    ForwardDirectory.RemoveAt(ForwardDirectory.Count - 1) ' pop end
                Else
                    MsgBox("Error 23623762", MsgBoxStyle.Critical, "Error")
                    Exit Sub
                End If
            End If

            BackButton.Enabled = True
            If ForwardDirectory.Count = 0 Then
                If ForwardButton.Focused = True Then
                    ForwardButton.Enabled = False
                    Me.ActiveControl = Nothing ' focus something nuetral
                Else
                    ForwardButton.Enabled = False
                End If
            Else
                ForwardButton.Enabled = True
            End If
        End If
    End Sub

    Dim backButtonFlag1 As Boolean = False
    Private Sub NewForwardDirectory(Dir As Object) ' Dir type = string or SearchResult
        If Not IsNothing(PreviousDirectory) AndAlso PreviousDirectory.Count > 0 Then
            If Not IsNothing(Dir) Then

                If backButtonFlag1 = False Then
                    ForwardDirectory.Add(Dir) ' pop start
                Else
                    CurrentSearch.Dispose()
                    CurrentSearch = Nothing
                End If

                PreviousDirectory.RemoveAt(PreviousDirectory.Count - 1) ' pop end

            End If

            ForwardButton.Enabled = True
            If PreviousDirectory.Count = 0 Then
                If BackButton.Focused = True Then
                    BackButton.Enabled = False
                    Me.ActiveControl = Nothing ' focus something nuetral
                Else
                    BackButton.Enabled = False
                End If
            Else
                BackButton.Enabled = True
            End If
        End If
    End Sub

    Private Sub BackButton_Click() Handles BackButton.Click
        If Not IsNothing(PreviousDirectory) AndAlso PreviousDirectory.Count > 0 Then

            If TypeOf (PreviousDirectory(PreviousDirectory.Count - 1)) Is SearchResult Then ' search view
                Dim CurrentSearchTemp As SearchResult = PreviousDirectory(PreviousDirectory.Count - 1)

                If SearchIsCurrentlyDisplayed = True Then
                    CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                    NewForwardDirectory(CurrentSearch)
                Else ' non-search view currently displayed
                    NewForwardDirectory(CurrentDirectory)
                End If

                CurrentSearch = CurrentSearchTemp
                CurrentSearch.Show() ' display ' sets SearchIsCurrentlyDisplayed = True
                CurrentDirectory = "Search Results"
                UpButton.Enabled = False
            Else ' string
                Dim String1 As String = CStr(PreviousDirectory(PreviousDirectory.Count - 1))

                If SearchIsCurrentlyDisplayed = True Then
                    CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                    NewForwardDirectory(CurrentSearch)
                Else ' non-search view currently displayed
                    NewForwardDirectory(CurrentDirectory)
                End If

                If String1 = "" Then ' top level
                    DisplayTopLevel()
                Else
                    DisplayDirectory(String1) ' display
                    CurrentDirectory = String1 ' new current directory
                End If

                If Not CurrentDirectory = "" Then
                    AddressTextBox.Text = CurrentDirectory
                    UpButton.Enabled = True
                Else
                    AddressTextBox.Text = "Computer"
                    UpButton.Enabled = False
                End If
                ResetSearchText() ' refresh search box text

            End If
        End If
    End Sub

    Private Sub ForwardButton_Click() Handles ForwardButton.Click
        If Not IsNothing(ForwardDirectory) AndAlso ForwardDirectory.Count > 0 Then

            If TypeOf (ForwardDirectory(ForwardDirectory.Count - 1)) Is SearchResult Then ' search view
                Dim CurrentSearchTemp As SearchResult = ForwardDirectory(ForwardDirectory.Count - 1)

                If SearchIsCurrentlyDisplayed = True Then
                    CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                    NewPreviousDirectory(CurrentSearch, True)
                Else ' non-search view currently displayed
                    NewPreviousDirectory(CurrentDirectory, True)
                End If

                CurrentSearch = CurrentSearchTemp
                CurrentSearch.Show() ' display ' sets SearchIsCurrentlyDisplayed = True
                CurrentDirectory = "Search Results"
                UpButton.Enabled = False
            Else ' ForwardDirectory(ForwardDirectory.Count - 1) = STRING type
                Dim String1 As String = CStr(ForwardDirectory(ForwardDirectory.Count - 1))

                If SearchIsCurrentlyDisplayed = True Then
                    CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                    NewPreviousDirectory(CurrentSearch, True)
                Else ' non-search view currently displayed
                    NewPreviousDirectory(CurrentDirectory, True)
                End If

                If String1 = "" Then ' top level
                    DisplayTopLevel()
                Else
                    DisplayDirectory(String1) ' display
                    CurrentDirectory = String1 ' new current directory
                End If

                If Not CurrentDirectory = "" Then
                    AddressTextBox.Text = CurrentDirectory
                    UpButton.Enabled = True
                Else
                    AddressTextBox.Text = "Computer"
                    UpButton.Enabled = False
                End If
                ResetSearchText() ' refresh search box text

            End If
        End If
    End Sub

    ''' <summary>
    ''' Can only up button (go up one level) when viewing a normal directory in the main ListView.
    ''' </summary>
    Private Sub UpButton_Click(sender As Object, e As EventArgs) Handles UpButton.Click
        Dim Directories() As String = Split(CurrentDirectory, "\")
        Dim NewStr As String = ""
        Dim tally As Integer = 0
        If Directories.Length > 0 Then
            If Directories(Directories.Length - 1) = "" Then
                ReDim Preserve Directories(Directories.Length - 2) ' remove leading blank string
            End If
            For i As Integer = 0 To Directories.Length - 2
                NewStr &= Directories(i) & "\"
                tally += 1
            Next
            If tally = 0 Then
                ' now at top level
                NewPreviousDirectory(CurrentDirectory)
                DisplayTopLevel() ' CurrentDirectory = ""
                UpButton.Enabled = False
                Me.ActiveControl = Nothing  ' button was disabled while focused
                AddressTextBox.Text = "Computer"
            Else
                NewPreviousDirectory(CurrentDirectory)
                CurrentDirectory = NewStr
                DisplayDirectory(NewStr)
                UpButton.Enabled = True
                AddressTextBox.Text = NewStr
            End If
            ForwardDirectory.Clear()
            ForwardButton.Enabled = False
            ResetSearchText() ' refresh search box text
        End If
    End Sub

    Private Sub AddressTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles AddressTextBox.KeyDown
        If e.KeyCode = 13 Then ' enter key pressed
            If LCase(AddressTextBox.Text) = "computer" OrElse AddressTextBox.Text = "" Then
                If Not CurrentDirectory = AddressTextBox.Text AndAlso Not (CurrentDirectory = "" AndAlso (AddressTextBox.Text = "" OrElse LCase(AddressTextBox.Text) = "computer")) Then ' if already the same directory 
                    If SearchIsCurrentlyDisplayed = True Then
                        CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                        NewPreviousDirectory(CurrentSearch)
                    Else ' non-search view currently displayed
                        NewPreviousDirectory(CurrentDirectory)
                    End If
                    ForwardDirectory.Clear()
                    ForwardButton.Enabled = False
                End If

                DisplayTopLevel() ' sets CurrentDirectory = ""
                UpButton.Enabled = False
            ElseIf CheckDirectoryExists(AddressTextBox.Text) = True Then
                Dim CurrentDirectoryTemp As String = CurrentDirectory
                If CurrentDirectoryTemp.Length > 0 AndAlso CurrentDirectoryTemp(CurrentDirectoryTemp.Length - 1) = "\" Then
                    CurrentDirectoryTemp = CurrentDirectoryTemp.Remove(CurrentDirectoryTemp.Length - 1, 1)
                End If
                Dim AddressBar As String = AddressTextBox.Text
                If AddressBar.Length > 0 AndAlso AddressBar(AddressBar.Length - 1) = "\" Then
                    AddressBar = AddressBar.Remove(AddressBar.Length - 1, 1)
                End If

                If Not LCase(CurrentDirectoryTemp) = LCase(AddressBar) Then ' if not already the same directory 
                    If SearchIsCurrentlyDisplayed = True Then
                        CurrentSearch.Hide() ' sets SearchIsCurrentlyDisplayed = False
                        NewPreviousDirectory(CurrentSearch)
                    Else ' non-search view currently displayed
                        NewPreviousDirectory(CurrentDirectoryTemp)
                    End If

                    DisplayDirectory(AddressBar)
                    CurrentDirectory = AddressBar
                    UpButton.Enabled = True
                    ForwardDirectory.Clear()
                    ForwardButton.Enabled = False

                    ResetSearchText() ' refresh search box text
                End If
            Else
                MsgBox("Can't find " & """" & AddressTextBox.Text & """" & ". Check the spelling and try again.", MsgBoxStyle.Critical, "Warning")
            End If
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub SearchTextBox_Enter(sender As Object, e As EventArgs) Handles SearchTextBox.Enter
        If SearchTextBox.Focused = True AndAlso SearchTextBox.ForeColor = Color.DarkGray AndAlso SearchTextBox_TextChanged_Lock = False Then
            SearchTextBox_TextChanged_Lock = True
            SearchTextBox.Text = ""
            SearchTextBox_TextChanged_Lock = False
            SearchTextBox.ForeColor = SystemColors.WindowText
        End If
    End Sub

    ''' <summary>
    ''' Changes the SearchTextBox.Text to the cosmetic version that shows "Search " and "Current Directory Name" with changed color also.
    ''' </summary>
    Private Sub SearchTextBox_Leave() Handles SearchTextBox.Leave
        If SearchTextBox.Focused = False AndAlso (SearchTextBox.Text = "") Then
            ResetSearchText()
        End If
    End Sub

    Private Sub ResetSearchText()
        SearchTextBox_TextChanged_Lock = True
        SearchTextBox.Text = "  Search "
        Dim Str() As String = Split(AddressTextBox.Text, "\")
        If Str.Length > 0 Then
            If Not Str(Str.Length - 1) = "" Then
                SearchTextBox.Text &= Str(Str.Length - 1)
            ElseIf Str.Length >= 2 AndAlso Not Str(Str.Length - 2) = "" Then
                SearchTextBox.Text &= Str(Str.Length - 2)
            ElseIf LCase(Str(0)) = "computer" Then
                SearchTextBox.Text &= "Computer"
            End If
        End If
        SearchTextBox.ForeColor = Color.DarkGray
        SearchTextBox_TextChanged_Lock = False
    End Sub

    ''' <summary>
    ''' Get name of a given full path including its extension if present.
    ''' </summary>
    Public Function GetFileName(ByRef Str As String) As String
        Dim count As Integer = 0
        For c As Integer = Str.Length - 1 To 0 Step -1
            If Str(c) = "\"c Then
                If Not c = 0 Then ' doesn't have slash as last char (Not Str.Length - c = Str.Length)
                    Return Str.Substring(c + 1, count)
                End If
            End If
            count += 1
        Next
        Return Str
    End Function

    ''' <summary>
    ''' Checks if a full path is a shortcut, or basically if it has the .lnk extension.
    ''' </summary>
    ''' <param name="str">Full path string</param>
    Private Function CheckIfShortcut(ByRef Str As String) As Boolean ' checks if a arrow should be put on an icon's image (if shortcut)
        Dim i As Integer = Str.Length - 1
        If i - 3 >= 0 Then
            If Str(i) = "k"c Then
                If Str(i - 1) = "n"c Then
                    If Str(i - 2) = "l"c Then
                        If Str(i - 3) = "."c Then
                            Return True ' .lnk
                        End If
                    End If
                End If
            End If
        End If
        Return False
    End Function

    Private Sub SearchTextBox_KeyDown(sender As Object, e As KeyEventArgs) Handles SearchTextBox.KeyDown
        If e.KeyCode = 13 Then ' enter key pressed
            ' search
            SearchTextBox_TextChanged()
            e.SuppressKeyPress = True
        End If
    End Sub

    Dim SearchTextBox_TextChanged_Lock As Boolean = False ' prevent programmatical text changing to trigger the text changed event (without using RemoveHandler)
    Dim SearchCount As Integer = 0 ' tally of the number of searches performed
    Private Sub SearchTextBox_TextChanged() Handles SearchTextBox.TextChanged
        If SearchTextBox_TextChanged_Lock = False Then

            If SearchTextBox.Focused = True Then
                If Not SearchTextBox.Text = "" Then ' not blank and only changed by user
                    If SearchIsCurrentlyDisplayed = False Then ' first time entering into search listview (not textchanged or other repeated use of search right now)
                        SearchThreadAbort = True ' abort the thread if active
                        SyncLock SearchLock
                            SearchThreadAbort = False
                        End SyncLock

                        If Not CurrentDirectory = "Search Results" Then
                            NewPreviousDirectory(CurrentDirectory)
                        End If

                        ForwardDirectory.Clear()
                        ForwardButton.Enabled = False
                        UpButton.Enabled = False
                        CurrentDirectory = "Search Results"
                        CurrentSearch = New SearchResult ' create new search ListView custom class
                        CurrentSearch.Show(True)
                    End If

                    ' start new search
                    ThisForm.SearchChangeTimer.Stop()
                    ThisForm.SearchChangeTimer.Start()

                Else ' stop showing search ListView
                    ThisForm.SearchChangeTimer.Stop()
                    backButtonFlag1 = True
                    BackButton_Click() ' hide()
                    backButtonFlag1 = False
                    ForwardDirectory.Clear()
                    ForwardButton.Enabled = False
                    SearchTextBox_TextChanged_Lock = True
                    SearchTextBox.Text = ""
                    SearchTextBox_TextChanged_Lock = False
                    SearchTextBox.ForeColor = SystemColors.WindowText
                    SearchTextBox.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub SearchChangeTimer_Tick() Handles SearchChangeTimer.Tick
        ThisForm.SearchChangeTimer.Stop()

        ' start new search
        SearchThreadAbort = True ' abort the thread if active
        SyncLock SearchLock
            SearchThreadAbort = False

            If Not IsNothing(CurrentSearch) Then
                SearchThread = New System.Threading.Thread(Sub() SearchThreadWork())
                SearchThread.IsBackground = True
                SyncLock ProgressLock
                    progress1 = 0 ' reset progress
                    ProgressBar1.Visible = True
                    ThisForm.SearchProgress.Start()
                End SyncLock
                ThisForm.SearchUpdateTimer.Start()
                ThisForm.SearchThread.Start()
            End If
        End SyncLock
    End Sub

    Shared progress1 As Integer = 0, totalNodes As Integer = 0
    ''' <summary>
    ''' Updates the progress bar for an in progress search.
    ''' </summary>
    Private Sub ProgressTimer_Tick(sender As Object, e As EventArgs) Handles SearchProgress.Tick
        SyncLock ProgressLock
            ProgressBar1.Value = progress1  ' update progress
        End SyncLock
    End Sub

    ''' <summary>
    ''' Searching thread related work.
    ''' </summary>
    Private Sub SearchThreadWork()
        SyncLock SearchLock
            SearchThreadAlive = True

            Control.CheckForIllegalCrossThreadCalls = False
            AddressTextBox.Text = "Search Results"
            CurrentSearch.SearchBar = SearchTextBox.Text
            Dim SearchText As String = LCase(SearchTextBox.Text) ' case insensitive search

            ' check if new search text already exists in shown search ListView items
            If Not IsNothing(CurrentSearch.ListViewVar.Items) AndAlso CurrentSearch.ListViewVar.Items.Count > 0 Then
                CurrentSearch.ListViewVar.BeginUpdate() ' pause drawing of control

                CurrentSearch.ListViewVar.Items.Clear()
                ListViewMainIcons.Images.Clear()

                CurrentSearch.ListViewVar.EndUpdate() ' resume drawing of control
            End If

            If SearchThreadAbort = False Then
                'ThisForm.SearchUpdateTimer.Start()
                SearchSubstring(SearchText, CurrentSearch)
                SyncLock BufferLock
                    ThisForm.SearchUpdateTimer.Stop()
                    If SearchCollectionBuffer.Count > 0 Then
                        UpdateSearchCollection() ' update any more left over items not yet added
                    End If
                End SyncLock
            End If
            ThisForm.SearchProgress.Stop()
            SyncLock ProgressLock
                progress1 = 0 ' reset progress
                ProgressBar1.Visible = False
                ProgressBar1.Value = 0
            End SyncLock

            SearchThreadAlive = False
        End SyncLock
    End Sub

    ''' <summary>
    ''' Search for a substring match in node names and add them to a new SearchResult view. (Uses recursion)
    ''' </summary>
    ''' <param name="SearchString">Substring to be searched for.</param>
    ''' <param name="NextNode">Used exclusively for recursive calls to search child nodes. Called when a node contains child nodes.</param>
    Private Sub SearchSubstring(ByVal SearchString As String, ByRef SearchView As SearchResult, Optional ByRef NextNode As TreeNode = Nothing)
        If SearchString.Length > 0 Then
            Dim CurrNode As TreeNode
            If IsNothing(NextNode) Then ' normal non-recursive call
                CurrNode = CurrentNodeDisplayed
            Else ' recursive call
                CurrNode = NextNode
            End If

            For Each node As TreeNode In CurrNode.Nodes
                If SearchThreadAbort = False Then
                    SyncLock ProgressLock
                        progress1 += 1
                    End SyncLock
                    Dim Temp As String = GetFileName(node.Name)
                    If Temp.Length > 0 Then
                        If TestStringsEqual(LCase(Temp), SearchString) = True Then
                            Dim item As New ListViewItem
                            item.Text = node.Text
                            item.Name = node.Name
                            GetIcon(node, item, SearchView.Icons) ' get and assign icon
                            PopulateItemColumns(item, item.Name)
                            SyncLock BufferLock
                                SearchCollectionBuffer.Add(item)
                            End SyncLock
                        End If
                        If node.Nodes.Count > 0 Then ' has child nodes
                            SearchSubstring(SearchString, SearchView, node) ' recursive call
                        End If
                    End If
                Else
                    Exit For ' abort
                End If
            Next

        End If
    End Sub

    ''' <summary>
    ''' Updates search listview with a new batch of items.
    ''' </summary>
    Private Sub UpdateSearchCollection()
        SyncLock BufferLock
            If Not IsNothing(CurrentSearch) Then
                If SearchCollectionBuffer.Count > 0 Then
                    CurrentSearch.ListViewVar.BeginUpdate()

                    For i As Integer = 0 To SearchCollectionBuffer.Count - 1
                        If SearchThreadAbort = False Then
                            CurrentSearch.ListViewVar.Items.Add(SearchCollectionBuffer.Item(i))
                        Else
                            Exit For
                        End If
                    Next
                    SearchCollectionBuffer.Clear()

                    CurrentSearch.ListViewVar.EndUpdate()
                End If
            End If
        End SyncLock
    End Sub

    Private Sub SearchUpdateTimer_Tick(sender As Object, e As EventArgs) Handles SearchUpdateTimer.Tick
        SyncLock BufferLock
            SearchUpdateTimer.Stop()
            If SearchThreadAbort = False Then
                UpdateSearchCollection()
            End If
            SearchUpdateTimer.Start()
        End SyncLock
    End Sub

    ''' <summary>
    ''' Contains data relevant to the search results that a user has initiated, including a seperate ListView to display those results.
    ''' </summary>
    Public Class SearchResult
        Implements IDisposable
        Public WithEvents ListViewVar As New ListView  ' contains the item search results of a search
        Public Icons As New ImageList ' stores icons for the ListView
        Private ContextMenu As ContextMenuStrip = ThisForm.SearchContextMenuStrip
        Public AddressBar As String  ' contains text for the addressbar TextBox to display
        Public SearchBar As String  ' contains text for search TextBox to display
        Public SearchCompleted As Boolean = True  ' true means the search for the string present in SearchBar was not fully completed
        Private SearchColumnSortInfo As ColumnSortInfo ' used as info for ListViewMain

        Public Sub New()
            With ThisForm
                ' adopt same columns, position, and other things
                Me.AddressBar = "Search Results"
                Me.ListViewVar.Name = "SearchView_" & CStr(.SearchCount)
                .SearchCount += 1

                Me.ListViewVar.Location = .ListViewMain.Location
                Me.ListViewVar.Size = .ListViewMain.Size
                Me.ListViewVar.Font = .ListViewMain.Font
                Me.ListViewVar.ContextMenuStrip = ThisForm.SearchContextMenuStrip
                ThisForm.Controls.Add(Me.ListViewVar)
                For i As Integer = 0 To .ListViewMain.Columns.Count - 1
                    Dim col As New ColumnHeader
                    col.Name = .ListViewMain.Columns.Item(i).Name
                    col.Text = .ListViewMain.Columns.Item(i).Text
                    col.Width = .ListViewMain.Columns.Item(i).Width
                    Me.ListViewVar.Columns.Add(col)
                Next
                Me.ListViewVar.View = .ListViewMain.View
                .SplitContainer1.Panel2.Controls.Add(Me.ListViewVar)
                Me.ListViewVar.Dock = DockStyle.Fill
                Me.ListViewVar.SmallImageList = Me.Icons
                'Me.ListViewVar.SmallImageList = .ListViewMainIcons

                SearchColumnSortInfo = New ColumnSortInfo
                If Not IsNothing(.MainColumnSortInfo.CurrentColumn) Then
                    For Each col As ColumnHeader In .ListViewMain.Columns
                        If col.Text = .MainColumnSortInfo.CurrentColumn.Text Then
                            Me.SearchColumnSortInfo.CurrentColumn = Me.ListViewVar.Columns.Item(col.Index) ' set new column
                            Me.SearchColumnSortInfo.CurrentColumn.Text = col.Text
                        End If
                    Next

                    If Not Me.SearchColumnSortInfo.CurrentSortOrder = ThisForm.MainColumnSortInfo.CurrentSortOrder Then
                        Me.SearchColumnSortInfo.CurrentSortOrder = ThisForm.MainColumnSortInfo.CurrentSortOrder
                    End If
                    'Me.ListViewVar.SuspendLayout()
                    ' Create a comparer.
                    Me.ListViewVar.ListViewItemSorter = New ListViewComparer(Me.SearchColumnSortInfo.CurrentColumn.Index, Me.SearchColumnSortInfo.CurrentSortOrder)
                    ' Sort.
                    Me.ListViewVar.Sort()
                    'Me.ListViewVar.ResumeLayout()
                End If

            End With
        End Sub

        Protected disposed As Boolean = False ' Keep track of when the object is disposed. 

        ' This method disposes the base object's resources. 
        Protected Overridable Sub Dispose(ByVal disposing As Boolean)
            If Not Me.disposed Then
                If disposing Then
                    ' Insert code to free managed resources. 
                    ListViewVar.Items.Clear()
                    Me.Icons.Images.Clear()
                    ListViewVar.Visible = False
                    AddressBar = Nothing
                    SearchBar = Nothing
                    SearchColumnSortInfo.CurrentColumn = Nothing
                    SearchColumnSortInfo.CurrentSortOrder = Nothing
                End If
                ' Insert code to free unmanaged resources. 
            End If
            Me.disposed = True
        End Sub

        Public Overridable Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            'GC.SuppressFinalize(Me)
            GC.Collect()
        End Sub

        Public Sub Show(Optional displayOnlyFlag As Boolean = False)
            With ThisForm

                ' abort the thread if active
                .SearchChangeTimer.Stop()
                .SearchThreadAbort = True
                SyncLock .SearchLock
                    .SearchThreadAbort = False
                End SyncLock

                If displayOnlyFlag = False Then
                    .AddressTextBox.Text = Me.AddressBar
                    .SearchTextBox_TextChanged_Lock = True
                    .SearchTextBox.Text = Me.SearchBar
                    .SearchTextBox_TextChanged_Lock = False
                    .SearchTextBox.ForeColor = SystemColors.WindowText

                    For i As Integer = 0 To .ListViewMain.Columns.Count - 1
                        'Me.ListViewVar.Columns.Item(i).Name = .ListViewMain.Columns.Item(i).Name
                        Me.ListViewVar.Columns.Item(i).Text = .ListViewMain.Columns.Item(i).Text
                        Me.ListViewVar.Columns.Item(i).Width = .ListViewMain.Columns.Item(i).Width
                    Next
                    If Not IsNothing(.MainColumnSortInfo.CurrentColumn) AndAlso (IsNothing(Me.SearchColumnSortInfo.CurrentColumn) OrElse _
                        (Not Me.SearchColumnSortInfo.CurrentColumn.Text = ThisForm.MainColumnSortInfo.CurrentColumn.Text)) Then

                        For Each col As ColumnHeader In .ListViewMain.Columns
                            If col.Text = ThisForm.MainColumnSortInfo.CurrentColumn.Text Then
                                Me.SearchColumnSortInfo.CurrentColumn = Me.ListViewVar.Columns.Item(col.Index) ' set new column
                                Me.SearchColumnSortInfo.CurrentColumn.Text = col.Text
                            End If
                        Next

                        If Not Me.SearchColumnSortInfo.CurrentSortOrder = ThisForm.MainColumnSortInfo.CurrentSortOrder Then
                            Me.SearchColumnSortInfo.CurrentSortOrder = ThisForm.MainColumnSortInfo.CurrentSortOrder
                        End If

                        ' Create a comparer.
                        Me.ListViewVar.ListViewItemSorter = New ListViewComparer(Me.SearchColumnSortInfo.CurrentColumn.Index, Me.SearchColumnSortInfo.CurrentSortOrder)
                        ' Sort.
                        Me.ListViewVar.Sort()
                    End If

                    If Me.SearchCompleted = False Then ' continue searching if not completed before
                        .SearchChangeTimer_Tick()
                    End If

                End If

                .CurrentVisibleListView = Me.ListViewVar
                .SearchIsCurrentlyDisplayed = True
                .ListViewMain.Visible = False
                If Not IsNothing(Me.ListViewVar) Then
                    Me.ListViewVar.Visible = True
                End If
            End With
        End Sub

        Public Sub Hide()
            With ThisForm
                ' Stop thread if still active.
                .SearchChangeTimer.Stop()
                If .SearchThreadAlive = True Then
                    Me.SearchCompleted = False
                Else
                    Me.SearchCompleted = True
                End If
                .SearchThreadAbort = True ' abort the thread if active
                SyncLock .SearchLock
                    .SearchThreadAbort = False
                End SyncLock

                ' Save column and sort info to main ListView.
                For i As Integer = 0 To Me.ListViewVar.Columns.Count - 1
                    '.ListViewMain.Columns.Item(i).Name = Me.ListViewVar.Columns.Item(i).Name
                    .ListViewMain.Columns.Item(i).Text = Me.ListViewVar.Columns.Item(i).Text
                    .ListViewMain.Columns.Item(i).Width = Me.ListViewVar.Columns.Item(i).Width
                Next
                If Not IsNothing(Me.SearchColumnSortInfo.CurrentColumn) AndAlso (IsNothing(.MainColumnSortInfo.CurrentColumn) OrElse _
                    (Not Me.SearchColumnSortInfo.CurrentColumn.Text = .MainColumnSortInfo.CurrentColumn.Text)) Then

                    For Each col As ColumnHeader In Me.ListViewVar.Columns
                        If col.Text = Me.SearchColumnSortInfo.CurrentColumn.Text Then
                            .MainColumnSortInfo.CurrentColumn = .ListViewMain.Columns.Item(col.Index) ' set new column
                            .MainColumnSortInfo.CurrentColumn.Text = col.Text
                        End If
                    Next

                    If Not Me.SearchColumnSortInfo.CurrentSortOrder = .MainColumnSortInfo.CurrentSortOrder Then
                        .MainColumnSortInfo.CurrentSortOrder = Me.SearchColumnSortInfo.CurrentSortOrder
                    End If

                    ' Create a comparer.
                    .ListViewMain.ListViewItemSorter = New ListViewComparer(.MainColumnSortInfo.CurrentColumn.Index, .MainColumnSortInfo.CurrentSortOrder)
                    ' Sort.
                    Me.ListViewVar.Sort()
                End If

                'Me.AddressBar = "Search Results"
                Me.SearchBar = .SearchTextBox.Text
                .CurrentVisibleListView = .ListViewMain
                .SearchIsCurrentlyDisplayed = False
                .ListViewMain.Visible = True
                If Not IsNothing(Me.ListViewVar) Then
                    Me.ListViewVar.Visible = False
                End If
            End With
        End Sub

        Private Sub ListView_ItemActivate() Handles ListViewVar.ItemActivate
            If ListViewVar.SelectedItems.Count > 0 Then
                Dim Path As String = ListViewVar.SelectedItems.Item(0).Name ' open first selected item
                If File.Exists(Path) Then
                    RandomFile.OpenFile(Path)
                ElseIf Directory.Exists(Path) Then
                    If Path.Length > 0 AndAlso Not Path(Path.Length - 1) = "\" Then
                        Path &= "\"
                    End If
                    With ThisForm
                        Me.Hide()
                        .NewPreviousDirectory(Me)
                        .ForwardDirectory.Clear()
                        .ForwardButton.Enabled = False
                        .CurrentDirectory = Path
                        .AddressTextBox.Text = Path
                        .TreeViewAll.SelectedNode = .CurrentNodeDisplayed
                        .SearchTextBox_TextChanged_Lock = True
                        .SearchTextBox.Text = ""
                        .SearchTextBox_TextChanged_Lock = False
                        .DisplayDirectory(Path) ' hides this search
                        .ResetSearchText() ' refresh search box text
                    End With
                End If
            End If
        End Sub

        ''' <summary>
        ''' Handles the clicking of column headers.
        ''' </summary>
        Private Sub ListViewVar_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListViewVar.ColumnClick
            ThisForm.ColumnClick(sender, e, Me.ListViewVar, Me.SearchColumnSortInfo)
        End Sub
    End Class

    ' search ListView related
    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        If Not IsNothing(CurrentSearch) AndAlso SearchIsCurrentlyDisplayed = True AndAlso _
            Not IsNothing(CurrentSearch.ListViewVar.SelectedItems) AndAlso CurrentSearch.ListViewVar.SelectedItems.Count > 0 Then
            If Directory.Exists(CurrentSearch.ListViewVar.SelectedItems.Item(0).Name) OrElse File.Exists(CurrentSearch.ListViewVar.SelectedItems.Item(0).Name) Then
                With RandomFile
                    .ContainingButton(CurrentSearch.ListViewVar.SelectedItems.Item(0).Name)
                End With
            End If
        End If
    End Sub

    ''' <summary>
    ''' Test if a given string contains a given substring.
    ''' </summary>
    ''' <param name="StringTest">Searched string.</param>
    ''' <param name="SearchString">substring to be used for search.</param>
    ''' <returns>True if 2nd string parameter is found within the first string parameter.</returns>
    Private Function TestStringsEqual(ByRef StringTest As String, ByRef SearchString As String) As Boolean
        If StringTest.Length >= SearchString.Length Then

            Dim count As Integer
            Dim Maximum As Integer = StringTest.Length - SearchString.Length
            For i As Integer = 0 To Maximum
                If StringTest(i) = SearchString(0) Then ' first chars matched

                    Dim flag As Boolean = True
                    count = 1
                    For k As Integer = i + 1 To (SearchString.Length - 1) + i
                        If Not StringTest(k) = SearchString(count) Then
                            flag = False
                            Exit For
                        End If
                        count += 1
                    Next
                    If flag = True Then
                        Return True
                    End If

                End If
            Next

        End If
        Return False
    End Function

    Private Sub BreakButton_Click(sender As Object, e As EventArgs) Handles BreakButton.Click
        Debugger.Break()
    End Sub

    Dim MainColumnSortInfo As New ColumnSortInfo ' used as info for ListViewMain
    ''' <summary>
    ''' Contains the current way items are sorted and the column being sorted
    ''' </summary>
    Class ColumnSortInfo
        Public CurrentColumn As ColumnHeader ' contains the column
        Public CurrentSortOrder As SortOrder = SortOrder.None ' contains the current sort
    End Class
    ''' <summary>
    ''' Handles the clicking of column headers.
    ''' </summary>
    Private Sub ListViewMain_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ListViewMain.ColumnClick
        ColumnClick(sender, e, ListViewMain, MainColumnSortInfo)
    End Sub

    Private Sub ColumnClick(ByRef sender As Object, ByRef e As ColumnClickEventArgs, ByRef TheListView As ListView, ByRef TheColumnSortInfo As ColumnSortInfo)
        With TheColumnSortInfo
            ' Get the new sorting column.
            Dim NewColumn As ColumnHeader = TheListView.Columns(e.Column)

            ' Figure out the new sorting order.
            If .CurrentColumn Is Nothing Then
                ' New column. Sort ascending.
                .CurrentSortOrder = SortOrder.Ascending
            Else
                ' See if this is the same column.
                If NewColumn.Text = .CurrentColumn.Text AndAlso .CurrentColumn.Text.StartsWith("▲ ") Then
                    .CurrentSortOrder = SortOrder.Descending ' Same column. Switch the sort order.
                Else
                    .CurrentSortOrder = SortOrder.Ascending ' New column. Sort ascending.
                End If

                .CurrentColumn.Text = .CurrentColumn.Text.Substring(2) ' Remove the old sort indicator.
            End If

            ' Display the new sort order.
            .CurrentColumn = NewColumn
            If .CurrentSortOrder = SortOrder.Ascending Then
                .CurrentColumn.Text = "▲ " & .CurrentColumn.Text
            Else
                .CurrentColumn.Text = "▼ " & .CurrentColumn.Text
            End If

            ' Create a comparer.
            TheListView.ListViewItemSorter = New ListViewComparer(e.Column, .CurrentSortOrder)

            ' Sort.
            TheListView.Sort()
        End With
    End Sub

    ''' <summary>
    ''' Implements a comparer for ListView columns.
    ''' </summary>
    Private Class ListViewComparer
        Implements IComparer

        Private columnNumber As Integer
        Private CurrentSortOrder As SortOrder

        Public Sub New(ByVal columnNumberTmp As Integer, ByVal CurrSortOrderTmp As SortOrder)
            columnNumber = columnNumberTmp
            CurrentSortOrder = CurrSortOrderTmp
        End Sub

        ''' <summary>
        ''' Compare the items in the appropriate column.
        ''' </summary>
        Public Function Compare(ByVal obj1 As Object, ByVal obj2 As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim Item1 As ListViewItem = DirectCast(obj1, ListViewItem)
            Dim Item2 As ListViewItem = DirectCast(obj2, ListViewItem)

            ' get the sub-item values
            Dim X As String
            If Item1.SubItems.Count <= columnNumber Then
                X = ""
            Else
                X = Item1.SubItems(columnNumber).Text
            End If

            Dim Y As String
            If Item2.SubItems.Count <= columnNumber Then
                Y = ""
            Else
                Y = Item2.SubItems(columnNumber).Text
            End If

            ' compare them
            If CurrentSortOrder = SortOrder.Ascending Then
                If columnNumber = 3 Then ' file size
                    If X.Length > 3 AndAlso Y.Length > 3 Then
                        X = X.Remove(X.Length - 3, 3) ' remove " KB"
                        Y = Y.Remove(Y.Length - 3, 3) ' remove " KB"
                        Return Val(X).CompareTo(Val(Y))
                    Else
                        Return String.Compare(X, Y)
                    End If
                ElseIf IsNumeric(X) AndAlso IsNumeric(Y) Then
                    Return Val(X).CompareTo(Val(Y))
                ElseIf IsDate(X) AndAlso IsDate(Y) Then
                    Return DateTime.Parse(X).CompareTo(DateTime.Parse(Y))
                Else
                    Return String.Compare(X, Y)
                End If
            Else
                If columnNumber = 3 Then ' file size
                    If X.Length > 3 AndAlso Y.Length > 3 Then
                        X = X.Remove(X.Length - 3, 3) ' remove " KB"
                        Y = Y.Remove(Y.Length - 3, 3) ' remove " KB"
                        Return Val(Y).CompareTo(Val(X))
                    Else
                        Return String.Compare(Y, X)
                    End If
                ElseIf IsNumeric(X) AndAlso IsNumeric(Y) Then
                    Return Val(Y).CompareTo(Val(X))
                ElseIf IsDate(X) AndAlso IsDate(Y) Then
                    Return DateTime.Parse(Y).CompareTo(DateTime.Parse(X))
                Else
                    Return String.Compare(Y, X)
                End If
            End If
        End Function
    End Class
End Class