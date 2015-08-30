Module FileTree
    Private SubFiles As New ArrayList
    Private SubFolders As New ArrayList
    Private TmpStr(), FileExtension As String
    ''' <summary>
    ''' Flag for whether or not group value was provided and is being used.
    ''' </summary>
    Private groupFlag As Boolean = False

    ''' <summary>
    ''' Create tree of all sub-files.
    ''' </summary>
    ''' <param name="path">Full directory path</param>
    ''' <param name="group">Group that it is a part of</param>
    ''' <returns>List of all files found.</returns>
    Public Function CreateFileList(ByVal path As String, Optional ByVal group As String = "") As ArrayList
        ' clear files list first
        SubFiles.Clear()
        DoWork(path, group)
        Return SubFiles ' variable count will be 0 if no files found
    End Function

    ''' <summary>
    ''' Receives a folder directory and returns all sub-files within it.
    ''' </summary>
    Private Sub DoWork(ByVal path As String, Optional ByVal group As String = "")
        If group = "" Then
            Try
                Dim list = System.IO.Directory.GetFiles(path)

                GetFiles(list)

                ' perform file adding on all subfolders found as well
                GetFilesFromSubFolders(list, path)

            Catch e As UnauthorizedAccessException ' system files like "system volume information" invoke this
                'MsgBox(e.Message, MsgBoxStyle.Critical, "Critical Error")
            Catch ' other error
            End Try
        Else ' part of a group
            groupFlag = True
            For i As Integer = 0 To RandomFolder.FolderDirectories.Count - 1 ' searches and performs work for all directories in group
                If RandomFolder.FolderGroups(i) = group AndAlso RandomFolder.DisabledLine(i) = False Then
                    Dim pathTemp As String = RandomFolder.FolderDirectories(i)
                    Try
                        Dim list = System.IO.Directory.GetFiles(pathTemp)

                        GetFiles(list)

                        ' perform file adding on all subfolders found as well
                        GetFilesFromSubFolders(list, pathTemp)

                    Catch e As UnauthorizedAccessException ' system files like "system volume information" invoke this
                        'MsgBox(e.Message, MsgBoxStyle.Critical, "Critical Error")
                    Catch ' other error
                    End Try
                End If
            Next
            groupFlag = False
        End If
    End Sub

    Sub GetFiles(ByRef list() As String)
        If (Not My.Settings.WhitelistedFileTypes.Count = 0 AndAlso My.Settings.WhitelistEnabled = True) OrElse _
           (Not My.Settings.BlacklistedFileTypes.Count = 0 AndAlso My.Settings.BlacklistEnabled = True) Then ' whitelist contains extensions (add only those)
            For Each FilePath In list
                'TmpStr = Split(FilePath, ".")
                'FileExtension = TmpStr(TmpStr.Count - 1) ' Put last element into variable (to find extension)
                FileExtension = GetFileType(FilePath)

                Dim flag1 As Boolean = True ' true means a file path is allowed entry into the SubFiles list
                If My.Settings.WhitelistEnabled = True AndAlso Not My.Settings.WhitelistedFileTypes.Count = 0 Then
                    flag1 = False
                    For Each FileType In My.Settings.WhitelistedFileTypes
                        If TestStringEquivalence(FileExtension, FileType) Then
                            flag1 = True ' in whitelist, so add
                            Exit For
                        End If
                    Next
                End If
                If My.Settings.BlacklistEnabled = True Then
                    For Each ext In My.Settings.BlacklistedFileTypes
                        If TestStringEquivalence(FileExtension, ext) Then
                            flag1 = False ' Is in blacklist, so don't add
                        End If
                    Next
                End If

                If flag1 = True Then
                    If groupFlag = True Then
                        If Not SubFiles.Contains(FilePath) Then ' prevent duplicates for groups
                            If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FilePath)) Then
                                SubFiles.Add(FilePath)
                            End If
                        End If
                    Else
                        If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FilePath)) Then
                            SubFiles.Add(FilePath)
                        End If
                    End If
                End If
            Next
        Else ' whitelist/blacklist is empty (add all files found)
            For Each FilePath In list
                If groupFlag = True Then
                    If Not SubFiles.Contains(FilePath) Then ' prevent duplicates for groups
                        If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FilePath)) Then
                            SubFiles.Add(FilePath)
                        End If
                    End If
                Else
                    If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FilePath)) Then
                        SubFiles.Add(FilePath)
                    End If
                End If
            Next
        End If
    End Sub

    ''' <summary>
    ''' Test if two strings are equal. (Ignores case)
    ''' </summary>
    Function TestStringEquivalence(ByRef ext1 As String, ByRef ext2 As String) As Boolean
        If Not IsNothing(ext1) Then
            If ext1.Length = ext2.Length Then
                Dim ext1c = ext1.ToCharArray, ext2c = ext2.ToCharArray
                For i As Integer = 0 To ext1.Length - 1
                    If Not LCase(ext1c(i)) = LCase(ext2c(i)) Then
                        Return False
                    End If
                Next
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Get extension of file.
    ''' </summary>
    Public Function GetFileType(ByRef Str As String) As String
        If Str.Length > 1 Then
            Dim count As Integer = 0
            For i As Integer = Str.Length - 1 To 0 Step -1
                If Str(i) = "."c AndAlso Not (Str.Length - 1) = i Then
                    Return Str.Substring(i + 1, count)
                ElseIf Str(i) = "\"c Then
                    Return "" ' file has no extension
                End If
                count += 1
            Next
        End If
        Return ""
    End Function

    Sub GetFilesFromSubFolders(ByRef list() As String, ByRef path As String)
        list = System.IO.Directory.GetDirectories(path)
        If list.Length > 0 Then ' more folders found (links to folders don't count)
            For Each directory In list
                DoWork(directory) ' call self (recursion)
            Next
        End If
    End Sub

    ''' <summary>
    ''' Create tree of all sub-folders.
    ''' </summary>
    ''' <param name="path">Full directory path</param>
    ''' <param name="group">Group that it is a part of</param>
    ''' <returns>List of all folders found.</returns>
    Public Function CreateFolderList(ByVal path As String, Optional ByVal group As String = "") As ArrayList
        ' clear folders list first
        SubFolders.Clear()
        DoWork2(path, group)
        Return SubFolders ' variable count will be 0 if no files found
    End Function

    ''' <summary>
    ''' Receives a folder directory and returns all sub folders within it
    ''' </summary>
    Private Sub DoWork2(ByVal path As String, Optional ByVal group As String = "")
        If group = "" Then
            Try
                Dim list = System.IO.Directory.GetDirectories(path)

                For Each FolderPath In list
                    If groupFlag = True Then
                        If Not SubFolders.Contains(FolderPath) Then ' prevent duplicates for groups
                            If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FolderPath)) Then
                                SubFolders.Add(FolderPath)
                            End If
                        End If
                    Else
                        If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FolderPath)) Then
                            SubFolders.Add(FolderPath)
                        End If
                    End If
                Next

                If list.Length > 0 Then ' more folders found (links to folders don't count)
                    For Each directory In list
                        DoWork2(directory) ' call self (recursion)
                    Next
                End If

            Catch e As UnauthorizedAccessException ' system files like "system volume information" invoke this
                'MsgBox(e.Message, MsgBoxStyle.Critical, "Critical Error")
            Catch ' other error
            End Try
        Else ' part of a group
            groupFlag = True
            For i As Integer = 0 To RandomFolder.FolderDirectories.Count - 1  ' searches and performs work for all directories in group
                If RandomFolder.FolderGroups(i) = group AndAlso RandomFolder.DisabledLine(i) = False Then
                    Dim pathTemp As String = RandomFolder.FolderDirectories(i)
                    Try
                        Dim list = System.IO.Directory.GetDirectories(pathTemp)

                        For Each FolderPath In list
                            If groupFlag = True Then
                                If Not SubFolders.Contains(FolderPath) Then ' prevent duplicates for groups
                                    If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FolderPath)) Then
                                        SubFolders.Add(FolderPath)
                                    End If
                                End If
                            Else
                                If RandomFolder.RollOnceListFlag = False OrElse (RandomFolder.RollOnceListFlag = True AndAlso Not RandomFolder.RollOnceList.Contains(FolderPath)) Then
                                    SubFolders.Add(FolderPath)
                                End If
                            End If
                        Next

                        If list.Length > 0 Then ' more folders found (links to folders don't count)
                            For Each directory In list
                                DoWork2(directory) ' call self (recursion)
                            Next
                        End If

                    Catch e As UnauthorizedAccessException ' system files like "system volume information" invoke this
                        'MsgBox(e.Message, MsgBoxStyle.Critical, "Critical Error")
                    Catch ' other error
                    End Try
                End If
            Next
            groupFlag = False
        End If
    End Sub

End Module
