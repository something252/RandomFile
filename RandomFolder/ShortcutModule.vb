Imports IWshRuntimeLibrary

Module ShortcutModule

    ''' <summary>
    ''' Create a shortcut in a specified location.
    ''' </summary>
    ''' <param name="shortcutName">The name of the shortcut.</param>
    ''' <param name="creationDir">The folder where the shortcut should be created.</param>
    ''' <param name="targetFullpath">The full path to the target file.</param>
    ''' <param name="workingDir">The folder to use as working directory for the target of the shortcut.</param>
    ''' <param name="iconFile">The file that contains the icon for the shortcut.</param>
    ''' <param name="iconNumber">The position of the icon within the icon file. This is zero if the icon file is an icon itself.</param>
    Public Function CreateShortCut(ByVal shortcutName As String, ByVal creationDir As String, ByVal targetFullpath As String, _
                                    ByVal workingDir As String, ByVal iconFile As String, ByVal iconNumber As Integer) As Boolean
        If RandomFolder.shortcutCreationAllowed = True Then
            Try
                If Not IO.Directory.Exists(creationDir) Then
                    Dim retVal As DialogResult = MsgBox("""" & creationDir & """" & " does not exist. Do you wish to create it?", MsgBoxStyle.YesNo)
                    If retVal = DialogResult.Yes Then
                        Try
                            IO.Directory.CreateDirectory(creationDir)
                        Catch ex As Exception
                            MsgBox("Failed to create directory", MsgBoxStyle.Critical, "Error")
                            Return False
                        End Try
                    Else
                        Return False
                    End If
                End If

                Dim shortCut As IWshRuntimeLibrary.IWshShortcut
                Dim shell As New WshShell
                shortCut = CType(shell.CreateShortcut(creationDir & "\" & shortcutName & ".lnk"), IWshRuntimeLibrary.IWshShortcut)
                shortCut.TargetPath = targetFullpath
                shortCut.WindowStyle = 1
                'shortCut.Description = "Random Folder Shortcut"
                shortCut.WorkingDirectory = workingDir
                'shortCut.IconLocation = iconFile & ", " & iconNumber
                shortCut.Save()
                Return True
            Catch ex As System.Exception
                Return False
            End Try
        End If
        Return False
    End Function

    ''' <summary>
    ''' Determine an available name for the shortcut if one already exists.
    ''' </summary>
    Public Sub DetermineShortcuts(ByVal shortcutName As String, ByVal creationDir As String)
        If RandomFolder.shortcutCreationAllowed = True Then
            shortcutName = shortcutName & " - Shortcut"

            If My.Computer.FileSystem.FileExists(creationDir & "\" & shortcutName & ".lnk") Then ' if shortcut file already exists

                For count As Integer = 2 To 2000 Step 1
                    If My.Computer.FileSystem.FileExists(shortcutName & "\" & " (" & count & ")" & ".lnk") Then ' if shortcut file already exists
                        'MsgBox("file exists." & " (" & count & ")", MsgBoxStyle.Critical)
                    Else ' file doesn't exist
                        shortcutName = shortcutName & " (" & count & ")"
                        Exit For
                    End If
                Next
            End If

            CreateShortCut(shortcutName, creationDir, RandomFolder.ResultBox.Text, RandomFolder.TextBox1.Text, "C:\Windows\System32\imageres.dll", 4)
            If CInt(RandomFolder.TimesBox.Text) >= 2 Then
                CreateShortCut(shortcutName & " - Copy", creationDir, RandomFolder.ResultBox.Text, RandomFolder.TextBox1.Text, "C:\Windows\System32\imageres.dll", 4)
            End If
            For count As Integer = 3 To CInt(RandomFolder.TimesBox.Text) Step 1
                CreateShortCut(shortcutName & " - Copy (" & count & ")", creationDir, RandomFolder.ResultBox.Text, RandomFolder.TextBox1.Text, "C:\Windows\System32\imageres.dll", 4)
            Next
        End If
    End Sub
End Module
