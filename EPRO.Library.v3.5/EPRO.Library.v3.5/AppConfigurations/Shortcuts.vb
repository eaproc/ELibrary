Imports System.IO
Imports EPRO.Library.v3._5
Imports Shell32

Namespace AppConfigurations


    Public NotInheritable Class Shortcuts

        Public Structure ShortcutLink
            Sub New(ByVal FullName As String,
                    ByVal Arguments As String,
                    ByVal Target As String)

                Me.__FullName = FullName
                Me.__Arguments = Arguments
                Me.__Target = Target

            End Sub

            Private __FullName As String
            Private __Arguments As String
            Private __Target As String

            Public ReadOnly Property Name As String
                Get
                    Return EIO.getFileName(Me.FullName)
                End Get
            End Property
            Public ReadOnly Property FullName As String
                Get
                    Return Me.__FullName
                End Get
            End Property

            Public ReadOnly Property Target As String
                Get

                    Return Me.__Target
                End Get
            End Property


            Public ReadOnly Property Arguments As String
                Get
                    Return Me.__Arguments
                End Get
            End Property



        End Structure

        Public Shared Function getAllShortcutsOn(ByVal FolderPath As String) As IEnumerable(Of FileInfo)
            Dim rst As New List(Of FileInfo)

            Try


                Dim di As DirectoryInfo = New DirectoryInfo(FolderPath)

                Return di.GetFiles("*lnk")




            Catch ex As Exception




            End Try


            Return rst

        End Function


        Public Shared Function getShortcut(ByVal linkFile As FileInfo) As ShortcutLink
            Try

                Dim shell As Shell32.Shell = New Shell32.ShellClass()
                Dim folder As Shell32.Folder = shell.NameSpace(linkFile.DirectoryName)
                Dim folderItem As Shell32.FolderItem = folder.ParseName(linkFile.Name)
                Dim link As Shell32.ShellLinkObject = CType(folderItem.GetLink, Shell32.ShellLinkObject)

                Return New ShortcutLink(linkFile.FullName, link.Arguments, link.Path)

            Catch ex As Exception
                Return Nothing
            End Try
        End Function


        Public Shared Function createShortcut(ByVal PathLinkWithoutDot As String,
                                              ByVal TargetFile As String,
                                              Optional ByVal Arguments As String = "",
                                              Optional ByVal IconPath As String = "",
                                              Optional ByVal comments As String = "Launch") As Boolean

            Try


                Dim wshShell As IWshRuntimeLibrary.WshShellClass = New IWshRuntimeLibrary.WshShellClass()
                Dim shortcut As IWshRuntimeLibrary.IWshShortcut = CType(wshShell.CreateShortcut(PathLinkWithoutDot & ".lnk"), 
                    IWshRuntimeLibrary.IWshShortcut)

                shortcut.Arguments = Arguments
                shortcut.TargetPath = TargetFile
                shortcut.WorkingDirectory = FileIO.FileSystem.GetParentPath(TargetFile)
                If IconPath IsNot Nothing AndAlso IconPath <> "" Then
                    shortcut.IconLocation = IconPath
                End If

                shortcut.Description = comments
                shortcut.Save()


                Return RefreshDesktopShortcuts()

            Catch ex As Exception
                Modules.MyLogFile.Print("createShortcut", ex)
                Return False
            End Try
        End Function




        Public Shared Function RefreshDesktopShortcuts() As Boolean

            Try

                REM Clear Icon Cache for windows 7 and 8 to reflect right icon
                Dim cmd As New MicrosoftOS.CommandPrompt(False)

                If Not IO.Directory.Exists("C:\Windows\SysWOW64") Then
                    REM 32bit
                    cmd.Execute("ie4uinit.exe -ClearIconCache")

                Else
                    REM 64bit
                    REM cmd.Execute("ie4uinit.exe -ClearIconCache", "%windir%\sysnative")
                    cmd.Execute("cd %windir%\sysnative & ie4uinit.exe -ClearIconCache")
                    REM Interaction.Shell("cmd.exe /K ""cd %windir%\sysnative & ie4uinit.exe -ClearIconCache""", AppWinStyle.NormalFocus, True)

                End If

                cmd = Nothing

                '   OR
                ''                CD /d %userprofile%\AppData\Local
                ''                DEL IconCache.db / a
                ''EXIT


                ''                restart Process
                ''                explorer()


                Return True

            Catch ex As Exception
                Modules.MyLogFile.Print("RefreshDesktopShortcuts", ex)
                Return False
            End Try


        End Function






    End Class

End Namespace