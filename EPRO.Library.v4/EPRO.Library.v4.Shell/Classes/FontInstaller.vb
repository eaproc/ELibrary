Option Strict Off

Imports System.IO

Imports EPRO.Library.v3._5
Imports EPRO.Library.v4.Shell


Public Class FontInstaller

    Sub New(ByVal FontFilePath As String)

        Try
            Dim dFntFile As String = Environment.GetFolderPath(Environment.SpecialFolder.Fonts) & "\" & Path.GetFileName(FontFilePath)
            If Not File.Exists(dFntFile) Then
                If OperatingSystem.getOSType() = OperatingSystem.MicrosoftOS.WINDOWS_XP Then

                    File.Copy(FontFilePath, dFntFile, True)

                Else

                    Dim objShell As Object = CreateObject("Shell.Application")

                    Dim WShell As Object = CreateObject("WScript.Shell")

                    Dim objFSO As Object = CreateObject("Scripting.Filesystemobject")

                    Dim objNameSpace As Object = objShell.Namespace(FileIO.FileSystem.GetParentPath(FontFilePath))

                    Dim objFont As Object = objNameSpace.ParseName(EIO.getFileName(FontFilePath))

                    objFont.InvokeVerb("Install")

                    objFont = Nothing
                    objNameSpace = Nothing
                    objFSO = Nothing
                    WShell = Nothing
                    objShell = Nothing

                End If

            End If

            Me.___successful = True
        Catch ex As Exception
            Program.Logger.Write(ex)
            Throw New Exception(FontFilePath & " could not be installed", ex)
        End Try

    End Sub




    Private ___successful As Boolean = False
    Public ReadOnly Property IsSuccessful As Boolean
        Get
            Return Me.___successful
        End Get
    End Property


    Public Shared Sub LaunchXPFontActivator()
        Try


            Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Fonts))

        Catch ex As Exception

        End Try
    End Sub

End Class
