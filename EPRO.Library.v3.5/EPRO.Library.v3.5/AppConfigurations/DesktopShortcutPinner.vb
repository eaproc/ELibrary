Option Explicit On
Option Strict Off
Imports EPRO.Library.v3._5.Objects
Imports EPRO.Library.v3._5.Modules

Namespace AppConfigurations


    ''' <summary>
    ''' Currently supports win7 and windows 8.1
    ''' </summary>
    ''' <remarks></remarks>
    Public NotInheritable Class DesktopShortcutPinner

        ''' <summary>
        ''' Not case sensitive
        ''' </summary>
        ''' <param name="pShortcutNameWithoutExtension"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal pShortcutNameWithoutExtension As String,
                       Optional ByVal PinIt As Boolean = True)


            REM Worked perfectly with windows 7, windows 8.1

            'Const CSIDL_COMMON_PROGRAMS = &H17
            Dim ShellApp As Object, FSO As Object, Desktop As Object
            ShellApp = CreateObject("Shell.Application")
            FSO = CreateObject("Scripting.FileSystemObject")

            'Set StartMenuFolder = ShellApp.NameSpace(CSIDL_COMMON_PROGRAMS)
            Desktop = ShellApp.NameSpace(Environment.GetFolderPath(Environment.SpecialFolder.Desktop))

            Dim LnkFile As Object
            LnkFile = Desktop.Self.Path & String.Format("\{0}.lnk", pShortcutNameWithoutExtension)

            If (FSO.FileExists(LnkFile)) Then
                Dim verb As Object

                Dim desktopImtes As Object, item As Object
                desktopImtes = Desktop.Items()

                For Each item In desktopImtes
                    '-- - Debug.Print("item.Name: " & item.Name)
                    If (item.Name.ToString().ToLower() = pShortcutNameWithoutExtension.ToLower()) Then
                        For Each verb In item.Verbs


                            'Debug.Print("verb.Name: " & verb.Name)
                            If (
                                    EStrings.valueOf(verb.Name).equalsIgnoreCase("Pin to Tas&kbar") AndAlso PinIt
                                ) OrElse
                                (
                                    EStrings.valueOf(verb.Name).equalsIgnoreCase("Unpin from Tas&kbar") AndAlso Not PinIt
                                ) Then 'If (verb.Name = "锁定到任务栏(&K)")
                                verb.DoIt()
                                GoTo END__CREATING
                            End If

                            ''If IsWin10 Then



                            ''Else

                            ''    'Debug.Print("verb.Name: " & verb.Name)
                            ''    If (
                            ''        (verb.Name = "Pin to Tas&kbar") AndAlso PinIt
                            ''        ) OrElse
                            ''    (
                            ''        (verb.Name = "Unpin from Tas&kbar") AndAlso Not PinIt
                            ''        ) Then 'If (verb.Name = "锁定到任务栏(&K)")
                            ''        verb.DoIt()
                            ''        GoTo END__CREATING
                            ''    End If


                            ''End If


                        Next
                    End If
                Next

            Else
                Modules.MyLogFile.Log("Link does not exist: " & LnkFile)
            End If

END__CREATING:
            FSO = Nothing
            ShellApp = Nothing



        End Sub





    End Class

End Namespace
