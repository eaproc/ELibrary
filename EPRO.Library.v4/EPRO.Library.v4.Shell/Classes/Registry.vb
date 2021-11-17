Imports EPRO.Library.v3._5.Objects.EStrings
Imports EPRO.Library.v4.Shell.OperatingSystem
Imports EPRO.Library.v3._5.Objects
Imports EPRO.Library.v4.Shell.Modules.basRegistry

''' <summary>
''' Contains Functions that acts on Registry [Invoke as Admin]
''' </summary>
''' <remarks>Always invoke as administrator</remarks>
Public Class Registry


    REM Microsoft.win32.registrykey has limitation to which OS bit you are using
    REM 64 or 32

    REM Testing REgistry

    ''If CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                 "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp",
    ''                                                 "Display1"
    ''                                                  ) IsNot Nothing Then

    ''    CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                     "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\Display1",
    ''                                                     "Display2"
    ''                                                     )

    ''    CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                     "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\Display1",
    ''                                                     "Display21", "Display32"
    ''                                                     )

    ''    CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                    "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\Display1",
    ''                                                    "Display23", "16", Microsoft.Win32.RegistryValueKind.DWord
    ''                                                    )


    ''    DeleteAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                    "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp\Display1",
    ''                                                    "Display23"
    ''                                                    )


    ''    DeleteASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
    ''                                                     "SOFTWARE\Microsoft\Microsoft SQL Server\MSSQL10.SQLEXPRESS\MSSQLServer\SuperSocketNetLib\Tcp",
    ''                                                     "Display1"
    ''                                                     )

    ''End If




#Region "Creating App Extension"

    ''' <summary>
    ''' Create a customised extension that opens with an executor
    ''' </summary>
    ''' <param name="ExtensionWithoutDot">Just the name of the extension without the preceeding dot (.) like txt</param>
    ''' <param name="OpenWithExecPath">Executor full path</param>
    ''' <param name="IconPath">Icon Full Path</param>
    ''' <param name="extDescription">The description of the extension</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateApplicationExtension(ByVal ExtensionWithoutDot As String,
                                                      ByVal OpenWithExecPath As String,
                                                      ByVal IconPath As String,
                                                      Optional ByVal extDescription As String = Nothing) As Boolean

        ''Sample ==>
        '' dllFunctionsNet.CreateApplicationExtension(
        ''                "cfplan",
        ''                Application.ExecutablePath,
        ''                Application.StartupPath & "\Icons\FloorPlanFiles.ico"
        '')
        '' 



        Try


            My.Computer.Registry.ClassesRoot.CreateSubKey("." & ExtensionWithoutDot) _
    .SetValue("", ExtensionWithoutDot & "file", Microsoft.Win32.RegistryValueKind.String)


            If extDescription IsNot Nothing Then
                REM Set Extension Description
                My.Computer.Registry.ClassesRoot.CreateSubKey(ExtensionWithoutDot & "file") _
            .SetValue("", extDescription, Microsoft.Win32.RegistryValueKind.String)


            End If

            My.Computer.Registry.ClassesRoot.CreateSubKey(ExtensionWithoutDot & "file\shell\open\command") _
             .SetValue("", OpenWithExecPath & " %1", Microsoft.Win32.RegistryValueKind.String)

            My.Computer.Registry.ClassesRoot.CreateSubKey(ExtensionWithoutDot & "file\DefaultIcon") _
             .SetValue("", IconPath)


            EPRO.Library.v3._5.AppConfigurations.Shortcuts.RefreshDesktopShortcuts()

        Catch ex As Exception
            Return False
        End Try


        Return True


    End Function

#End Region



#Region "Using RegEdit"

    ''' <summary>
    ''' Main Levels in RegEdit
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum RegEditKeys
        HKEY_CLASSES_ROOT
        HKEY_CURRENT_USER
        HKEY_LOCAL_MACHINE
        HKEY_USERS
        HKEY_CURRENT_CONFIG
    End Enum



    ''' <summary>
    ''' Fetch Object Value from Registry if the path exist
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sKey">Key Name to Read in the path folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function readRegistryObjectValue(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKey As String) As Object



        REM Heading
        REM Path
        REM Key
        REM Value

        REM IF a path does not exist it returns nothing
        REM OpenSubKey opens the direct descendant folder only

        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        REM Reading and Writing Conflicts each other so try close immediately you finish using
        If Not IsNothing(rst) Then
            Dim rstStr As Object = rst.GetValue(sKey)

            rst.Close()
            rst = Nothing

            Return rstStr
        End If





        Return Nothing

    End Function


    ''' <summary>
    ''' Fetch String Value from Registry if the path exist
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sKey">Key Name to Read in the path folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function readRegistryStringValue(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKey As String) As String



        REM Heading
        REM Path
        REM Key
        REM Value

        REM IF a path does not exist it returns nothing
        REM OpenSubKey opens the direct descendant folder only

        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        REM Reading and Writing Conflicts each other so try close immediately you finish using
        If Not IsNothing(rst) Then
            Dim rstStr As String = EStrings.valueOf(rst.GetValue(sKey))

            rst.Close()
            rst = Nothing

            Return rstStr
        End If



        ''Try
        ''    Debug.Print(
        ''        rst.GetValue(sKey)
        ''        )

        ''    rst.CreateSubKey("Testing")
        ''Catch ex As Exception

        ''End Try





        Return vbNullString
    End Function



    ''' <summary>
    ''' Fetch Double Value from Registry if the path exist
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sKey">Key Name to Read in the path folder</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function readRegistryDoubleValue(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKey As String) As Double



        REM Heading
        REM Path
        REM Key
        REM Value

        REM IF a path does not exist it returns nothing
        REM OpenSubKey opens the direct descendant folder only

        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath)
        ' Dim realPath As String = ParseRegistryPath(sPath)





        If Not IsNothing(rst) Then Return EDouble.valueOf(rst.GetValue(sKey))


        Return 0

    End Function




    ''' <summary>
    ''' Create a Subkey [Folder] in the path specified or open it if it already exists
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sSubKeyName">Sub Key Folder to Create</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateASubKey(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sSubKeyName As String) As Microsoft.Win32.RegistryKey




        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath, True)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        Try

            If Not IsNothing(rst) Then Return rst.CreateSubKey(sSubKeyName)

        Catch ex As Exception

        End Try

        Return Nothing

    End Function



    ''' <summary>
    ''' Delete a Subkey [Folder] in the path specified
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sSubKeyName">Sub Key Folder to Deletes</param>
    ''' <returns></returns>
    ''' <remarks>Supports Recursive</remarks>
    Public Shared Function DeleteASubKey(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sSubKeyName As String) As Boolean




        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath, True)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        Try

            If Not IsNothing(rst) Then rst.DeleteSubKeyTree(sSubKeyName)
            Return True

        Catch ex As Exception

        End Try

        Return False

    End Function


    ''' <summary>
    ''' Create a Key in the path specified
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sKeyName">Sub Key Name to Create</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CreateAKey(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKeyName As String,
                                           ByVal sKeyValue As String,
                                           Optional ByVal valueKind As Microsoft.Win32.RegistryValueKind = Microsoft.Win32.RegistryValueKind.String) As Boolean



        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath, True)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        Try

            If Not IsNothing(rst) Then rst.SetValue(sKeyName, sKeyValue, valueKind)
            Return True

        Catch ex As Exception

        End Try

        Return False

    End Function



    ''' <summary>
    ''' Delete a Key in the path specified
    ''' </summary>
    ''' <param name="sLevel">Registry Main Levels [Local_Machine ...]</param>
    ''' <param name="sPath">Path to read not including the Main Level Folder Name</param>
    ''' <param name="sKeyName">Sub Key Name to Delete</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function DeleteAKey(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKeyName As String) As Boolean



        Dim rst As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(sLevel, sPath, True)
        ' Dim realPath As String = ParseRegistryPath(sPath)




        Try

            If Not IsNothing(rst) Then rst.DeleteValue(sKeyName)
            Return True

        Catch ex As Exception

        End Try

        Return False

    End Function


    ''' <summary>
    ''' Get a SubKey Folder if it exist
    ''' </summary>
    ''' <param name="sPath"></param>
    ''' <param name="OpenInWriteMode">If it is in write mode  ... anything you read will return nothing. 
    ''' Otherwise, will pop exception if writing in read mode</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getRegistrySubKeyFolder(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           Optional ByVal OpenInWriteMode As Boolean = False) As Microsoft.Win32.RegistryKey

        Select Case sLevel

            Case Is = RegEditKeys.HKEY_LOCAL_MACHINE
                Try
                    If isThereWow64Folder() Then
                        Return OpenSubKey(Microsoft.Win32.RegistryHive.LocalMachine, sPath, OpenInWriteMode)
                    Else
                        Return Microsoft.Win32.Registry.LocalMachine.OpenSubKey(sPath, OpenInWriteMode)
                    End If

                Catch ex As Exception
                    Return Nothing
                End Try


            Case Is = RegEditKeys.HKEY_CURRENT_USER
                Try
                    If isThereWow64Folder() Then
                        Return OpenSubKey(Microsoft.Win32.RegistryHive.CurrentUser, sPath, OpenInWriteMode)
                    Else
                        Return Microsoft.Win32.Registry.CurrentUser.OpenSubKey(sPath, OpenInWriteMode)
                    End If

                Catch ex As Exception
                    Return Nothing
                End Try


        End Select

        Return Nothing
    End Function


    ''' <summary>
    ''' Check through the sub folder if the key exists
    ''' </summary>
    ''' <param name="sLevel"></param>
    ''' <param name="sPath"></param>
    ''' <param name="sKeySearching"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function searchKeyLikeInSubFolder(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal sKeySearching As String) As Boolean

        Try
            Dim subFolder As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(
                                                                              sLevel,
                                                               sPath
                                                           )

            If subFolder Is Nothing Then Return False

            Dim sKeys() As String = subFolder.GetValueNames()
            If sKeys Is Nothing Then Return False

            For Each sKey As String In sKeys

                If sKey.ToLower = sKeySearching.ToLower Then Return True

            Next

        Catch ex As Exception

        End Try


        Return False

    End Function

    ''' <summary>
    ''' Check through the sub folder if it exists if there is any key with a data [Like] Not Match All
    ''' </summary>
    ''' <param name="sLevel"></param>
    ''' <param name="sPath"></param>
    ''' <param name="data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function searchDataLikeInSubFolder(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal data As String) As Boolean

        Try
            Dim subFolder As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(
                                                                              sLevel,
                                                               sPath
                                                           )

            If subFolder Is Nothing Then Return False

            Dim sKeys() As String = subFolder.GetValueNames()
            If sKeys Is Nothing Then Return False

            For Each sKey As String In sKeys
                Dim sData As String = readRegistryStringValue(sLevel, sPath, sKey)
                If sData <> vbNullString Then
                    REM If Found
                    If sData.IndexOf(data) >= 0 Then Return True
                End If

            Next

        Catch ex As Exception

        End Try


        Return False

    End Function


    ''' <summary>
    ''' Check through the sub folder if it exists if there is any key with a data [Like] Not Match All.
    ''' NB: But must match both data1 and data2
    ''' </summary>
    ''' <param name="sLevel"></param>
    ''' <param name="sPath"></param>
    ''' <param name="data"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function searchDataLikeInSubFolder(ByVal sLevel As RegEditKeys,
                                           ByVal sPath As String,
                                           ByVal data As String, ByVal data2 As String) As Boolean

        Dim subFolder As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(
                                                                          sLevel,
                                                           sPath
                                                       )

        If subFolder Is Nothing Then Return False

        Dim sKeys() As String = subFolder.GetValueNames()
        If sKeys Is Nothing Then Return False

        For Each sKey As String In sKeys
            Dim sData As String = readRegistryStringValue(sLevel, sPath, sKey)
            If sData <> vbNullString Then
                REM If Found
                If sData.IndexOf(data) >= 0 And sData.IndexOf(data2) >= 0 Then Return True
            End If

        Next

        Return False

    End Function



    '' ''' <summary>
    '' ''' Make the path given in a regedit readable format
    '' ''' </summary>
    '' ''' <param name="sPath"></param>
    '' ''' <returns></returns>
    '' ''' <remarks></remarks>
    ''Private Shared Function ParseRegistryPath(ByVal sPath As String) As String
    ''    REM Rule IS it is either parsed full or not parsed at all
    ''    If sPath.IndexOf("\\") >= 0 Then Return sPath

    ''    Return Replace(sPath, "\", "\\")

    ''End Function

    ''' <summary>
    ''' Delete all subfolders in a path except sum listed ones
    ''' </summary>
    ''' <param name="sLevel"></param>
    ''' <param name="sPath"></param>
    ''' <param name="ExceptSubKeys"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function deleteAllSubKeysInThisPathExcept(ByVal sLevel As RegEditKeys,
                                         ByVal sPath As String,
                                         ByVal ExceptSubKeys() As String) As Boolean

        Dim subFolder As Microsoft.Win32.RegistryKey = getRegistrySubKeyFolder(
                                                                          sLevel,
                                                           sPath
                                                       )

        If subFolder Is Nothing Or ExceptSubKeys Is Nothing Then Return False

        Dim subKeys() As String = subFolder.GetSubKeyNames()
        If subKeys Is Nothing Then Return True

        For Each sSubKey As String In subKeys
            If Not ExceptSubKeys.Contains(sSubKey) Then

                DeleteASubKey(sLevel, sPath, sSubKey)

            End If


        Next

        Return True

    End Function


#End Region



End Class
