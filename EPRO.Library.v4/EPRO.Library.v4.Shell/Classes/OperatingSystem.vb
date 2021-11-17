Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Security.Cryptography.X509Certificates


Imports EPRO.Library.v3._5.EIO
Imports EPRO.Library.v3._5.Objects
Imports EPRO.Library.v3._5.Modules

Imports EPRO.Library.v4.Shell.Registry
Imports EPRO.Library.v4.Shell.Modules

Imports CODERiT.Logger.v._3._5.Exceptions

''' <summary>
''' Perform Static Operating System Functions
''' </summary>
''' <remarks></remarks>
Public Class OperatingSystem

#Region "Properties and Consts"


    ''' <summary>
    ''' Microsoft Operating System Types.
    ''' Currently Supporting a few
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum MicrosoftOS
        WINDOWS_XP
        WINDOWS_VISTA
        WINDOWS_7
        WINDOWS_8
        WINDOWS_8_1
        UNKNOWN
        WINDOWS_10
    End Enum

    Public Const TASK_MANAGER_PROCESS As String = "taskmgr.exe"
    Public Const COMMAND_PROMPT_PROCESS As String = "cmd.exe"
    Public Const REGISTRY_TOOLS_PROCESS As String = "regedit.exe"
    Public Const SYSTEM_RESTORE_PROCESS As String = "rstrui.exe"

    ''' <summary>
    ''' for both 32bit and 64bit
    ''' </summary>
    ''' <remarks></remarks>
    Public Const GENERAL_PROGRAM_FILES_FOLDER As String = "C:\Program Files"
    Public Const PROGRAM_FILES_FOLDER_32BIT_IN_64BIT_OS As String = "C:\Program Files (x86)"

    Public Const WOW_64_FOLDER As String = "C:\Windows\SysWOW64"

#End Region

#Region "Methods"

#Region "Private"

#Region "Finding Windows"

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, _
    ByRef lpdwProcessId As Integer) As Integer
    End Function


    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function FindWindow( _
     ByVal lpClassName As String, _
     ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function FindWindowByClass( _
     ByVal lpClassName As String, _
     ByVal zero As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function FindWindowByCaption( _
     ByVal zero As IntPtr, _
     ByVal lpWindowName As String) As IntPtr
    End Function

    'Imports System.Text
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
    Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function

#End Region



    ' ''' <summary>
    ' ''' Get Service Status. 
    ' ''' I made sure the service exist before calling this process else it will return stopped
    ' ''' </summary>
    ' ''' <param name="serviceName"></param>
    ' ''' <returns></returns>
    ' ''' <remarks>NB: I created this because a local creation of service controller is static</remarks>
    'Private Shared Function getServiceStatus(ByVal serviceName As String) As ServiceProcess.ServiceControllerStatus

    '    Dim aService As New ServiceProcess.ServiceController(serviceName)

    '    Try
    '        Return aService.Status
    '    Catch ex As Exception
    '        Return ServiceProcess.ServiceControllerStatus.Stopped
    '    End Try

    'End Function



#End Region

#Region "Public"


#Region "ReOrder Windows"

    <DllImport("user32.dll", SetLastError:=True)> _
    Private Shared Function SetWindowPos(ByVal hWnd As IntPtr, ByVal hWndInsertAfter As IntPtr, ByVal X As Integer, ByVal Y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal uFlags As Integer) As Boolean
    End Function

    Private Const SWP_NOSIZE As Integer = &H1
    Private Const SWP_NOMOVE As Integer = &H2

    Private Shared ReadOnly HWND_TOPMOST As New IntPtr(-1)
    Private Shared ReadOnly HWND_NOTOPMOST As New IntPtr(-2)

    Public Shared Function SetWindow___TopMostZOrder(pHandle As IntPtr) As Boolean
        Return SetWindowPos(pHandle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End Function

    Public Shared Function SetWindow___NormalZOrder(pHandle As IntPtr) As Boolean
        Return SetWindowPos(pHandle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End Function

#End Region


#Region "Fetching Windows"

    ''' <summary>
    ''' Check if there is any window on with this exact caption. NB: NOT Case - Sensitive
    ''' </summary>
    ''' <param name="WindowsCaption"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isWindowOn(ByVal WindowsCaption As String) As Boolean

        Return CBool(
            Val(FindWindowByCaption(Nothing, WindowsCaption).ToInt32)
            )


    End Function

    ''' <summary>
    ''' Returns a Windows Caption Using Handle
    ''' </summary>
    ''' <param name="hWnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetWindowCaptionText(ByVal hWnd As IntPtr) As String

        REM Allocate correct string length first
        Dim length As Integer = GetWindowTextLength(hWnd)
        Dim sb As StringBuilder = New StringBuilder(length + 1)
        GetWindowText(hWnd, sb, sb.Capacity)
        Return sb.ToString()

    End Function


    ''' <summary>
    ''' Get's Windows Handle By Using it's caption
    ''' </summary>
    ''' <param name="Caption"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetWindowHandle__UsingCaption(ByVal Caption As String) As IntPtr

        Return FindWindowByCaption(Nothing, Caption)

    End Function


    ''' <summary>
    ''' Gets a ProcessID from Pointer
    ''' </summary>
    ''' <param name="Hwnd"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetProcessID_From_Handle(ByVal Hwnd As IntPtr) As Integer

        Dim ProcID As Integer = 0

        REM test returned here is not the Handle.. The handle is passed by reference
        ''Dim test As Integer = GetWindowThreadProcessId(Hwnd, ProcID)

        GetWindowThreadProcessId(Hwnd, ProcID)

        Return ProcID

    End Function


#End Region


    ''Operating system	Version number	dwMajorVersion	dwMinorVersion	Other
    ''Windows 8.1	6.3*	6	3	OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    ''Windows Server 2012 R2	6.3*	6	3	OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    ''Windows 8	6.2	6	2	OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    ''Windows Server 2012	6.2	6	2	OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    ''Windows 7	6.1	6	1	OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    ''Windows Server 2008 R2	6.1	6	1	OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    ''Windows Server 2008	6.0	6	0	OSVERSIONINFOEX.wProductType != VER_NT_WORKSTATION
    ''Windows Vista	6.0	6	0	OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION
    ''Windows Server 2003 R2	5.2	5	2	GetSystemMetrics(SM_SERVERR2) != 0
    ''Windows Home Server	5.2	5	2	OSVERSIONINFOEX.wSuiteMask & VER_SUITE_WH_SERVER
    ''Windows Server 2003	5.2	5	2	GetSystemMetrics(SM_SERVERR2) == 0
    ''Windows XP Professional x64 Edition	5.2	5	2	(OSVERSIONINFOEX.wProductType == VER_NT_WORKSTATION) && (SYSTEM_INFO.wProcessorArchitecture==PROCESSOR_ARCHITECTURE_AMD64)
    ''Windows XP	5.1	5	1	Not applicable
    ''Windows 2000	5.0	5	0	Not applicable




    ''Windows NT 4.0 - 4.0.1381
    ''Windows 98 - 4.10.1998 - 4.10.67766222
    ''Windows ME - 4.90.3000 - 4.90.3002
    ''Windows 2000 - 5.0.2195
    ''Windows XP RC1 - 5.1.2505
    ''Windows XP RC2 - 5.1.2526
    ''Windows XP - 5.1.2600
    ''Windows Server 2003 - 5.2.3790
    ''Windows XP x64 - 5.2.3790
    ''Windows Server 2003 - 5.2.3790
    ''Windows Media Center 2002 - 5.1.2600.1106 - 5.1.2600.1142
    ''Windows Media Center 2004 - 5.1.2600.1217 - 5.1.2600.2180
    ''Windows Media Center 2005 - 5.1.2700.2180 - 5.1.2715.2812
    ''Windows Vista - 6.0.6000
    ''Windows Vista SP1 - 6.0.6001
    ''Windows Vista SP2 - 6.0.6002
    ''Windows Server 2008 - 6.0.6001
    ''Windows Server 2008 SP2 - 6.0.6002
    ''Windows 7 - 6.1.7600
    ''Windows 7 SP1 - 6.1.7601
    ''Windows Server 2008 R2 - 6.1.7600
    ''Windows Server 2008 R2 SP1 - 6.0.7601




    ''' <summary>
    ''' Get the major versions of Microsoft Operating Systems
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getOSType() As MicrosoftOS
        '
        ' Fix Windows 10 bug
        '
        Dim pProductName = Registry.readRegistryStringValue(RegEditKeys.HKEY_LOCAL_MACHINE,
                                                            "SOFTWARE\Microsoft\Windows NT\CurrentVersion",
                                                             "ProductName")
        If pProductName IsNot Nothing AndAlso pProductName.IndexOf("Windows 10", StringComparison.CurrentCultureIgnoreCase) >= 0 Then Return MicrosoftOS.WINDOWS_10


        Dim OSver As Version = Environment.OSVersion.Version
        If OSver.Major = 5 Then
            Return MicrosoftOS.WINDOWS_XP

        ElseIf OSver.Major = 6 And OSver.Minor = 0 Then
            Return MicrosoftOS.WINDOWS_VISTA


        ElseIf OSver.Major = 6 And OSver.Minor = 1 Then
            Return MicrosoftOS.WINDOWS_7


        ElseIf OSver.Major = 6 And OSver.Minor = 2 Then
            Return MicrosoftOS.WINDOWS_8


        ElseIf OSver.Major = 6 And OSver.Minor = 3 Then
            Return MicrosoftOS.WINDOWS_8_1

        ElseIf OSver.Major = 10 Then
            Return MicrosoftOS.WINDOWS_10
        Else
            Return MicrosoftOS.UNKNOWN

        End If

    End Function

    ''' <summary>
    ''' Get true if this is 64bit OS
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isThis64BitOS() As Boolean

        Return isThereWow64Folder()

    End Function


    ''' <summary>
    ''' check if this folder exist [C:\Windows\SysWOW64] Which means is it 64bit
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isThereWow64Folder() As Boolean
        REM The only problem with this method for now is 
        REM If user installed the OS in another drive not C:
        REM

        REM Get OS Installed Path

        Return FileIO.FileSystem.DirectoryExists(
                                       WOW_64_FOLDER
                                     )

    End Function


    ''' <summary>
    ''' Determines Which Program File Folder is In - Use
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function get32BitProgramFileFolder() As String


        If FileIO.FileSystem.DirectoryExists(PROGRAM_FILES_FOLDER_32BIT_IN_64BIT_OS) Then
            Return PROGRAM_FILES_FOLDER_32BIT_IN_64BIT_OS
        Else
            Return GENERAL_PROGRAM_FILES_FOLDER
        End If


    End Function



#Region "InstalledApplicationManagement"
    ''' <summary>
    ''' Check if component is installed by checking how the displayName is represented in registry
    ''' </summary>
    ''' <param name="pPartOfAppDisplayName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function IsApplicationInstalled(ByVal pPartOfAppDisplayName As String) As Boolean
        Dim InstalledAppsDisplayName As List(Of String) = New List(Of String)()

        ''Me.TextBox1.Text &= "CHECKING ------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall" & Environment.NewLine
        ''Me.TextBox1.Text &= "---------------------------------------------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= Environment.NewLine


        Dim uninstall As Microsoft.Win32.RegistryKey =
            Registry.getRegistrySubKeyFolder(Registry.RegEditKeys.HKEY_LOCAL_MACHINE, "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")

        If uninstall IsNot Nothing Then
            For Each sApp As String In uninstall.GetSubKeyNames()
                For Each pValueNames As String In uninstall.OpenSubKey(sApp).GetValueNames()
                    If pValueNames.equalsIgnoreCase("DisplayName") Then
                        InstalledAppsDisplayName.Add(uninstall.OpenSubKey(sApp).GetValue(pValueNames).ToString())
                    End If
                Next
            Next

        End If




        ''Me.TextBox1.Text &= Environment.NewLine
        ''Me.TextBox1.Text &= "CHECKING ------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Uninstall" & Environment.NewLine
        ''Me.TextBox1.Text &= "---------------------------------------------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= Environment.NewLine

        uninstall =
            Registry.getRegistrySubKeyFolder(Registry.RegEditKeys.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Uninstall")

        If uninstall IsNot Nothing Then
            For Each sApp As String In uninstall.GetSubKeyNames()
                For Each pValueNames As String In uninstall.OpenSubKey(sApp).GetValueNames()
                    If pValueNames.equalsIgnoreCase("DisplayName") Then
                        InstalledAppsDisplayName.Add(uninstall.OpenSubKey(sApp).GetValue(pValueNames).ToString())
                    End If
                Next
            Next
        End If




        ''Me.TextBox1.Text &= Environment.NewLine
        ''Me.TextBox1.Text &= "CHECKING ------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" & Environment.NewLine
        ''Me.TextBox1.Text &= "---------------------------------------------------------------------" & Environment.NewLine
        ''Me.TextBox1.Text &= Environment.NewLine

        uninstall =
            Registry.getRegistrySubKeyFolder(Registry.RegEditKeys.HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Uninstall")

        If uninstall IsNot Nothing Then
            For Each sApp As String In uninstall.GetSubKeyNames()
                For Each pValueNames As String In uninstall.OpenSubKey(sApp).GetValueNames()
                    If pValueNames.equalsIgnoreCase("DisplayName") Then
                        InstalledAppsDisplayName.Add(uninstall.OpenSubKey(sApp).GetValue(pValueNames).ToString())
                    End If
                Next
            Next
        End If




        Return InstalledAppsDisplayName.Contains(pPartOfAppDisplayName, New LikeIgnoreCaseComparer())


    End Function


#End Region


#Region "Processes"
    ''' <summary>
    ''' Checks if a process is running
    ''' </summary>
    ''' <param name="ProcessNameWithEXE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isProcessRunning(ByVal ProcessNameWithEXE As String) As Boolean
        Return Process.GetProcessesByName(getFileNameWithoutExtension(ProcessNameWithEXE)).Length > 0

    End Function

    ''' <summary>
    ''' Kill process if it is available
    ''' </summary>
    ''' <param name="ProcessNameWithEXE"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function KillProcess(ByVal ProcessNameWithEXE As String) As Boolean
        Try


            If isProcessRunning(ProcessNameWithEXE) Then
                Dim P As Process() = Process.GetProcessesByName(getFileNameWithoutExtension(ProcessNameWithEXE))
                REM There might be more than one instances of it
                For Each proc As Process In P
                    proc.Kill()
                Next

            End If

            Return True


        Catch ex As Exception
            Program.Logger.Log(New EException(ex))
        End Try
        Return False

    End Function

#End Region

#Region "Services"
    ''' <summary>
    ''' Checks if a particular service is installed on this OS
    ''' </summary>
    ''' <param name="serviceName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isServiceAvailable(ByVal serviceName As String) As Boolean

        Try
            Dim aService As New ServiceProcess.ServiceController(serviceName)

            REM The next line will generate error if it is not available
            Debug.Print(aService.Status.ToString())
            Return True
        Catch ex As Exception

            Return False

        End Try

    End Function

    ''' <summary>
    ''' gets a service status. It returns ServiceProcess.ServiceControllerStatus.StopPending  if the service is not found or error occured
    ''' </summary>
    ''' <param name="serviceName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getServiceStatus(ByVal serviceName As String) As ServiceProcess.ServiceControllerStatus

        Try
            Dim aService As New ServiceProcess.ServiceController(serviceName)


            Return aService.Status

        Catch ex As Exception

            Return ServiceProcess.ServiceControllerStatus.StopPending

        End Try

    End Function


    ''' <summary>
    ''' Stops a service if it exists. I guess you should invoke this as an administrator 
    ''' NB: This is Sync Method. It doesnt return until it is done:)
    ''' </summary>
    ''' <param name="serviceName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function stopService(ByVal serviceName As String) As Boolean

        If Not OperatingSystem.isServiceAvailable(serviceName) Then Return False

        Dim aService As New ServiceProcess.ServiceController(serviceName)

        If getServiceStatus(serviceName) <> ServiceProcess.ServiceControllerStatus.Stopped Then
            Try
                aService.Stop()
            Catch ex As Exception
                Return False  REM Cant stop Service
            End Try

        End If

        While getServiceStatus(serviceName) <> ServiceProcess.ServiceControllerStatus.Stopped

            REM Try Again after 1s
            Threading.Thread.Sleep(1000)

        End While

        Return True

    End Function





    ''' <summary>
    ''' Starts a service if it exists. I guess you should invoke this as an administrator 
    ''' NB: This is Sync Method. It doesnt return until it is done:)
    ''' </summary>
    ''' <param name="serviceName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function startService(ByVal serviceName As String) As Boolean

        If Not OperatingSystem.isServiceAvailable(serviceName) Then Return False

        Dim aService As New ServiceProcess.ServiceController(serviceName)

        If getServiceStatus(serviceName) <> ServiceProcess.ServiceControllerStatus.Running Then
            Try
                aService.Start()
            Catch ex As Exception
                Return False  REM Cant start Service
            End Try

        End If

        While getServiceStatus(serviceName) <> ServiceProcess.ServiceControllerStatus.Running

            REM Try Again after 1s
            Threading.Thread.Sleep(1000)

        End While

        Return True

    End Function

#End Region






#Region "DisablingOSFunctionality"

    ''' <summary>
    ''' Allow or Disallow the Add and Remove Page
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisAllowAddAndRemoveProgramPage(Optional ByVal setValue As Boolean = True)
        ''        System Key: [HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\
        ''Uninstall]
        ''Data Type: REG_DWORD (DWORD Value)
        ''Value Data: (0 = default, 1 = enable restriction)

        CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                  "Software\Microsoft\Windows\CurrentVersion\Policies",
                                  "Uninstall"
                                  )


        CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                       "Software\Microsoft\Windows\CurrentVersion\Policies\Uninstall",
                                       "NoAddRemovePrograms", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )


    End Sub


    ''' <summary>
    ''' Disable or Enable Microsoft Installer
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableMSI(Optional ByVal setValue As Boolean = True)


        Dim value As Long = 2
        If Not setValue Then value = 0

        ' ''Code:
        ' ''My.Computer.Registry.SetValue("HKLM\Software\Policies\Microsoft\Windows\Installer",
        REM  "DisableMSI", "1", Microsoft.Win32.RegistryValueKind.DWord)


        CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                   "Software\Policies\Microsoft\Windows",
                                   "Installer"
                                   )

        CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                       "Software\Policies\Microsoft\Windows\Installer",
                                       "DisableMSI", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )


    End Sub


    ''' <summary>
    ''' Disable or Enable USB Devices
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableUSBDevices(Optional ByVal setValue As Boolean = True)

        ' ''"Hidden Files and Folders" Option
        REM 3- enable  4-Disable
        ' ''Code:
        ' ''    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\usbstor"
        ' ''"Start", "3", Microsoft.Win32.RegistryValueKind.DWord)

        CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                      "SYSTEM\CurrentControlSet\Services",
                                      "usbstor"
                                      )

        Dim value As Long = 4
        If Not setValue Then value = 3

        CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                      "SYSTEM\CurrentControlSet\Services\usbstor",
                                      "Start", CStr(ELong.valueOf(value)),
                                      Microsoft.Win32.RegistryValueKind.DWord
                                      )


    End Sub


    ''' <summary>
    ''' Show or Hide Hidden Files and Folders
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub ShowHiddenFilesAndFolder(Optional ByVal setValue As Boolean = True)

        ' ''"Hidden Files and Folders" Option
        REM 1- show  2-Don't show
        ' ''Code:
        ' ''    HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
        ' ''My.Computer.Registry.SetValue(regloc, "Hidden", "0", Microsoft.Win32.RegistryValueKind.DWord)



        Dim value As Long = 1
        If Not setValue Then value = 2

        CreateAKey(RegEditKeys.HKEY_CURRENT_USER,
                                      "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
                                      "Hidden", CStr(ELong.valueOf(value)),
                                      Microsoft.Win32.RegistryValueKind.DWord
                                      )


    End Sub



    ''' <summary>
    ''' Disable or Enable System Restore
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableSystemRestore(Optional ByVal setValue As Boolean = True)
        ' ''System Restore

        ' ''Code:
        ' ''"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore", 
        REM "DisableSR", "1", Microsoft.Win32.RegistryValueKind.DWord)

        CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                     "SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore",
                                     "DisableSR", CStr(ELong.valueOf(setValue)),
                                     Microsoft.Win32.RegistryValueKind.DWord
                                     )

        If setValue Then
            REM Close it if it is already on
            OperatingSystem.KillProcess(SYSTEM_RESTORE_PROCESS)
        End If
    End Sub


    ''' <summary>
    ''' Disable or Enable a TaskManager
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableTaskManager(Optional ByVal setValue As Boolean = True)
        ''"HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskMgr", "1", Microsoft.Win32.RegistryValueKind.DWord)
        If getRegistrySubKeyFolder(RegEditKeys.HKEY_CURRENT_USER,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies"
                                   ) Is Nothing Then Return


        If getRegistrySubKeyFolder(RegEditKeys.HKEY_CURRENT_USER,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies\System"
                                   ) Is Nothing Then

            CreateASubKey(RegEditKeys.HKEY_CURRENT_USER,
                                       "Software\Microsoft\Windows\CurrentVersion\Policies",
                                       "System"
                                       )


        End If


        CreateAKey(RegEditKeys.HKEY_CURRENT_USER,
                                       "Software\Microsoft\Windows\CurrentVersion\Policies\System",
                                       "DisableTaskMgr", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )

        If setValue Then
            REM Close it if it is already on
            OperatingSystem.KillProcess(TASK_MANAGER_PROCESS)
        End If
    End Sub



    ''' <summary>
    ''' Disable or Enable a Registry Tools like RegEdit
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableRegistryTools(Optional ByVal setValue As Boolean = True)
        ' ''Registry Tools

        ' ''Code:
        ' ''My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\System", 
        REM "DisableRegistryTools", "1", Microsoft.Win32.RegistryValueKind.DWord)



        REM This path should be on the system
        If getRegistrySubKeyFolder(RegEditKeys.HKEY_CURRENT_USER,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies"
                                   ) Is Nothing Then Return


        CreateASubKey(RegEditKeys.HKEY_CURRENT_USER,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies",
                                   "System"
                                   )

        CreateAKey(RegEditKeys.HKEY_CURRENT_USER,
                                       "Software\Microsoft\Windows\CurrentVersion\Policies\System",
                                       "DisableRegistryTools", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )

        If setValue Then
            REM Close it if it is already on
            OperatingSystem.KillProcess(REGISTRY_TOOLS_PROCESS)
        End If
    End Sub



    ''' <summary>
    ''' Disable or Enable a Command Prompt
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableCommandPrompt(Optional ByVal setValue As Boolean = True)
        ' ''Command Prompt

        ' ''Code:
        ' ''My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Policies\Microsoft\Windows\System",
        REM  "DisableCMD", "1", Microsoft.Win32.RegistryValueKind.DWord)


        CreateASubKey(RegEditKeys.HKEY_CURRENT_USER,
                                   "Software\Policies\Microsoft\Windows",
                                   "System"
                                   )

        CreateAKey(RegEditKeys.HKEY_CURRENT_USER,
                                       "Software\Policies\Microsoft\Windows\System",
                                       "DisableCMD", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )


        If setValue Then
            REM Close it if it is already on
            OperatingSystem.KillProcess(COMMAND_PROMPT_PROCESS)
        End If
    End Sub


    ''' <summary>
    ''' Disable or Enable a System Date Time Changing
    ''' </summary>
    ''' <param name="setValue"></param>
    ''' <remarks></remarks>
    Public Shared Sub DisableSytemDateTime(Optional ByVal setValue As Boolean = True)
        ' ''Registry Tools

        ' ''Code:
        ' ("HKEY_LOCAL_MACHINE\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", 
        REM "NoControlPanel", "1", Microsoft.Win32.RegistryValueKind.DWord)



        REM This path should be on the system
        If getRegistrySubKeyFolder(RegEditKeys.HKEY_LOCAL_MACHINE,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies"
                                   ) Is Nothing Then Return


        CreateASubKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                   "Software\Microsoft\Windows\CurrentVersion\Policies",
                                   "Explorer"
                                   )

        CreateAKey(RegEditKeys.HKEY_LOCAL_MACHINE,
                                       "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer",
                                       "NoControlPanel", CStr(ELong.valueOf(setValue)),
                                       Microsoft.Win32.RegistryValueKind.DWord
                                       )


    End Sub



#End Region


#Region "Screen Capture"
    ''' <summary>
    ''' Get Full Screenshot With Cursor. NB: No RawFormat Data
    ''' </summary>
    ''' <param name="_cursor">The cursor of the calling form</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getScreenShotWithCursor(ByRef _cursor As Cursor) As Bitmap
        Dim b As Bitmap = New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)

        Dim g As Graphics = Graphics.FromImage(b)
        g.CopyFromScreen(My.Computer.Screen.Bounds.Location, New Point(0, 0), My.Computer.Screen.Bounds.Size)
        _cursor.Draw(g, New Rectangle(Cursor.Position, _cursor.Size))
        g.Dispose()

        Return b
    End Function

    ''' <summary>
    ''' Return Screenshot without a cursor. NB: No RawFormat Data
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getScreenShotWithoutCursor() As Bitmap

        With Screen.PrimaryScreen.Bounds

            Dim screenSize As Size = New Size(.Width, .Height)

            Dim screenGrab As New Bitmap(.Width, .Height)

            Dim g As Graphics = Graphics.FromImage(screenGrab)

            g.CopyFromScreen(New Point(0, 0), New Point(0, 0), screenSize)

            g.Dispose()

            Return screenGrab

        End With


    End Function

#End Region


#Region "Certificates Management"

    ''' <summary>
    ''' Checks if the specified certificate file is installed in the specified location
    ''' </summary>
    ''' <param name="certificateFullPath">The full path to the certificate locally</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Shared Function isCertificatePresentOn(
                                                 ByVal certificateFullPath As String,
                                                 Optional ByVal certificateStoreLocation As StoreLocation =
                                                 StoreLocation.CurrentUser,
                                                 Optional ByVal certificateStore As StoreName = StoreName.Root
                                                                                                      ) As Boolean

        Try

            Return (
                fetchCertificateOn(
                            New StringBuilder(certificateFullPath),
                            certificateStoreLocation, certificateStore
                            ) IsNot Nothing
                        )



        Catch ex As Exception

        End Try

        Return False

    End Function


    ''' <summary>
    ''' Checks if a Certificate has not expired
    ''' </summary>
    ''' <param name="certificateFullPath">The full path to the certificate locally</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Shared Function isCertificateValid(
                                               ByVal certificateFullPath As String,
                                               Optional ByVal certificateStoreLocation As StoreLocation =
                                               StoreLocation.CurrentUser,
                                               Optional ByVal certificateStore As StoreName = StoreName.Root
                                                                                                    ) As Boolean

        Try

            If isCertificatePresentOn(certificateFullPath, certificateStoreLocation, certificateStore) Then

                Return CDate(
                    New X509Certificate2(certificateFullPath).GetExpirationDateString
                    ) > Now

            End If

        Catch ex As Exception

        End Try

        Return False

    End Function


#Region "Certificate Add Monitor"
    REM I dont know if it is same information that is displayed on Delete Action
    REM Also NOTE: This is not tested yet on other OS.. Just Windows 7 Only

    ''    Private NotInheritable Class clsCertAddMonitor
    ''        Implements IDisposable



    ''        Private MyThreadManager As clsMultiThreading = Nothing
    ''        Private Const Caption As String = "Security Warning"
    ''        Private isDone As Boolean = False
    ''        Private Const MaximumRunTime As Byte = 10 REM it should more than this else there is a problem
    ''        Private RuntimeCount As Byte = 0


    ''        Sub New()

    ''            Me.MyThreadManager = New clsMultiThreading()
    ''            Me.MyThreadManager.PlaceThreadInInfiniteLoop(True, AddressOf Me.Check_If_Window_is_Out, 500)

    ''        End Sub


    ''        Private Sub Check_If_Window_is_Out(ByVal Sender As Threading.Thread)

    ''            REM Incase the class is done and the thread is not terminated
    ''            If Me.isDone Then Return

    ''            Try
    ''                Me.RuntimeCount += 1
    ''                Dim CertWindowHwnd As IntPtr = GetWindowHandle__UsingCaption(Caption)

    ''                If Not IsNothing(CertWindowHwnd) Then

    ''                    basMain.MyLogFile.Log("CertWindow Handle: " & CertWindowHwnd.ToInt32, "clsOperatingSystem",
    ''                                   "Check_IF_Window_is_out")

    ''                    If CertWindowHwnd.ToInt32 <> 0 Then _
    ''                        AppActivate(
    ''                                GetProcessID_From_Handle(CertWindowHwnd)
    ''                                )

    ''                    REM Since am in a dll I need to use send wait
    ''                    REM if am in an application that can respond to user's input then i can use send method
    ''                    SendKeys.SendWait("{TAB}")
    ''                    SendKeys.SendWait("{ENTER}")

    ''                    basMain.MyLogFile.Log("Sending keys to process ID", "clsOperatingSystem",
    ''                                   "Check_IF_Window_is_out")
    ''                    If IsNothing(GetWindowHandle__UsingCaption(Caption)) Then
    ''                        basMain.MyLogFile.Log("GetWindow Handle returned Nothing: on " & Caption)
    ''END_THREAD:
    ''                        Me.MyThreadManager.ExitAsyncThread(Sender)
    ''                        Me.isDone = True

    ''                        basMain.MyLogFile.Log("Thread is done")

    ''                    End If



    ''                End If


    ''                If Me.RuntimeCount > MaximumRunTime Then GoTo END_THREAD
    ''                basMain.MyLogFile.Log("I am checking  out window no of times: " & RuntimeCount)

    ''            Catch ex As Threading.ThreadAbortException
    ''                basMain.MyLogFile.Log("ThreadAbortException Occurred Check if window is out", "clsOperatingSystem",
    ''                                       "Check_IF_Window_is_out", ex.Message)
    ''                Me.isDone = True
    ''            Catch ex As Exception
    ''                basMain.MyLogFile.Log("Normal Exception Occurred Check if window is out", "clsOperatingSystem",
    ''                                       "Check_IF_Window_is_out", ex.Message)
    ''                REM  Am not expecting this error thou
    ''                Me.isDone = True

    ''            End Try


    ''        End Sub


    ''#Region "IDisposable Support"
    ''        Private disposedValue As Boolean ' To detect redundant calls

    ''        ' IDisposable
    ''        Protected Sub Dispose(ByVal disposing As Boolean)
    ''            If Not Me.disposedValue Then
    ''                If disposing Then
    ''                    ' TODO: dispose managed state (managed objects).
    ''                    If Me.MyThreadManager IsNot Nothing Then
    ''                        Me.MyThreadManager.AbortAll()
    ''                        Me.MyThreadManager = Nothing
    ''                    End If
    ''                End If

    ''                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
    ''                ' TODO: set large fields to null.
    ''            End If
    ''            Me.disposedValue = True
    ''        End Sub

    ''        ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    ''        'Protected Overrides Sub Finalize()
    ''        '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    ''        '    Dispose(False)
    ''        '    MyBase.Finalize()
    ''        'End Sub

    ''        ' This code added by Visual Basic to correctly implement the disposable pattern.
    ''        Public Sub Dispose() Implements IDisposable.Dispose
    ''            ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    ''            Dispose(True)
    ''            GC.SuppressFinalize(Me)
    ''        End Sub
    ''#End Region

    ''    End Class

#End Region


    ''' <summary>
    ''' Adds a certificate to the specified location but it does not do it if it is already present and valid
    ''' </summary>
    ''' <param name="certificateFullPath">The full path to the certificate locally</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Shared Function addCertificateOn(
                                               ByVal certificateFullPath As String,
                                               Optional ByVal certificateStoreLocation As StoreLocation =
                                               StoreLocation.CurrentUser,
                                               Optional ByVal certificateStore As StoreName = StoreName.Root
                                                                                                    ) As Boolean



        Try
            If Not isCertificateValid(certificateFullPath, certificateStoreLocation, certificateStore) Then

                Dim store As New X509Store(certificateStore, certificateStoreLocation)
                store.Open(OpenFlags.ReadWrite)


                REM I have to recheck to make sure the cert file exist
                If System.IO.File.Exists(certificateFullPath) Then


                    If getOSType() = MicrosoftOS.WINDOWS_XP Then
                        REM Class -> Multi-threading is bad on XP
                        REM but direct from application still works fine
                        store.Add(New X509Certificate2(certificateFullPath))


                    Else
                        REM Others for now works fine

                        ''REM Add Monitor here for auto-response
                        ''Dim Monitor As New clsCertAddMonitor()

                        store.Add(New X509Certificate2(certificateFullPath))

                        ''Monitor.Dispose()
                        ''Monitor = Nothing REM Free Space

                    End If

                Else
                    Return False
                End If




                store.Close()

                REM Am collapsing the two into 1 Since in both cases it returns true
                '' ''    Return True

                '' ''Else

                '' ''    REM It is already present
                '' ''    Return True

            End If

            Return True

        Catch ex As Exception
            REM Incase user selected NO
            Program.Logger.Log(
                New EException(
                    "Exception under clsOperatingSystem addCertificateOn", ex)
                )

        End Try


        Return False

    End Function


    ''' <summary>
    ''' Delete a certificate to the specified location but it does not do it if it is NOT already present 
    ''' </summary>
    ''' <param name="certificateFullPath">The full path to the certificate locally</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Shared Function deleteCertificateOn(
                                               ByVal certificateFullPath As StringBuilder,
                                               Optional ByVal certificateStoreLocation As StoreLocation =
                                               StoreLocation.CurrentUser,
                                               Optional ByVal certificateStore As StoreName = StoreName.Root
                                                                                                    ) As Boolean



        Try
            Dim Cert As X509Certificate2 = fetchCertificateOn(certificateFullPath, certificateStoreLocation, certificateStore)

            If Cert IsNot Nothing Then

                Dim store As New X509Store(certificateStore, certificateStoreLocation)
                store.Open(OpenFlags.ReadWrite)


                store.Remove(Cert)


                store.Close()
            Else

                REM It is already Deleted
                Return True

            End If

        Catch ex As Exception
            REM Incase use selected NO

        End Try


        Return False

    End Function


    ''' <summary>
    ''' Delete a certificate to the specified location but it does not do it if it is NOT already present 
    ''' </summary>
    ''' <param name="certificateName">Example "CN= EPRO CYBERSOFT"</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Shared Function deleteCertificateOn(
                                               ByVal certificateName As String,
                                               Optional ByVal certificateStoreLocation As StoreLocation =
                                               StoreLocation.CurrentUser,
                                               Optional ByVal certificateStore As StoreName = StoreName.Root
                                                                                                    ) As Boolean



        Try
            Dim Cert As X509Certificate2 = fetchCertificateOn(certificateName, certificateStoreLocation,
                                                              certificateStore)

            If Cert IsNot Nothing Then

                Dim store As New X509Store(certificateStore, certificateStoreLocation)
                store.Open(OpenFlags.ReadWrite)


                store.Remove(Cert)


                store.Close()
            Else

                REM It is already Deleted
                Return True

            End If

        Catch ex As Exception
            REM Incase use selected NO

        End Try


        Return False

    End Function



    ''' <summary>
    ''' fetch a certificate Locally
    ''' </summary>
    ''' <param name="certificateFullPath">The full path to the certificate locally</param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Overloads Shared Function fetchCertificateOn(
                                               ByVal certificateFullPath As String
                                               ) As X509Certificate2



        Try

            REM I have to recheck to make sure the cert file exist
            If System.IO.File.Exists(certificateFullPath) Then


                Return (New X509Certificate2(certificateFullPath))


            End If


        Catch ex As Exception


        End Try


        Return Nothing

    End Function

    ''' <summary>
    ''' fetch a certificate on store using local certificate match
    ''' </summary>
    ''' <param name="CertificateFullPath">Full path of the local certificate</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Overloads Shared Function fetchCertificateOn(
                                                       ByVal CertificateFullPath As StringBuilder,
                                                ByVal certificateStoreLocation As StoreLocation,
                                                ByVal certificateStore As StoreName
                                               ) As X509Certificate2



        Try

            Dim Cert As X509Certificate2 = fetchCertificateOn(
                                          CertificateFullPath.ToString)

            If Cert Is Nothing Then Return Cert


            Return fetchCertificateOn(Cert.Subject, certificateStoreLocation, certificateStore)



        Catch ex As Exception


        End Try


        Return Nothing

    End Function

    ''' <summary>
    ''' fetch a certificate on store using name
    ''' </summary>
    ''' <param name="CertificateName">Example "CN= EPRO CYBERSOFT"</param>
    ''' <param name="certificateStoreLocation"></param>
    ''' <param name="certificateStore"></param>
    ''' <returns></returns>
    ''' <remarks>NB: Return false might be due to invalid certificate file</remarks>
    Public Overloads Shared Function fetchCertificateOn(
                                                       ByVal CertificateName As String,
                                                ByVal certificateStoreLocation As StoreLocation,
                                                ByVal certificateStore As StoreName
                                               ) As X509Certificate2



        Try

            Dim store As New X509Store(certificateStore, certificateStoreLocation)
            store.Open(OpenFlags.ReadOnly)


            Dim storecollection As X509Certificate2Collection = CType(store.Certificates, X509Certificate2Collection)

            For Each x509 As X509Certificate2 In storecollection

                If x509.Subject.ToLower = CertificateName.ToLower Then store.Close() : Return x509
                Debug.Print("{0} isNOT {1}", x509.Subject.ToLower, CertificateName.ToLower)

            Next x509


            'Close the store.
            store.Close()


        Catch ex As Exception


        End Try


        Return Nothing

    End Function


#End Region

#Region "Start Up App Managements"

    ''' <summary>
    ''' Add an Executor to run at start up
    ''' </summary>
    ''' <param name="ExecutorPath"></param>
    ''' <remarks></remarks>
    Public Shared Sub addExecutorTo_RunAtStartUp(ByVal ExecutorPath As String)

        CreateAKey(
             RegEditKeys.HKEY_LOCAL_MACHINE,
                "SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
                getFileNameWithoutExtension(ExecutorPath),
                ExecutorPath)


    End Sub

    ''' <summary>
    ''' Removes an executor from run at start up
    ''' </summary>
    ''' <param name="ExecutorPath"></param>
    ''' <remarks></remarks>
    Public Shared Sub removeExecutorFrom_RunAtStartUp(ByVal ExecutorPath As String)
        DeleteAKey(
            RegEditKeys.HKEY_LOCAL_MACHINE,
               "SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
               getFileNameWithoutExtension(ExecutorPath)
                )
    End Sub

#End Region


#End Region



#End Region



End Class
