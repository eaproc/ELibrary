Option Strict Off


''' <summary>
''' perform Firewall Operations
''' </summary>
''' <remarks></remarks>
Public Class Firewall


#Region "Enums"

    Public Enum TrafficProtocols
        TCP
        UDP
    End Enum

#End Region







    ''' <summary>
    ''' Disable or Enable Firewall
    ''' </summary>
    ''' <param name="Enabled"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function setFireWallEnable(ByVal Enabled As Boolean) As Boolean
        'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile
        'EnableFirewall=1

        Return True
    End Function




    ''' <summary>
    ''' Disable or Enable Firewall
    ''' </summary>
    ''' <param name="Enabled"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function setExceptionsEnable(ByVal Enabled As Boolean) As Boolean
        'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile
        'DoNotAllowExceptions

        Return True
    End Function




    ''' <summary>
    ''' Create an Inbound Port Exception. NB: Winxp allows only 1 port per declaration
    ''' </summary>
    ''' <param name="sFirewallName"></param>
    ''' <param name="sFirewallDescription"></param>
    ''' <param name="sPort">Currently Just TCP Protocol</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function createInboundPortExceptions(ByVal sFirewallName As String,
                                                ByVal sFirewallDescription As String,
                                                ByVal sPort As String,
                                                Optional pProtocol As TrafficProtocols = TrafficProtocols.TCP) As Boolean



        'Set Constants
        Const NET_FW_IP_PROTOCOL_UDP = 17
        Const NET_FW_IP_PROTOCOL_TCP = 6

        Const NET_FW_SCOPE_ALL = 0
        Const NET_FW_SCOPE_LOCAL_SUBNET = 1

        'Declare variables
        Dim errornum As Int32

        ' Create the firewall manager object.
        Dim fwMgr As Object
        fwMgr = CreateObject("HNetCfg.FwMgr")

        ' Get the current profile for the local firewall policy.
        Dim profile As Object
        profile = fwMgr.LocalPolicy.CurrentProfile


        Dim port As Object
        port = CreateObject("HNetCfg.FWOpenPort")

        port.Name = sFirewallName
        If pProtocol = TrafficProtocols.TCP Then
            port.Protocol = NET_FW_IP_PROTOCOL_TCP
        ElseIf pProtocol = TrafficProtocols.UDP Then
            port.Protocol = NET_FW_IP_PROTOCOL_UDP

        End If

        port.Port = Val(sPort)

        'If using Scope, don't use RemoteAddresses
        port.Scope = NET_FW_SCOPE_ALL
        'Use this line to scope the port to Local Subnet only
        'port.Scope = NET_FW_SCOPE_LOCAL_SUBNET

        port.Enabled = True
        'Use this line instead if you want to add the port, but disabled
        'port.Enabled = FALSE

        On Error Resume Next
        profile.GloballyOpenPorts.Add(port)
        errornum = Err.Number

        If errornum <> 0 Then
            Program.Logger.Print("Adding the port failed. Error Number: " & errornum)
            Return False

        End If




        '' ''REM Currently Editing Registry isnt Working with FireWall ********************************************
        '' ''REM Protocol TCP = 6
        '' ''REM INbound and outbound is same with win xp and below
        '' ''REM Check the version of OS
        '' ''REM Edit the Registry
        '' ''Dim osVer As MicrosoftOS = getOSType

        '' ''Select Case osVer

        '' ''    Case Is = MicrosoftOS.WINDOWS_XP

        '' ''        REM HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List
        '' ''        REM Create this if not available
        '' ''        REM StandardProfile\GloballyOpenPorts\List
        '' ''        REM 1433:TCP
        '' ''        REM 1433:TCP:*:Enabled:SQL Ports Allowed
        '' ''        REM * Means All IP

        '' ''        REM Create StandardProfile FolderKey
        '' ''        If getRegistrySubKeyFolder(
        '' ''                                           RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                           "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile"
        '' ''                                           ) Is Nothing Then
        '' ''            REM Try to create it
        '' ''            'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
        '' ''            If CreateASubKey(
        '' ''                                                RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                                "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy",
        '' ''                                                "StandardProfile"
        '' ''                                                ) Is Nothing Then
        '' ''                Return False

        '' ''            End If

        '' ''        End If


        '' ''        REM Create GloballyOpenPorts FolderKey
        '' ''        If getRegistrySubKeyFolder(
        '' ''                                           RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                           "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts"
        '' ''                                           ) Is Nothing Then
        '' ''            REM Try to create it
        '' ''            'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
        '' ''            If CreateASubKey(
        '' ''                                                RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                                "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile",
        '' ''                                                "GloballyOpenPorts"
        '' ''                                                ) Is Nothing Then
        '' ''                Return False

        '' ''            End If

        '' ''        End If



        '' ''        REM Create List FolderKey
        '' ''        If getRegistrySubKeyFolder(
        '' ''                                           RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                           "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List"
        '' ''                                           ) Is Nothing Then
        '' ''            REM Try to create it
        '' ''            'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
        '' ''            If CreateASubKey(
        '' ''                                                RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                                "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts",
        '' ''                                                "List"
        '' ''                                                ) Is Nothing Then
        '' ''                Return False

        '' ''            End If

        '' ''        End If




        '' ''        If Not IsNumeric(sPort) Then Return False REM ONLY Numbers are allowed

        '' ''        REM Try to check the number range ... not higher that short int




        '' ''        REM Try Create Key
        '' ''        REM HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List
        '' ''        REM 1433:TCP
        '' ''        REM 1433:TCP:*:Enabled:SQL Ports Allowed
        '' ''        Return CreateAKey(
        '' ''                                                RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                               "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List",
        '' ''                                               String.Format(
        '' ''                                                   "{0}:TCP",
        '' ''                                                   sPort
        '' ''                                                   ),
        '' ''                                               String.Format("{0}:TCP:*:Enabled:{1}",
        '' ''                                                             sPort, sFirewallName
        '' ''                                                             )
        '' ''                                                )





        '' ''    Case Is = MicrosoftOS.WINDOWS_7
        '' ''        'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
        '' ''        If getRegistrySubKeyFolder(
        '' ''                                            RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                            "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules"
        '' ''                                            ) Is Nothing Then
        '' ''            REM Try to create it
        '' ''            'HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
        '' ''            If CreateASubKey(
        '' ''                                                RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                                "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy",
        '' ''                                                "FirewallRules"
        '' ''                                                ) Is Nothing Then
        '' ''                Return False

        '' ''            End If

        '' ''        End If




        '' ''        REM Try create the key
        '' ''        REM v2.10|Action=Allow|Active=TRUE|Dir=In|Protocol=6|LPort=1403|Name=Crazy Firewall|Desc=Default Port for SQL Server Remote Connections|
        '' ''        Return CreateAKey(
        '' ''                                                 RegEditKeys.HKEY_LOCAL_MACHINE,
        '' ''                                                "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules",
        '' ''                                                sFirewallName,
        '' ''                                                String.Format("v2.10|Action=Allow|Active=TRUE|Dir=In|Protocol=6|LPort={0}|Name={1}|Desc={2}|",
        '' ''                                                              sPort, sFirewallName, sFirewallDescription)
        '' ''                                                 )



        '' ''End Select

        Return True
    End Function




    ''' <summary>
    ''' Checks if this FireWallName Has been configured
    ''' </summary>
    ''' <param name="sFirewallName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isInboundPortExceptionValid(ByVal sFirewallName As String, ByVal sPort As String) As Boolean

        REM Protocol TCP = 6
        REM INbound and outbound is same with win xp and below
        REM Check the version of OS
        REM Edit the Registry
        Dim osVer As OperatingSystem.MicrosoftOS = OperatingSystem.getOSType

        Select Case osVer

            Case Is = OperatingSystem.MicrosoftOS.WINDOWS_XP
                REM For now am not sure about windows xp

                REM Try Create Key
                REM HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List
                REM 1433:TCP
                REM 1433:TCP:*:Enabled:SQL Ports Allowed
                Return Registry.searchDataLikeInSubFolder(
                                                        Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                       "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\GloballyOpenPorts\List",
                                                        sFirewallName, sPort
                                                        )



                'Return True

                'Case Is = MicrosoftOS.WINDOWS_7
            Case Else
                REM Assuming Vista and the rest works like this

                REM Try create the key
                REM v2.10|Action=Allow|Active=TRUE|Dir=In|Protocol=6|LPort=1403|Name=Crazy Firewall|Desc=Default Port for SQL Server Remote Connections|


                Return Registry.searchDataLikeInSubFolder(
                                                         Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                        "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules",
                                                        sFirewallName, sPort
                                                         )



        End Select



        Return False
    End Function



    ''' <summary>
    ''' Only for 64 Bit Systems and Windows XP. Adds ICMPv4 to Firewall
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddICMPv4Exception() As Boolean

        If OperatingSystem.getOSType = OperatingSystem.MicrosoftOS.WINDOWS_XP Then

            Try
                'SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\IcmpSettings",
                '"AllowInboundEchoRequest"

                If Registry.CreateASubKey(Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                                 "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile",
                                                                 "IcmpSettings"
                                                                  ) IsNot Nothing Then

                    If Registry.CreateAKey(Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                               "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\IcmpSettings",
                                                                "AllowInboundEchoRequest", "1",
                                                                Microsoft.Win32.RegistryValueKind.DWord
                                                                 ) Then
                        Return True
                    End If

                End If

                Return False

            Catch ex As Exception
                Return False
            End Try
        Else

            Try

                Dim CurrentProfiles As Object

                Const NET_FW_IP_PROTOCOL_ICMPv4 = 1
                'Const NET_FW_IP_PROTOCOL_ICMPv6 = 58

                'Action
                Const NET_FW_ACTION_ALLOW = 1

                ' Create the FwPolicy2 object.
                Dim fwPolicy2 As Object
                fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

                ' Get the Rules object
                Dim RulesObject As Object
                RulesObject = fwPolicy2.Rules

                CurrentProfiles = fwPolicy2.CurrentProfileTypes

                'Create a Rule Object.
                Dim NewRule As Object
                NewRule = CreateObject("HNetCfg.FWRule")

                NewRule.Name = "ICMP_Rule"
                NewRule.Description = "Allow ICMP network traffic"
                NewRule.Protocol = NET_FW_IP_PROTOCOL_ICMPv4
                '            NewRule.IcmpTypesAndCodes = "*:*"
                NewRule.IcmpTypesAndCodes = "2:*,3:*,4:*,5:*,8:*,9:*,10:*,11:*,12:*,13:*,17:*"
                NewRule.Enabled = True
                NewRule.Grouping = "@firewallapi.dll,-23255"
                NewRule.Profiles = CurrentProfiles
                NewRule.Action = NET_FW_ACTION_ALLOW

                'Add a new rule
                RulesObject.Add(NewRule)

                Return True

            Catch ex As Exception

                Return False

            End Try



        End If


    End Function

    ''' <summary>
    ''' Checks if the ICMPv4 is Already Enabled
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isICMPv4Allowed() As Boolean
        If OperatingSystem.getOSType = OperatingSystem.MicrosoftOS.WINDOWS_XP Then

            Return Registry.searchKeyLikeInSubFolder(
                                                        Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                       "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\IcmpSettings",
                                                       "AllowInboundEchoRequest"
                                                         )
        Else
            Return Registry.searchDataLikeInSubFolder(
                                                        Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                       "SYSTEM\CurrentControlSet\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules",
                                                       "ICMP_Rule"
                                                         )
        End If


    End Function





    ''' <summary>
    ''' Allows the Specified Executor to access UDP Ports
    ''' </summary>
    ''' <param name="pApplicationName"></param>
    ''' <param name="pExecutorFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddApplicationUDPEnableException(ByVal pApplicationName As String,
                                                            ByVal pExecutorFullPath As String) As Boolean

        Try



            Dim CurrentProfiles As Object

            ' Protocol
            Const NET_FW_IP_PROTOCOL_UDP = 17


            'Action
            Const NET_FW_ACTION_ALLOW = 1

            ' Create the FwPolicy2 object.
            Dim fwPolicy2 As Object
            fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

            ' Get the Rules object
            Dim RulesObject As Object
            RulesObject = fwPolicy2.Rules

            CurrentProfiles = fwPolicy2.CurrentProfileTypes

            'Create a Rule Object.
            Dim NewRule As Object
            NewRule = CreateObject("HNetCfg.FWRule")

            NewRule.Name = pApplicationName
            NewRule.Description = "Allow my application network UDP traffic"
            NewRule.Applicationname = pExecutorFullPath
            NewRule.Protocol = NET_FW_IP_PROTOCOL_UDP
            'NewRule.LocalPorts = 4000
            NewRule.Enabled = True
            NewRule.Grouping = "@firewallapi.dll,-23255"
            NewRule.Profiles = CurrentProfiles
            NewRule.Action = NET_FW_ACTION_ALLOW

            'Add a new rule
            RulesObject.Add(NewRule)

            Return True



        Catch ex As Exception
            Program.Logger.Print(ex)
        End Try


        Return False

    End Function


    ''' <summary>
    ''' Allows the Specified Executor to access TCP Ports
    ''' </summary>
    ''' <param name="ApplicationName"></param>
    ''' <param name="ExecutorFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AddApplicationTCPEnableException(ByVal ApplicationName As String,
                                                            ByVal ExecutorFullPath As String) As Boolean

        Try



            Dim CurrentProfiles As Object

            ' Protocol
            Const NET_FW_IP_PROTOCOL_TCP = 6

            'Action
            Const NET_FW_ACTION_ALLOW = 1

            ' Create the FwPolicy2 object.
            Dim fwPolicy2 As Object
            fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

            ' Get the Rules object
            Dim RulesObject As Object
            RulesObject = fwPolicy2.Rules

            CurrentProfiles = fwPolicy2.CurrentProfileTypes

            'Create a Rule Object.
            Dim NewRule As Object
            NewRule = CreateObject("HNetCfg.FWRule")

            NewRule.Name = ApplicationName
            NewRule.Description = "Allow my application network tcp traffic"
            NewRule.Applicationname = ExecutorFullPath
            NewRule.Protocol = NET_FW_IP_PROTOCOL_TCP
            'NewRule.LocalPorts = 4000
            NewRule.Enabled = True
            NewRule.Grouping = "@firewallapi.dll,-23255"
            NewRule.Profiles = CurrentProfiles
            NewRule.Action = NET_FW_ACTION_ALLOW

            'Add a new rule
            RulesObject.Add(NewRule)

            Return True



        Catch ex As Exception
            Program.Logger.Print(ex)
        End Try


        Return False

    End Function


    ''' <summary>
    ''' Check if this application is allowed for TCP Communication
    ''' </summary>
    ''' <param name="ApplicationName"></param>
    ''' <param name="ExecutorFullPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isApplicationTCPEnableExceptionAdded(ByVal ApplicationName As String,
                                                           ByVal ExecutorFullPath As String) As Boolean


        Dim osVer As OperatingSystem.MicrosoftOS = OperatingSystem.getOSType

        Select Case osVer

            Case Is = OperatingSystem.MicrosoftOS.WINDOWS_XP

                Return Registry.searchDataLikeInSubFolder(
                                                        Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                       "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\StandardProfile\AuthorizedApplications\List",
                                                        ApplicationName, ExecutorFullPath
                                                        )



                Return True

                'Case Is = MicrosoftOS.WINDOWS_7 
            Case Else
                REM Vista and the rest are like this


                Return Registry.searchDataLikeInSubFolder(
                                                         Registry.RegEditKeys.HKEY_LOCAL_MACHINE,
                                                        "SYSTEM\ControlSet001\services\SharedAccess\Parameters\FirewallPolicy\FirewallRules",
                                                        ApplicationName, ExecutorFullPath
                                                         )



        End Select



        Return False


    End Function



End Class
