Imports System.Management
Imports System.IO
Imports System.Net.NetworkInformation
Imports System.Net


Imports EPRO.WaitAsyncMgr


Imports EPRO.Library.v3._5.Objects.EStrings
Imports EPRO.Library.v3._5.Modules
Imports EPRO.Library.v3._5.Objects


Public Class Network
    ''Public Sub GetMAc(ByVal hostname As String)
    ''    Dim theManagementScope As New ManagementScope(String.Format("\\{0}\root\cimv2", hostname))
    ''    Const theQueryString As String = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 1"

    ''    Dim theObjectQuery As New ObjectQuery(theQueryString)

    ''    Dim theSearcher As New ManagementObjectSearcher(theManagementScope, theObjectQuery)
    ''    Dim theResultsCollection As ManagementObjectCollection = theSearcher.Get()

    ''    For Each currentResult As ManagementObject In theResultsCollection
    ''        MessageBox.Show(currentResult("MacAddress").ToString())
    ''    Next
    ''End Sub


    '' ''Public Sub DoGetHostAddresses(ByVal hostName As [String])

    '' ''    Dim ips As IPAddress()
    '' ''    'From the AddressList ,... The only one that has a valid address is the real connection
    '' ''    Try

    '' ''        Dim ipN As IPHostEntry = Dns.GetHostByName(hostName)
    '' ''        ''Dim iiP As IPHostEntry = Dns.GetHostEntry(hostName)
    '' ''        Dim iiP As IPHostEntry = Dns.Resolve(hostName)
    '' ''        Debug.Print(iiP.HostName)
    '' ''    Catch ex As Exception

    '' ''    End Try
    '' ''    ips = Dns.GetHostAddresses(hostName)

    '' ''    Debug.Print("GetHostAddresses(" + hostName + ") returns: ")

    '' ''    Dim index As Integer
    '' ''    For index = 0 To ips.Length - 1
    '' ''        Debug.Print(ips(index).ToString)
    '' ''    Next index
    '' ''End Sub


    '' ''' <summary>
    '' ''' Returns this system IP Address
    '' ''' </summary>
    '' ''' <remarks></remarks>
    ''Public Shared Function getMyIpAddress() As String

    ''    Try
    ''        If My.Computer.Network.IsAvailable Then

    ''            ' Dim Ips As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(My.Computer.Name)

    ''            Dim Ips() As System.Net.IPAddress = System.Net.Dns.GetHostAddresses("")

    ''            'Return Ips.AddressList(0).ToString
    ''            For Each IP As IPAddress In Ips

    ''                If clsNetwork.parseIP(IP) Then Return IP.ToString

    ''            Next

    ''        End If


    ''        Return "0.0.0.0"

    ''    Catch ex As Exception
    ''        Return "0.0.0.0"
    ''    End Try
    ''End Function

   




    ''' <summary>
    ''' Indicate if this computer can connect to internet
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CanConnectToInternet() As Boolean
        REM Try atleast 3times
        Try
            Dim iCount As Int16 = 3
            Do Until iCount < 1

                If My.Computer.Network.Ping("google.com") Then Return True

                iCount = CShort(iCount - 1)
                Threading.Thread.Sleep(500)
            Loop

        Catch ex As Exception

        End Try
        Return False
    End Function


    ''' <summary>
    ''' Get the http document of a url. the html codes of the page
    ''' </summary>
    ''' <param name="url"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetPageSource(ByVal url As String) As String
        Dim strReply As String = ""

        Try
            Dim objHttpRequest As HttpWebRequest = CType(HttpWebRequest.Create(url), HttpWebRequest)

            REM Just for testing
            'objHttpRequest.Method = "POST"


            Dim objHttpResponse As HttpWebResponse = CType(objHttpRequest.GetResponse(), HttpWebResponse)
            Dim objStrmReader As New StreamReader(objHttpResponse.GetResponseStream())

            strReply = objStrmReader.ReadToEnd()

        Catch ex As Exception
            strReply = ""
        End Try

        Return strReply
    End Function


    ''' <summary>
    ''' Check for software updating using the url provided. The version is returned in html <version> </version>
    ''' </summary>
    ''' <param name="CurrentVersion">The version Documented online. Returned ByRef</param>
    ''' <param name="InstalledApplicationVersion">The Version installed on this pc</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckForUpdate(ByVal UpdateSoftwareUrl As String,
                                          ByVal InstalledApplicationVersion As String,
                                          Optional ByRef CurrentVersion As String =
                                         Constants.vbNullString) As Boolean

        Dim aWait As New WaitAsync(, "Checking for update ...")
        'Current Version is the oNline Version
        CurrentVersion = GetPageSource(UpdateSoftwareUrl)

        aWait.Dispose()


        If CurrentVersion <> "" Then
            CurrentVersion = ExtractStringFromHtml(CurrentVersion, "<version>", "</version>")
            If CompareVersions(CurrentVersion,
                                                                    InstalledApplicationVersion) > 0 Then
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' Check for software updating using the url provided. The version is returned in html <version> </version>
    ''' This copy is not perfect yet
    ''' </summary>
    ''' <param name="CurrentVersion">The version Documented online. Returned ByRef</param>
    ''' <param name="InstalledApplicationVersion">The Version installed on this pc</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckForUpdate(ByVal UpdateSoftwareUrl As String,
                                          ByVal InstalledApplicationVersion As String,
                                          ByRef MyWebHttpRequest As EHttpRequest,
                                          Optional ByRef CurrentVersion As String =
                                         Constants.vbNullString) As Boolean

        Dim aWait As New WaitAsync(, "Checking for update ...")
        'Current Version is the oNline Version
        'CurrentVersion = GetPageSource(UpdateSoftwareUrl)
        'Dim WebClientFormattedURL As clsWebClient.WebClientURLFormat = New clsWebClient.WebClientURLFormat(UpdateSoftwareUrl)
        CurrentVersion = MyWebHttpRequest.UploadValues(UpdateSoftwareUrl)


        aWait.Dispose()


        If CurrentVersion <> "" Then
            CurrentVersion = ExtractStringFromHtml(CurrentVersion, "<version>", "</version>")
            If CompareVersions(CurrentVersion,
                                                                    InstalledApplicationVersion) > 0 Then
                Return True
            End If
        End If

        Return False
    End Function

    ''' <summary>
    ''' Check for software updating using the url provided. The version is returned in html <version> </version>
    ''' Also Url with <url></url>
    ''' </summary>
    ''' <param name="CurrentVersion">The version Documented online. Returned ByRef</param>
    ''' <param name="InstalledApplicationVersion">The Version installed on this pc</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckForUpdate(ByVal UpdateSoftwareUrl As String,
                                          ByVal InstalledApplicationVersion As String,
                                          ByRef CurrentVersion As String,
                                          ByRef CurrentVersionURL As String,
                                          Optional ByVal pSilent As Boolean = False) As Boolean

        REM This line slows down operation so make it display wait

        Dim aWait As WaitAsync = Nothing
        If Not pSilent Then [aWait] = New WaitAsync(, "Checking for update ...")

        'Current Version is the oNline Version
        Dim sHTML As String = GetPageSource(UpdateSoftwareUrl)

        If [aWait] IsNot Nothing Then aWait.Dispose()


        If sHTML <> "" Then

            CurrentVersion = ExtractStringFromHtml(sHTML, "<version>", "</version>")
            CurrentVersionURL = ExtractStringFromHtml(sHTML, "<url>", "</url>")

            If CompareVersions(CurrentVersion,
                                                                    InstalledApplicationVersion) > 0 Then
                Return True
            End If
        End If

        Return False
    End Function



    '   There is nothing yet I can do about this function .. It hangs the system on first call to fetch devices on network
    ''' <summary>
    ''' Process of Fetching Systems on LAN could be disturbing so allow enough time to list ...
    ''' Once first list is completed and count is more than one in any of the lsv ... increase interval to
    ''' 10 secs but this assignment can be edited by users in options
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetSystemsOnLAN() As List(Of String)
        Dim lsv As List(Of String) = New List(Of String)
        Dim childEntry As System.DirectoryServices.DirectoryEntry
        Dim ParentEntry As New System.DirectoryServices.DirectoryEntry
        Try
            ParentEntry.Path = "WinNT:"
            For Each childEntry In ParentEntry.Children
                Select Case childEntry.SchemaClassName
                    Case "Domain"

                        Dim SubChildEntry As System.DirectoryServices.DirectoryEntry
                        Dim SubParentEntry As New System.DirectoryServices.DirectoryEntry
                        SubParentEntry.Path = "WinNT://" & childEntry.Name
                        For Each SubChildEntry In SubParentEntry.Children
                            Select Case SubChildEntry.SchemaClassName
                                Case "Computer"
                                    'Debug.Print(SubChildEntry.Name)
                                    lsv.Add(SubChildEntry.Name.ToString)

                            End Select
                        Next
                End Select

            Next
        Catch Excep As Exception
            'MsgBox("Error While Reading Directories : " + Excep.Message.ToString)
        Finally
            ParentEntry = Nothing
        End Try

        Return lsv
    End Function


#Region "Fetching MacAddress NICS"


    ''' <summary>
    ''' Get Default NIC Mac Address .. because by default it returns it in order of activity
    ''' </summary>
    ''' <returns>The Mac Address without any delimiter [12 Chars]</returns>
    ''' <remarks></remarks>
    Public Shared Function GetMacAddress() As String
        Dim nic As NetworkInterface = Nothing
        Dim macAddress As String = ""

        For Each nic In NetworkInterface.GetAllNetworkInterfaces()
            macAddress = nic.GetPhysicalAddress().ToString.Trim()

            If macAddress <> "" Then
                Return macAddress
            End If
        Next

        Return macAddress
    End Function

    Public Shared Function GetAllNetworkInterfaces() As List(Of NetworkInterface)
        Return NetworkInterface.GetAllNetworkInterfaces().ToList()
    End Function


    Public Shared Function GetNetworkInterfacesWithValidIPv4() As List(Of NetworkInterface)
        Dim pRst As New List(Of NetworkInterface)()

        For Each NIC As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()


            If NIC.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 OrElse
                NIC.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then


                For Each ip As UnicastIPAddressInformation In NIC.GetIPProperties().UnicastAddresses

                    If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then

                        pRst.Add(NIC)
                        Exit For

                    End If

                Next


            End If


        Next
        Return pRst
    End Function



#End Region




    ''' <summary>
    ''' Checks if the ips are on the same network. like 192.168.5 and 192.168.56 are not
    ''' </summary>
    ''' <param name="Ip1"></param>
    ''' <param name="Ip2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function isIPNetworkSame(ByVal Ip1 As String, ByVal Ip2 As String) As Boolean

        If Not parseIP(Ip1) Or Not parseIP(Ip2) Then Return False

        Dim IP1Split() As String = Split(Ip1, ".")
        Dim IP2Split() As String = Split(Ip2, ".")

        If String.Format("{0}.{1}.{2}", IP1Split) <> String.Format("{0}.{1}.{2}", IP2Split) Then Return False

        Return True

    End Function


    ''' <summary>
    ''' Get an Ip relating to another IP from many Ips
    ''' </summary>
    ''' <param name="Ips"></param>
    ''' <param name="relatingToIp"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getValidIP_Relating_To(ByVal Ips() As String, ByVal relatingToIp As String) As String

        'MyLogFile.Log("A Client is trying to Connect with these IPs: " & MessageReceived(3))

        Dim ValidIP As String = relatingToIp

        For IpCount As Int16 = 0 To CShort(Ips.Count - 1)


            If Network.isIPNetworkSame(Ips(IpCount), relatingToIp) Then
                ValidIP = Ips(IpCount)
                ' MyLogFile.Log("Found Valid IPs: " & ValidIP)
                Exit For
            End If


        Next

        Return ValidIP

    End Function





#Region "Fetching IPs"

#Region "Structures"

    ''' <summary>
    ''' It splits the structure into 4 parts. Throws Invalid IP Exception
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure IPv4BreakdownStructure


#Region "Enums and Constants"

        ''        Class A	1.0.0.1 to 126.255.255.254	Supports 16 million hosts on each of 127 networks.
        ''Class B	128.1.0.1 to 191.255.255.254	Supports 65,000 hosts on each of 16,000 networks.
        ''Class C	192.0.1.1 to 223.255.254.254	Supports 254 hosts on each of 2 million networks.
        ''Class D	224.0.0.0 to 239.255.255.255	Reserved for multicast groups.
        ''Class E	240.0.0.0 to 254.255.255.254	Reserved for future use, or Research and Development Purposes.


        Public Enum IPClasses
            CLASS_A
            CLASS_B
            CLASS_C
            CLASS_D
            CLASS_E
            UNKNOWN
        End Enum


#End Region


#Region "Properties"



        Private _______IPv4Part1 As String
        Private _______IPv4Part2 As String
        Private _______IPv4Part3 As String
        Private _______IPv4Part4 As String




        Public ReadOnly Property IPv4Part1 As String
            Get
                Return Me._______IPv4Part1
            End Get
        End Property


        Public ReadOnly Property IPv4Part2 As String
            Get
                Return Me._______IPv4Part2
            End Get
        End Property


        Public ReadOnly Property IPv4Part3 As String
            Get
                Return Me._______IPv4Part3
            End Get
        End Property


        Public ReadOnly Property IPv4Part4 As String
            Get
                Return Me._______IPv4Part4
            End Get
        End Property


        Public ReadOnly Property CDRNumber As Byte
            Get
                Return CByte(Me.ToBinaryNotation().Replace("0", "").Where(Function(x) x = "1").Count())
            End Get
        End Property


        '        A	1 – 126*	0	N.H.H.H	255.0.0.0	126 (27 – 2)	16,777,214 (224 – 2)
        'B	128 – 191	10	N.N.H.H	255.255.0.0	16,382 (214 – 2)	65,534 (216 – 2)
        'C	192 – 223	110	N.N.N.H	255.255.255.0	2,097,150 (221 – 2)	254 (28 – 2)
        'D	224 – 239	1110	Reserved for Multicasting
        'E	240 – 254
        Public ReadOnly Property IPClass As IPClasses
            Get
                Select Case Me.IPv4Part1.toInt32()
                    Case 1 To 126
                        Return IPClasses.CLASS_A

                    Case 128 To 191
                        Return IPClasses.CLASS_B

                    Case 192 To 223

                        Return IPClasses.CLASS_C

                    Case 224 To 239
                        Return IPClasses.CLASS_D

                    Case 240 To 254
                        Return IPClasses.CLASS_E

                End Select

                Return IPClasses.UNKNOWN
            End Get
        End Property


#End Region



#Region "Constructors"



        ''' <summary>
        ''' Use for parsing in other base like binary and hex
        ''' </summary>
        ''' <param name="pPart1"></param>
        ''' <param name="pPart2"></param>
        ''' <param name="pPart3"></param>
        ''' <param name="pPart4"></param>
        ''' <remarks></remarks>
        Friend Sub New(pPart1 As String,
                 pPart2 As String,
                 pPart3 As String,
                 pPart4 As String)
            Me._______IPv4Part1 = pPart1
            Me._______IPv4Part2 = pPart2
            Me._______IPv4Part3 = pPart3
            Me._______IPv4Part4 = pPart4
        End Sub

        ''' <summary>
        ''' Use internally only
        ''' </summary>
        ''' <param name="pIPv4Address">Decimal Notation</param>
        ''' <remarks></remarks>
        Friend Sub New(pIPv4Address As String,
                       pIgnoreChecking As Boolean)
            If Not pIgnoreChecking AndAlso Not parseIP(pIPv4Address) Then Throw New Exception("Please, pass in a valid IPv4 Address with (.) delimited notation")
            Dim IPSegs() As String = Split(pIPv4Address, ".")
            Me._______IPv4Part1 = IPSegs(0)
            Me._______IPv4Part2 = IPSegs(1)
            Me._______IPv4Part3 = IPSegs(2)
            Me._______IPv4Part4 = IPSegs(3)
        End Sub
        Public Sub New(pPart1 As Byte,
                       pPart2 As Byte,
                       pPart3 As Byte,
                       pPart4 As Byte)
            Me.New(String.Format(
            "{0}.{1}.{2}.{3}", pPart1,
                                pPart2, pPart3, pPart4
            ))
        End Sub

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pIPv4Address">Decimal Notation</param>
        ''' <remarks></remarks>
        Public Sub New(pIPv4Address As String)
            Me.New(pIPv4Address, False)

        End Sub



#End Region






        Public Overrides Function ToString() As String
            Return String.Format(
            "{0}.{1}.{2}.{3}", Me._______IPv4Part1,
            Me._______IPv4Part2, Me._______IPv4Part3,
            Me._______IPv4Part4
            )
        End Function

        Public Overloads Function ToString(ByVal pWithDelimiter As Boolean) As String
            If pWithDelimiter Then Return Me.ToString()
            Return Me.ToString().Replace(".", "")
        End Function




        ''' <summary>
        ''' Be sure this class is decimal notation else it will throw exception
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToBinaryNotation() As String
            Return String.Format(
            "{0}.{1}.{2}.{3}",
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(Me._______IPv4Part1), 2
                    ),
                8),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part2), 2
                    ),
                8),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part3), 2
                    ),
                8),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part4), 2
                    ),
                8)
            )

        End Function


        Public Function ToPingableBinaryNotation() As Long
            Return Convert.ToInt64(Me.ToBinaryNotation.Replace(".", ""), 2)
        End Function



        ''' <summary>
        ''' Be sure this class is decimal notation else it will throw exception.
        ''' To ping just remove the (.) and do like this ping 0xc0a83865
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ToHexNotation() As String
            Return String.Format(
            "{0}.{1}.{2}.{3}",
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(Me._______IPv4Part1), 16
                    ),
                2),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part2), 16
                    ),
                2),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part3), 16
                    ),
                2),
            WrapUp(
                Convert.ToString(
                    EInt.valueOf(_______IPv4Part4), 16
                    ),
                2)
            )

        End Function

        Public Function ToPingableHexNotation() As String
            Return "0x" & Me.ToHexNotation().Replace(".", "")
        End Function









        Public Shared Function FromBinary(pBinaryNotation As String) As IPv4BreakdownStructure
            Dim p = New IPv4BreakdownStructure(pBinaryNotation, True)
            Return New IPv4BreakdownStructure(
                Convert.toInt32(p.IPv4Part1, 2).ToString(),
                Convert.toInt32(p.IPv4Part2, 2).ToString(),
                Convert.toInt32(p.IPv4Part3, 2).ToString(),
                Convert.toInt32(p.IPv4Part4, 2).ToString()
                )
        End Function

        Public Shared Function CreateDelimitedBinaryNotation(pBinaryString As String) As String
            If pBinaryString.Length <> 32 Then Throw New Exception("It must be 32 in length")
            Dim p = EStrings.SplitChunk(pBinaryString, 8).ToArray()
            Return New IPv4BreakdownStructure(p(0), p(1), p(2), p(3)).ToString()
        End Function



    End Structure


    ''' <summary>
    ''' Displays more details of IPv4 including Binary Notations
    ''' </summary>
    ''' <remarks></remarks>
    Public Structure IPv4AdressDetails
        Private _______IPv4DecimalNotation As IPv4BreakdownStructure

        Private _______MaskDecimalNotation As IPv4BreakdownStructure

        Private _______BroadcastAddressDecimalNotation As IPv4BreakdownStructure

        Private _______MaskWildCard As IPv4BreakdownStructure

        Private _______IPv4NetworkDecimalNotation As IPv4BreakdownStructure

        Private _______HostMinimumIPv4DecimalNotation As IPv4BreakdownStructure
        Private _______HostMaximumIPv4DecimalNotation As IPv4BreakdownStructure

        Private _______MaximumPosibleHosts As Long


        Public ReadOnly Property IPv4DecimalNotation As IPv4BreakdownStructure
            Get
                Return _______IPv4DecimalNotation
            End Get
        End Property


        Public ReadOnly Property MaskDecimalNotation As IPv4BreakdownStructure
            Get
                Return _______MaskDecimalNotation
            End Get
        End Property

        Public ReadOnly Property BroadcastAddressDecimalNotation As IPv4BreakdownStructure
            Get
                Return _______BroadcastAddressDecimalNotation
            End Get
        End Property

        Public ReadOnly Property MaskWildCard As IPv4BreakdownStructure
            Get
                Return _______MaskWildCard
            End Get
        End Property

        Public ReadOnly Property IPv4NetworkDecimalNotation As IPv4BreakdownStructure
            Get
                Return _______IPv4NetworkDecimalNotation
            End Get
        End Property

        Public ReadOnly Property HostMinimumIPv4DecimalNotation As IPv4BreakdownStructure
            Get
                Return _______HostMinimumIPv4DecimalNotation
            End Get
        End Property

        Public ReadOnly Property HostMaximumIPv4DecimalNotation As IPv4BreakdownStructure
            Get
                Return _______HostMaximumIPv4DecimalNotation
            End Get
        End Property


        Public ReadOnly Property MaximumPosibleHosts As Long
            Get
                Return _______MaximumPosibleHosts
            End Get
        End Property

        Public ReadOnly Property IPv4CDRNotation As String
            Get
                Return String.Format("{0}/{1}", Me._______IPv4DecimalNotation.ToString(),
                                     Me._______MaskDecimalNotation.CDRNumber
                                     )
            End Get
        End Property




















        ''' <summary>
        ''' Displays more details of IPv4 including Binary Notations. Throws Invalid IP Exception
        ''' </summary>
        ''' <param name="pIPv4DecimalNotation"></param>
        ''' <param name="pIPv4MaskDecimalNotation"></param>
        ''' <remarks></remarks>
        Public Sub New(pIPv4DecimalNotation As String,
                       pIPv4MaskDecimalNotation As String)

            Me._______IPv4DecimalNotation = New IPv4BreakdownStructure(pIPv4DecimalNotation)
            Me._______MaskDecimalNotation = New IPv4BreakdownStructure(pIPv4MaskDecimalNotation)


            '
            '   Network IP
            '
            Dim pIPBinary = Me._______IPv4DecimalNotation.ToBinaryNotation().Replace(".", "")
            Dim pMaskBinary = Me._______MaskDecimalNotation.ToBinaryNotation().Replace(".", "")
            Dim pNetWorkBinary = Strings.Left(pIPBinary, Me._______MaskDecimalNotation.CDRNumber)

            If pNetWorkBinary.Length > 32 Then Throw New Exception("INVALID Network IP Configuration Found!")
            If pNetWorkBinary.Length < 32 Then pNetWorkBinary &= StrDup(32 - pNetWorkBinary.Length, "0"c)

            Me._______IPv4NetworkDecimalNotation = IPv4BreakdownStructure.FromBinary(
                                                    IPv4BreakdownStructure.CreateDelimitedBinaryNotation(pNetWorkBinary)
                                                    )

            '
            '   Minimum Host IP
            '
            Dim pMin = Strings.Left(pIPBinary, Me._______MaskDecimalNotation.CDRNumber)
            If pMin.Length < 31 Then pMin &= StrDup(31 - pMin.Length, "0"c)
            If pMin.Length < 32 Then pMin &= "1"
            Me._______HostMinimumIPv4DecimalNotation = IPv4BreakdownStructure.FromBinary(
                                                   IPv4BreakdownStructure.CreateDelimitedBinaryNotation(pMin)
                                                   )

            '
            '   Maximum Host IP
            '
            pMin = Strings.Left(pIPBinary, Me._______MaskDecimalNotation.CDRNumber)
            If pMin.Length < 31 Then pMin &= StrDup(31 - pMin.Length, "1"c)
            If pMin.Length < 32 Then pMin &= "0"
            Me._______HostMaximumIPv4DecimalNotation = IPv4BreakdownStructure.FromBinary(
                                                   IPv4BreakdownStructure.CreateDelimitedBinaryNotation(pMin)
                                                   )


            '
            '   Broadcast IP 
            '
            Dim pTemp = Me._______HostMaximumIPv4DecimalNotation
            Me._______BroadcastAddressDecimalNotation = New IPv4BreakdownStructure(
                pTemp.IPv4Part1, pTemp.IPv4Part2, pTemp.IPv4Part3, (pTemp.IPv4Part4.toInt32() + 1).ToString()
                )



            '
            '   Mask Wild Card
            '

            pMin = StrDup(Me._______MaskDecimalNotation.CDRNumber, "0"c)
            If pMin.Length < 32 Then pMin &= StrDup(32 - Me._______MaskDecimalNotation.CDRNumber, "1"c)
            Me._______MaskWildCard = IPv4BreakdownStructure.FromBinary(
                                                   IPv4BreakdownStructure.CreateDelimitedBinaryNotation(pMin)
                                                   )




            '
            ' Maximum Possible Host
            '

            Dim p = Math.Pow(2, 32 - Me._______MaskDecimalNotation.CDRNumber) - 2
            Me._______MaximumPosibleHosts = 0
            If p > 0 Then Me._______MaximumPosibleHosts = p.toLong()







        End Sub





    End Structure

#End Region





    ''' <summary>
    ''' Returns all valid Ip connection on this Computer. Separated with Comma(,)
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getAllIpsOnMyComputer() As String

        Try
            Dim Result As String = vbNullString

            If My.Computer.Network.IsAvailable Then

                ' Dim Ips As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(My.Computer.Name)

                Dim Ips() As System.Net.IPAddress = System.Net.Dns.GetHostAddresses("")

                'Return Ips.AddressList(0).ToString
                For Each IP As IPAddress In Ips

                    If Network.parseIP(IP) Then Result &= IP.ToString & ","

                Next

                If Result <> vbNullString Then Result = Left(Result, Len(Result) - 1)


            End If


            Return Result

        Catch ex As Exception
            Return "0.0.0.0"
        End Try



    End Function


    Public Shared Function getAllIpDetailsThisPC() As List(Of IPv4AdressDetails)
        Dim pRst As New List(Of IPv4AdressDetails)()
        Try
            If My.Computer.Network.IsAvailable Then

                For Each NIC As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                    Dim p = getAllIpDetailsForNIC(NIC)
                    If p.Count > 0 Then pRst.AddRange(p)

                Next

            End If


        Catch ex As Exception
            Program.Logger.Print(ex)
        End Try
        Return pRst
    End Function


    Public Shared Function getAllIpDetailsForNIC(NIC As NetworkInterface) As List(Of IPv4AdressDetails)
        Dim pRst As New List(Of IPv4AdressDetails)()
        Try

            If NIC.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 OrElse
                NIC.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then


                For Each ip As UnicastIPAddressInformation In NIC.GetIPProperties().UnicastAddresses

                    If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then

                        pRst.Add(New IPv4AdressDetails(ip.Address.ToString(),
                                                       ip.IPv4Mask.ToString()
                                                       ))

                    End If

                Next


            End If

        Catch ex As Exception
            Program.Logger.Print(ex)
        End Try
        Return pRst
    End Function


    ''' <summary>
    ''' Check Whether the IP Entered is in a correct format [IP v4]. 
    ''' Min: 0.0.0.0
    ''' Max: 255.255.255.255
    ''' </summary>
    ''' <param name="IP">Ip Address to check</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function parseIP(ByVal IP As IPAddress) As Boolean
        REM 3 Locations of (.)
        REM [.] must not start and must not end
        REM Must Contain 4 Fragments
        REM Each Fragment Must not exceed a byte 255
        REM Run thru each fragment, Each must not exceed Max: 3chars and Min: 1
        REM All Fragment must be numeric


        Return Network.parseIP(IP.ToString())

    End Function




    ''' <summary>
    ''' Check Whether the IP Entered is in a correct format [IP v4]. 
    ''' Min: 0.0.0.0
    ''' Max: 255.255.255.255
    ''' </summary>
    ''' <param name="IP">Ip Address to check</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function parseIP(ByVal IP As String) As Boolean
        REM 3 Locations of (.)
        REM [.] must not start and must not end
        REM Must not contain [..]
        REM Must Contain 4 Fragments
        REM Each Fragment Must not exceed a byte 255
        REM Run thru each fragment, Each must not exceed Max: 3chars and Min: 1
        REM All Fragment must be numeric


        If IP.Length = 0 Then Return False
        If IP.Substring(0, 1).Equals(".") Then Return False
        If IP.Substring(IP.Length - 1, 1).Equals(".") Then Return False
        If IP.IndexOf("..") > 0 Then Return False

        Dim IPSegs() As String = Split(IP, ".")
        If IPSegs.Length <> 4 Then Return False


        REM Let us check the content of each segment

        For Each Segment As String In IPSegs
            If Segment Is Nothing Then Return False
            REM Trim incase of space
            If Segment.Length > 3 Or Segment.Trim.Length < 1 Then Return False

            REM Assume no space and still meet the length
            If Not IsNumeric(Segment) Then Return False

            REM Since it is numeric, check if it is a byte value
            If Val(Segment) > 255 Then Return False

        Next




        REM If all segments passed the test then return true
        Return True
    End Function



    ''' <summary>
    ''' Returns this system IP Address
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Function getMyIpAddress() As String

        Try
            If My.Computer.Network.IsAvailable Then

                For Each NIC As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces


                    If NIC.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 OrElse
                        NIC.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then


                        For Each ip As UnicastIPAddressInformation In NIC.GetIPProperties().UnicastAddresses

                            If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then


                                Return ip.Address.ToString()


                            End If

                        Next


                    End If




                Next

            End If


            Return "0.0.0.0"

        Catch ex As Exception
            Return "0.0.0.0"
        End Try
    End Function



    ''' <summary>
    ''' Coverts IP Address to HostName
    ''' </summary>
    ''' <param name="IP"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetHostNameFromIP(ByRef IP As String) As String
        Try
            Dim host As System.Net.IPHostEntry
            REM It is obsolete but it works better than getHostEntry function
            host = System.Net.Dns.GetHostByAddress(IP)

            Return host.HostName
        Catch ex As Exception

            Return vbNullString

        End Try

    End Function


#End Region







End Class
