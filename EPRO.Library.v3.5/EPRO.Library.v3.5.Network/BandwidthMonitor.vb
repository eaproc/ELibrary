
Imports System.Net.NetworkInformation
Imports System.Threading


Imports EPRO.Library.v3._5
Imports EPRO.Library.v3._5.MicrosoftOS


''Anyways, It is not possible to be connected without a network interface

''Also it is possible to have the bandwidth monitor class created but not running

''if adapter is not connected
''if adapter does not support ipv4



''' <summary>
''' Monitors the BandwidthUsage of an Interface. Events are on separate threads
''' </summary>
''' <remarks></remarks>
Public Class BandwidthMonitor
    Implements IDisposable




#Region "Properties"




    ''' <summary>
    ''' The Network Interface we are monitoring
    ''' </summary>
    ''' <remarks></remarks>
    Private Property MonitoringInterface As NetworkInterface = Nothing


    ''' <summary>
    ''' This will be the Maximum Reference of the Bandwith Receive Usage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property PassedIn___TargetReceiveBandwithInKb As Long


    ''' <summary>
    ''' This will be the Maximum Reference of the Bandwith Receive Usage
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private ReadOnly Property PassedIn___TargetReceiveBandwithInBytes As Long
        Get
            Return Me.PassedIn___TargetReceiveBandwithInKb * 1024
        End Get
    End Property



    ''' <summary>
    ''' Thread Manager for this Class
    ''' </summary>
    ''' <remarks></remarks>
    Private _________thrFillUpStats As Thread
    Private _________thr__LowerThan_Infomer As Thread



    ''' <summary>
    ''' Indicate if this class is working
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isClassOnline As Boolean
        Get
            Return _________thrFillUpStats IsNot Nothing AndAlso _________thrFillUpStats.IsAlive
        End Get
    End Property


    ''' <summary>
    ''' Keep the Last Receive Byte Value to know when the interface is reset
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property ReceivedBytesStartingPoint As ULong = 0


    ''' <summary>
    ''' Save the Current Total Bytes Received
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property TotalBytesReceived As ULong = 0


    ''' <summary>
    ''' The accumulated total bytes even when the Interface has been restarted
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Property AccumulatedTotalReceivedBytes As ULong = 0



    ''' <summary>
    ''' Gets the Total Bytes Received on this interface since this class has been initialized
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTotalBytesReceived As ULong
        Get
            Return Me.AccumulatedTotalReceivedBytes
        End Get
    End Property


    ''' <summary>
    ''' Checks if the TargetReceived  in bytes size has been ment
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property isBandwidthReceivedUsedUp As Boolean
        Get
            Return (Me.AccumulatedTotalReceivedBytes >= (Me.PassedIn___TargetReceiveBandwithInBytes))
        End Get
    End Property


    ''' <summary>
    ''' Returns the Target Received Left in Bytes ... but returns 0 if targe is already met
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTargetRecievedLeftInBytes As Long
        Get
            ''Program.Logger.Log("Fetching Bandwidth:")
            ''Program.Logger.Log("AccumulatedTotalReceivedBytes:" & Me.AccumulatedTotalReceivedBytes)

            If Me.isBandwidthReceivedUsedUp Then
                Return 0
            Else
                Return CLng(Me.PassedIn___TargetReceiveBandwithInBytes - Me.AccumulatedTotalReceivedBytes)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the Target Received Left in Kb ... but returns 0 if targe is already met
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTargetRecievedLeftInKb As Long
        Get
            If Me.isBandwidthReceivedUsedUp Then
                Return 0
            Else
                Return CLng(Me.getTargetRecievedLeftInBytes / 1024)
            End If
        End Get
    End Property

    ''' <summary>
    ''' Returns the Target Received Left in Mb ... but returns 0 if targe is already met
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTargetRecievedLeftInMb As Long
        Get

            Return CLng(Me.getTargetRecievedLeftInMbSingleFormat)

        End Get
    End Property


    ''' <summary>
    ''' Returns the Target Received Left in Mb and in single 2 decimal places
    ''' ... but returns 0 if targe is already met
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property getTargetRecievedLeftInMbSingleFormat As Single
        Get
            If Me.isBandwidthReceivedUsedUp Then
                Return 0
            Else
                Return CSng(FormatNumber(CSng(Me.getTargetRecievedLeftInKb / 1024), 2, TriState.True))
            End If
        End Get
    End Property



    Private ______MonitorFetchIntervalMilliSecs As Int32 = 1500

    Private ______AlertRunningLowerThanInKb As Long


#End Region



#Region "Events"

    Public Event BandwidthIsRunningLow()

#End Region


#Region "Constructors"



    ''' <summary>
    ''' Create new Bandwidth Monitor. If there is a problem initializing this class due to NICs. It doesnt count
    ''' </summary>
    ''' <param name="NIC"></param>
    ''' <param name="pTargetReceiveBandwithInKb">The Target Recieve Bandwidth</param>
    ''' <remarks></remarks>
    Private Sub New(ByVal NIC As NetworkInterface,
            ByVal pTargetReceiveBandwithInKb As ULong)

        Me.PassedIn___TargetReceiveBandwithInKb = CLng(pTargetReceiveBandwithInKb)
        Me.MonitoringInterface = NIC


        If Me.MonitoringInterface IsNot Nothing Then

            REM Save initial byte receive
            If Me.MonitoringInterface.Supports(NetworkInterfaceComponent.IPv4) Then
                Me.ReceivedBytesStartingPoint = CULng(Me.MonitoringInterface.GetIPv4Statistics().BytesReceived)
                Me.TotalBytesReceived = Me.ReceivedBytesStartingPoint


                ''Program.Logger.Log("Starting Bandwidth:")
                ''Program.Logger.Log("ReceivedBytesStartingPoint:" & Me.ReceivedBytesStartingPoint)

            End If

            REM Place the everysec 1.5secs thread
            _________thrFillUpStats = New Thread(
                AddressOf SetNewStateValues
                ) With {.Name = "_________thrFillUpStats",
                        .IsBackground = True
                       }

            _________thrFillUpStats.SetApartmentState(ApartmentState.MTA)
            _________thrFillUpStats.Start()


        End If

    End Sub



    ''' <summary>
    ''' Use the Server IP Address to get which NIC is connected if many NIC is available. Else default NIC is used
    ''' </summary>
    ''' <param name="NICServerIPonSameNetwork"></param>
    ''' <param name="pTargetReceiveBandwithInKb"></param>
    ''' <remarks></remarks>
    Public Sub New(ByVal NICServerIPonSameNetwork As String,
            ByVal pTargetReceiveBandwithInKb As ULong,
            Optional ByVal pAlertLowerThanInKb As Long = 0
            )
        Me.New(FetchMatchingInterface(NICServerIPonSameNetwork), pTargetReceiveBandwithInKb)
        Me.______AlertRunningLowerThanInKb = pAlertLowerThanInKb
    End Sub



    ''' <summary>
    ''' Use default NIC 
    ''' </summary>
    ''' <param name="pTargetReceiveBandwithInKb"></param>
    ''' <remarks></remarks>
    Public Sub New(
            ByVal pTargetReceiveBandwithInKb As ULong)
        Me.New(FetchMatchingInterface(Network.getMyIpAddress()), pTargetReceiveBandwithInKb)
    End Sub



#End Region

#Region "Methods"

    Private Sub onAlertLowerThanInKb()
        _________thr__LowerThan_Infomer = New Thread(
           AddressOf Run__onAlertLowerThanInKb
           ) With {.Name = "_________thr__LowerThan_Infomer",
                   .IsBackground = True
                  }

        _________thr__LowerThan_Infomer.SetApartmentState(ApartmentState.MTA)
        _________thr__LowerThan_Infomer.Start()


        Me.______AlertRunningLowerThanInKb = 0

    End Sub
    Private Sub Run__onAlertLowerThanInKb()

        Try

            RaiseEvent BandwidthIsRunningLow()

        Catch ex As ThreadAbortException
        Catch ex As Exception

            Program.Logger.Print(ex)
        End Try

    End Sub




    ''' <summary>
    ''' Set new values to this class depending on the state of the Interface
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SetNewStateValues()

        Try
            While Not Me.disposedValue

                If Me.MonitoringInterface IsNot Nothing Then
                    If Me.MonitoringInterface.OperationalStatus = OperationalStatus.Up Then

                        If Me.MonitoringInterface.Supports(NetworkInterfaceComponent.IPv4) Then

                            Dim NIC_IPv4_Stats As IPv4InterfaceStatistics =
                                Me.MonitoringInterface.GetIPv4Statistics()

                            If Me.TotalBytesReceived > NIC_IPv4_Stats.BytesReceived Then
                                REM This interface has been restarted
                                REM Reset and dont get a reading
                                REM Am using Me.TotalBytesReceived Interchangeably with ReceivedByteStartPoint
                                Me.ReceivedBytesStartingPoint = CULng(NIC_IPv4_Stats.BytesReceived)
                                Me.TotalBytesReceived = Me.ReceivedBytesStartingPoint

                                ''Program.Logger.Log("Resetting Bandwidth:")
                                ''Program.Logger.Log("ReceivedBytesStartingPoint:" & Me.ReceivedBytesStartingPoint)


                            Else
                                REM We are on Course
                                REM Am only collecting the differences
                                Me.AccumulatedTotalReceivedBytes = CULng(
                                                    Me.AccumulatedTotalReceivedBytes + (NIC_IPv4_Stats.BytesReceived -
                                                    Me.TotalBytesReceived)
                                )
                                Me.TotalBytesReceived = CULng(NIC_IPv4_Stats.BytesReceived)

                                ''Program.Logger.Log("Setting Bandwidth:")
                                ''Program.Logger.Log("AccumulatedTotalReceivedBytes:" & Me.AccumulatedTotalReceivedBytes)
                                ''Program.Logger.Log("TotalBytesReceived:" & Me.TotalBytesReceived)



                            End If



                        End If

                    End If

                End If



                If Me.______AlertRunningLowerThanInKb > 0 AndAlso Me.getTargetRecievedLeftInKb() <= Me.______AlertRunningLowerThanInKb Then
                    Me.onAlertLowerThanInKb()

                End If


                Thread.Sleep(Me.______MonitorFetchIntervalMilliSecs)
            End While

        Catch ex As ThreadAbortException
        Catch ex As Exception

            Program.Logger.Print(ex)

        End Try


        Program.Logger.Print(Thread.CurrentThread.Name & " is quiting ...")
    End Sub


    ''' <summary>
    ''' Fetch the Network interface. If no match found it returns the first interface
    ''' </summary>
    ''' <param name="IPv4Address"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function FetchMatchingInterface(ByVal IPv4Address As String) As NetworkInterface

        Dim AllNiCs() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces

        If AllNiCs Is Nothing Then Return Nothing

        If AllNiCs.Count = 0 Then Return Nothing

        REM Keep first one
        Dim returnNIC As NetworkInterface = AllNiCs(0)

        For Each NIC As NetworkInterface In AllNiCs


            If NIC.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Or
                NIC.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then


                For Each ip As UnicastIPAddressInformation In NIC.GetIPProperties().UnicastAddresses

                    If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                        ' Debug.Print(NIC.Name)
                        'If clsNetwork.parseIP(ip.Address.ToString()) Then


                        'Debug.Print(ip.Address.ToString())

                        If Network.isIPNetworkSame(ip.Address.ToString(), IPv4Address) Then
                            returnNIC = NIC
                            Exit For
                        End If


                        'End If

                    End If

                Next


            End If




        Next


        Return returnNIC

    End Function


#End Region





#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not Me._________thrFillUpStats Is Nothing AndAlso _________thrFillUpStats.IsAlive Then _________thrFillUpStats.Abort()
                Me._________thrFillUpStats = Nothing

                Try
                    If Not Me._________thr__LowerThan_Infomer Is Nothing AndAlso _________thr__LowerThan_Infomer.IsAlive Then _________thr__LowerThan_Infomer.Abort()
                    _________thr__LowerThan_Infomer = Nothing
                Catch ex As Exception
                    Program.Logger.Print(ex)
                End Try
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
