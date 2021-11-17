Imports System.Threading
Imports System.Printing


Imports EPRO.Library.v3._5.Modules
Imports EPRO.Library.v3._5.Objects



Namespace MicrosoftOS



    Public Class PrintingMonitoring
        Implements IDisposable

        '
        '   Some KEY NOTES
        '   WHEN You are using application that consume much threads, Use handlers INSTEAD of Events
        '   Event General act async and might not get called
        '


#Region "Constructors"

        Public Sub New(pInvokeEventsOnNewThread As Boolean)
            Me._________MonitoringPool = New Dictionary(Of String, PrintJobDescriptor)()
            Me._________MonitoringDevices = New List(Of String)()


            Me.InvokeEventsOnNewThread = pInvokeEventsOnNewThread


            Me._________thrPrintMonitor = New Thread(AddressOf Me.runThrPrintMonitor) With {
                .Name = "_________thrPrintMonitor",
                .IsBackground = True
            }
            Me._________thrPrintMonitor.SetApartmentState(ApartmentState.MTA)
            Me._________thrPrintMonitor.Start()


            _________LockAddAndRemovePrinters = New Object()
        End Sub
        Public Sub New()

            Me.New(True)
        End Sub

#End Region


#Region "Methods"


        Private Sub onDocumentPrinted(pJob As PrintJobDescriptor)
            If Me.DocumentPrinted IsNot Nothing Then _
            DocumentPrinted(pJob)

        End Sub


        Private Sub onDocumentPrintingTerminated(pJob As PrintJobDescriptor)
            If Me.DocumentPrintingTerminated IsNot Nothing Then _
                DocumentPrintingTerminated(pJob)
        End Sub



        Private Sub runThrPrintMonitor()
            Try
                While Not Me.IsDisposed

                    '
                    '   Check Current Printer Classes in Collection and Process Them if Need Be
                    '
                    If Me._________MonitoringPool.Count > 0 Then
                        Dim pTempClassPointer As Dictionary(Of String, PrintJobDescriptor) =
                            New Dictionary(Of String, PrintJobDescriptor)()

                        For Each p In Me._________MonitoringPool.Values

                            If p.IsOverDueForUpdate() Then
                                '
                                '   Raise Event Terminated
                                '
                                If Me.InvokeEventsOnNewThread Then

                                    Dim thrTemp As New Thread(Sub() Me.onDocumentPrintingTerminated(p)) With {
                                        .IsBackground = True
                                    }
                                    thrTemp.Start()
                                Else
                                    If Me.DocumentPrintingTerminated IsNot Nothing Then DocumentPrintingTerminated(p)
                                End If
                                Thread.Sleep(2000)

                            ElseIf p.IsJobCompleted() Then
                                '
                                '   Raise Event Done
                                '
                                If Me.InvokeEventsOnNewThread Then
                                    Dim thrTemp As New Thread(Sub() Me.onDocumentPrinted(p)) With {
                                        .IsBackground = True
                                    }
                                    thrTemp.Start()
                                Else
                                    If Me.DocumentPrinted IsNot Nothing Then DocumentPrinted(p)
                                End If
                                Thread.Sleep(2000)
                            Else

                                pTempClassPointer.Add(p.JobUniqueID, p)
                            End If

                        Next

                        Me._________MonitoringPool.Clear()
                        Me._________MonitoringPool = pTempClassPointer

                    End If







                    '
                    '   Process Each Printer added
                    '
                    Dim pMonitoringDeviceCopy As String() = Nothing
                    SyncLock Me._________LockAddAndRemovePrinters
                        If Me._________MonitoringDevices.Count > 0 Then
                            pMonitoringDeviceCopy = System.Array.CreateInstance(GetType(String), Me._________MonitoringDevices.Count).Cast(Of String)().ToArray()
                            Me._________MonitoringDevices.CopyTo(pMonitoringDeviceCopy)
                        Else
                            pMonitoringDeviceCopy = New String() {}
                        End If
                    End SyncLock



                    Try


                        For Each pPrinterName In pMonitoringDeviceCopy

                            Dim pPrinterQueue = Me.getPrinterQueue(pPrinterName)
                            Dim pJobCollection As PrintJobInfoCollection = Nothing

                            If pPrinterQueue IsNot Nothing Then
                                pPrinterQueue.Refresh()
                                pJobCollection = pPrinterQueue.GetPrintJobInfoCollection()
                            End If

                            If pJobCollection IsNot Nothing AndAlso pJobCollection.Count() > 0 Then

                                For Each pJob In pJobCollection
                                    Dim pDescriptor As PrintJobDescriptor = Nothing
                                    Dim pUniqueID = PrintJobDescriptor.ConstructUniqueID(pPrinterName, pJob.JobIdentifier)
                                    If Me._________MonitoringPool.ContainsKey(pUniqueID) Then _
                                 pDescriptor = Me._________MonitoringPool(pUniqueID)

                                    If pDescriptor Is Nothing Then
                                        ' Dont bother adding again if this is repetitive IsPrinting or IsPrinted
                                        If Not pJob.IsPrinting AndAlso Not pJob.IsPrinted Then
                                            pDescriptor = New PrintJobDescriptor(pJob)
                                            Me._________MonitoringPool.Add(pUniqueID, pDescriptor)
                                        End If
                                    End If

                                    If pDescriptor IsNot Nothing Then pDescriptor.Update(pJob)


                                Next



                            End If


                        Next

                    Catch ex As Exception

                        Program.Logger.Print(ex)

                    End Try



















                    Thread.Sleep(THREAD______SLEEP_____INTERVAL__MILLI_SECS)
                End While
            Catch ex As ThreadAbortException

            Catch ex As Exception
                Program.Logger.Print(ex)

            Finally
                ' Log
                Program.Logger.Print(Thread.CurrentThread.Name & " is quiting ...")

            End Try
        End Sub




        Public Sub AddPrinter(pPrinterName As String)
            SyncLock Me._________LockAddAndRemovePrinters
                '
                '
                If Not Me._________MonitoringDevices.Contains(pPrinterName, New IEqualityComparerIgnoreCase()) Then _
                    Me._________MonitoringDevices.Add(pPrinterName)
            End SyncLock

        End Sub



        Public Sub RemovePrinter(pPrinterName As String)
            SyncLock Me._________LockAddAndRemovePrinters

                '
                '
                If Me._________MonitoringDevices.Contains(pPrinterName, New IEqualityComparerIgnoreCase()) Then _
                    Me._________MonitoringDevices.Remove(pPrinterName)

            End SyncLock


        End Sub

        Public Sub ClearAllPrinters()
            SyncLock Me._________LockAddAndRemovePrinters

                '
                '
                Me._________MonitoringDevices.Clear()

            End SyncLock


        End Sub




        Private Function getPrinterQueue(pPrinterName As String) As PrintQueue
            Try


                Dim localPrintServer2 = New LocalPrintServer()
                Dim pQueues = localPrintServer2.GetPrintQueues()

                Dim pQueue As PrintQueue = Nothing


                If pQueues IsNot Nothing Then
                    pQueue = pQueues.Where(Function(x) x.Name.Equals(pPrinterName)).FirstOrDefault()
                End If

                Return pQueue
            Catch ex As Exception
                Program.Logger.Print(ex)
                Return Nothing
            End Try

        End Function


#End Region



#Region "Enums and Consts"

        Const THREAD______SLEEP_____INTERVAL__MILLI_SECS As Int32 = 50

#End Region



#Region "Events"

        Public Delegate Sub dlgDocumentPrinted(pJob As PrintJobDescriptor)
        Public Delegate Sub dlgDocumentPrintingTerminated(pJob As PrintJobDescriptor)

#End Region



#Region "Properties"

        Public DocumentPrinted As dlgDocumentPrinted
        Public DocumentPrintingTerminated As dlgDocumentPrintingTerminated


        Private _________MonitoringDevices As List(Of String)
        Private _________MonitoringPool As Dictionary(Of String, PrintJobDescriptor)


        Private _________thrPrintMonitor As Thread


        Public ReadOnly Property Printers As IEnumerable(Of String)
            Get
                Return Me._________MonitoringDevices
            End Get
        End Property


        Private _________LockAddAndRemovePrinters As Object


        Public Property InvokeEventsOnNewThread As Boolean

#End Region






#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls
        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return disposedValue
            End Get
        End Property
        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If Me._________thrPrintMonitor IsNot Nothing AndAlso Me._________thrPrintMonitor.IsAlive Then _
                        Me._________thrPrintMonitor.Abort()
                    Me._________thrPrintMonitor = Nothing
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

End Namespace