Imports System.Threading
Imports System.Windows.Forms

Namespace SplashCoordinator

    '   Events runs on Co-ordinator Main thread
    '   passed in actions runs on a separate background thread


    ''' <summary>
    ''' Initiates Coordinator on Constructor. Show your splash first before calling coordinator. Events runs on Main thread
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ProcessCoordinator(Of Initiator As {IProgressEvent, IProcessCoordinator, Control})
        Implements IDisposable


#Region "Properties"

        Private tmrCoordinate As Thread
        Private ProcessesToRun As List(Of Action)
        REM         Private ExitAction As Action

        Private TotalProcessToRunCount As UInt16
        Private TotalProcessesRan As UInt16


        Private CountUpTo As UInt16
        Private CurrentCount As Int32

        Private NextIntervalMarkPoint As Int32
        Private NormalInterval As UInt16

        Private ProgressHandler As Initiator

        Public ReadOnly Property Initializer As Initiator
            Get
                Return Me.ProgressHandler
            End Get
        End Property

        REM Private waitingForSignal As Boolean

#End Region




#Region "Constructors"

        ''' <summary>
        ''' Organizes the way processes are run under splash screens. Events runs on Main thread
        ''' </summary>
        ''' <param name="__CountUpTo">Must be greater than total process times 5</param>
        ''' <param name="__ProcessesToRun">Processes you wish to execute under splash screens</param>
        ''' <remarks></remarks>
        Public Sub New(
                ByVal __CountUpTo As UInt16,
                 ByVal __ProgressHandler As Initiator,
                ParamArray __ProcessesToRun() As Action)

            Me.ProcessesToRun = New List(Of Action)()
            If __ProcessesToRun IsNot Nothing Then Me.ProcessesToRun.AddRange(__ProcessesToRun)
            REM If __ExitAction Is Nothing Then Throw New Exception("Exit Action Can not be Null")
            REM Me.ExitAction = __ExitAction
            Me.ProgressHandler = __ProgressHandler


            Me.CountUpTo = __CountUpTo
            Me.TotalProcessToRunCount = CUShort(Me.ProcessesToRun.Count)

            If Me.TotalProcessToRunCount > 0 Then
                If Me.CountUpTo < (Me.TotalProcessToRunCount * 5) Then Throw New Exception("Count Up To Parameter must be greater than  or equals to processes count * 5")

                Me.NormalInterval = CUShort(Me.CountUpTo / Me.TotalProcessToRunCount)
                Me.CountUpTo += Me.NormalInterval  REM Extend Counter

                Me.SetNextInterval()


            End If

            REM Run a start thread here to let the splash form show
            Me.tmrCoordinate = New Thread(AddressOf Me.runThread)
            With tmrCoordinate
                REM AddHandler .Tick, AddressOf Me.tmrCoordinate_Tick
                REM .Interval = 5
                .IsBackground = True
                .Start()
            End With

        End Sub

#End Region



#Region "Methods"

        Private ReadOnly Property hasAnyProcessLeft As Boolean
            Get
                Return Me.NextIntervalMarkPoint > 0
            End Get
        End Property


        Private Sub SetNextInterval()

            Select Case Me.ProcessesToRun.Count

                Case Is > 0
                    Me.NextIntervalMarkPoint = (Me.NormalInterval * Me.TotalProcessesRan) + Me.NormalInterval

                    ''Case Is > 1
                    ''    Me.NextIntervalMarkPoint = (Me.NormalInterval * Me.TotalProcessesRan) + Me.NormalInterval

                    ''Case Is = 1
                    ''    Me.NextIntervalMarkPoint = (Me.NormalInterval * Me.TotalProcessesRan) + Me.NormalInterval

                Case Else REM Is = 0
                    Me.NextIntervalMarkPoint = 0

            End Select
        End Sub

        Private Function PopNextAction() As Action
            Dim ac As Action = Me.ProcessesToRun(0)
            Me.ProcessesToRun.RemoveAt(0)
            Return ac
        End Function


        Private Function isProgressHandlerValid() As Boolean
            Return Me.ProgressHandler IsNot Nothing AndAlso Me.ProgressHandler.IsHandleCreated AndAlso Not Me.ProgressHandler.IsDisposed
        End Function


        Private Sub runThread()
            Try
                While Not Me.IsDisposed
                    Me.tmrCoordinate_Tick(Nothing, Nothing)
                    Thread.Sleep(5)
                End While
            Catch ex As Exception
                REM Abort or finished
            End Try

        End Sub


        Private Sub tmrCoordinate_Tick(sender As Object, e As EventArgs)

            ''REM Since this is multi Apartment
            ''REM wait for signal
            ''While Me.waitingForSignal

            ''    Application.DoEvents()
            ''End While

            If Me.IsDisposed Then
                REM If Me.tmrCoordinate IsNot Nothing Then Me.tmrCoordinate.Stop()
                Throw New Exception("that is enough")

            Else
                Me.CurrentCount += 1

                If Me.CurrentCount = Me.CountUpTo Then
                    If Me.isProgressHandlerValid() Then Me.ProgressHandler.Invoke(Sub() Me.ProgressHandler.onExitAction())
                    Me.disposedValue = True REM Me.tmrCoordinate.Stop()
                    Return
                ElseIf Me.CurrentCount = Me.NextIntervalMarkPoint Then
                    REM This forks the thread on some occassion
                    REM Me.waitingForSignal = True
                    Me.PopNextAction().Invoke()
                    Me.TotalProcessesRan = CUShort(Me.TotalProcessesRan + 1)

                    Me.SetNextInterval()
                End If

                If Me.isProgressHandlerValid() Then _
                    Me.ProgressHandler.Invoke(Sub() Me.ProgressHandler.ProgressChanged(Me.CountUpTo, Me.CurrentCount))

            End If
        End Sub


        ''' <summary>
        ''' Terminate this class quitely. It will call the onTerminateAction
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Terminate()
            If Not Me.IsDisposed Then
                Me.disposedValue = True
                Try

                    If Me.isProgressHandlerValid() Then Me.ProgressHandler.Invoke(Sub() Me.ProgressHandler.onTerminateAction())

                Catch ex As Exception
                    REM Ignore any error during invoke
                End Try
            End If
        End Sub


#End Region


#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls


        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return Me.disposedValue

            End Get
        End Property
        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    If Me.tmrCoordinate IsNot Nothing AndAlso Me.tmrCoordinate.IsAlive Then tmrCoordinate.Abort()
                    tmrCoordinate = Nothing

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
        ''' <summary>
        ''' This will abort the thread and will not return to your next line of call. Also, if you use application.exit .. 
        ''' That might not stop this thread for some time before it finally stops. So use terminate when necessary
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class



End Namespace