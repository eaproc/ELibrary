Imports System.Threading

Imports EPRO.Library.v3._5.Objects


Namespace TimerCore

''' <summary>
''' Not reusable. Only initialize and dispose. Also, it doesnt handle cross-threading
''' </summary>
''' <remarks></remarks>
Public Class SecsCounterBasic
    Implements IDisposable


    ''' <summary>
    ''' Not reusable. Only initialize and dispose. Also, it doesnt handle cross-threading
    ''' </summary>
    ''' <param name="MinsToCountTo"></param>
    ''' <remarks></remarks>
    Public Sub New(MinsToCountTo As Long)

        Me._______CountToSecs = MinsToCountTo * 60
        ____CustomSecsLeftAlert = New List(Of Long)()


        _____onSecondIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE
        _____onMinuteIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE
        _____onTimeLeft___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE
        _____onDayChanged___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE

            _____onSecondIncrement___Thread_____IsAlive = False

            Me.__PreviousStartedDateTime = Types.NullableDateTime.NULL_TIME
    End Sub



#Region "Events"

        Public Event onSecondIncrement(sender As SecsCounterBasic)
        Public Event onMinuteIncrement(sender As SecsCounterBasic)
        Public Event onHalfMinuteIncrement(sender As SecsCounterBasic)
        Public Event onQuarterMinuteIncrement(sender As SecsCounterBasic)


        Public Event onTimeUp(sender As SecsCounterBasic)
        Public Event onTimeLeft(sender As SecsCounterBasic)

        ''' <summary>
        ''' It auto reset the day to new date after calling this event
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <remarks></remarks>
        Public Event onDayChanged(sender As SecsCounterBasic)

#End Region


#Region "Properties"

        Public Shared Property IsDebugMode As Boolean


        Private _____thrMainSecCounter As Thread
        Private _____thr___onSecondIncrement_Informer As Thread

        Private _____thr___onMinuteIncrement_Informer As Thread
        Private _____thr___onHalfMinuteIncrement_Informer As Thread
        Private _____thr___onQuarterMinuteIncrement_Informer As Thread

        Private _____thr___onTimeLeft_Informer As Thread
        Private _____thr___onDayChanged_Informer As Thread



        Private _____onSecondIncrement___Thread_____IsAlive As Boolean
        Private _____onSecondIncrement___OccurringAction As RecurringThreadAction
        ''' <summary>
        ''' This is the same thing Other big intervals will use
        ''' </summary>
        ''' <remarks></remarks>
        Private _____onMinuteIncrement___OccurringAction As RecurringThreadAction

        Private _____onTimeLeft___OccurringAction As RecurringThreadAction
        Private _____onDayChanged___OccurringAction As RecurringThreadAction




        Private ____SecsUsed As Long, _______CountToSecs As Long,
            __StartedTime As DateTime,
            __PreviousStartedDateTime As Types.NullableDateTime

        Private ____CustomSecsLeftAlert As List(Of Long)

    Public ReadOnly Property MinsUsed As Int32
        Get
            Return CInt(EDateTime.getMins(Me.SecsUsed))
        End Get
    End Property

    Public ReadOnly Property SecsUsed As Long
        Get
            Return ____SecsUsed
        End Get
    End Property

    Public ReadOnly Property MinsLeft As Int32
        Get
            Return CInt(EDateTime.getMins(Me.SecsLeft))
        End Get
    End Property

    Public ReadOnly Property SecsLeft As Long
        Get
            Return _______CountToSecs - ____SecsUsed
        End Get
    End Property

        Public ReadOnly Property StartedTime As DateTime
            Get
                Return __StartedTime
            End Get
        End Property

        ''' <summary>
        ''' Contains the previous date time before the Started Time in this class changed. 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property PreviousStartedDateTime As Types.NullableDateTime
            Get
                Return Me.__PreviousStartedDateTime
            End Get
        End Property

#End Region


#Region "Enums and Const"

    Public Enum RecurringThreadAction

        IGNORE___IF_IN_USE
        ABORT___IF_IN_USE

    End Enum

#End Region


#Region "Methods"

    ''' <summary>
    ''' Only used once
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Start()

        If Me.disposedValue Then Throw New ObjectDisposedException(Me.GetType().Name)
        If Me._____thrMainSecCounter IsNot Nothing Then Return

        Me._____thrMainSecCounter = New Thread(AddressOf Me.Run) With {.Name = "_____thrMainSecCounter", .IsBackground = True}
        _____thrMainSecCounter.SetApartmentState(ApartmentState.MTA)
        _____thrMainSecCounter.Start()


    End Sub


        Public Sub Run()

            If IsDebugMode Then Program.Logger.Print("SecsCounterBasic Starting ...")
            Try
                Me.__StartedTime = Now
                While Not Me.disposedValue

                    If Me.____SecsUsed >= Me._______CountToSecs Then RaiseEvent onTimeUp(Me) : Exit While


                    Me.____SecsUsed += 1
                    Me.Invoke___onSecondIncrement()


                    If Me.____SecsUsed Mod 60 = 0 Then Me.Invoke___onMinuteIncrement()

                    If Me.____SecsUsed Mod 30 = 0 Then Me.Invoke___onHalfMinuteIncrement()

                    If Me.____SecsUsed Mod 15 = 0 Then Me.Invoke___onQuarterMinuteIncrement()


                    If Not EDateTime.EqualsDateWithoutTime(Now, Me.StartedTime) Then
                        Me.Invoke___onDayChanged()
                        Me.__PreviousStartedDateTime = New Types.NullableDateTime(Me.StartedTime)
                        Me.__StartedTime = Now
                    End If



                    If Me.____CustomSecsLeftAlert.Contains(Me.SecsLeft) Then
                        Me.____CustomSecsLeftAlert.Remove(Me.SecsLeft)
                        Me.Invoke___onTimeLeft()

                    End If


                    Thread.Sleep(1000)
                End While

            Catch ex As ThreadAbortException

            Catch ex As Exception
                Program.Logger.Print(ex)

            Finally

                If IsDebugMode Then Program.Logger.Print(Thread.CurrentThread.Name & " is quiting ...")

            End Try

            Me.disposedValue = True
        End Sub





        ''' <summary>
        ''' Use to explain which action the thread delivering this event should take if it is blocked (In-Use) on the client side
        ''' </summary>
        ''' <param name="pRecurringThreadAction"></param>
        ''' <remarks></remarks>
    Public Sub Set_____onSecondIncrement___OccurringAction(pRecurringThreadAction As RecurringThreadAction)
        Me._____onSecondIncrement___OccurringAction = pRecurringThreadAction
    End Sub
        ''' <summary>
        ''' Use to explain which action the thread delivering this event should take if it is blocked (In-Use) on the client side
        ''' </summary>
        ''' <param name="pRecurringThreadAction"></param>
        ''' <remarks></remarks>
    Public Sub Set_____onMinuteIncrement___OccurringAction(pRecurringThreadAction As RecurringThreadAction)
        Me._____onMinuteIncrement___OccurringAction = pRecurringThreadAction

    End Sub
        ''' <summary>
        ''' Use to explain which action the thread delivering this event should take if it is blocked (In-Use) on the client side
        ''' </summary>
        ''' <param name="pRecurringThreadAction"></param>
        ''' <remarks></remarks>
    Public Sub Set_____onTimeLeft___OccurringAction(pRecurringThreadAction As RecurringThreadAction)
        Me._____onTimeLeft___OccurringAction = pRecurringThreadAction

    End Sub
        ''' <summary>
        ''' Use to explain which action the thread delivering this event should take if it is blocked (In-Use) on the client side
        ''' </summary>
        ''' <param name="pRecurringThreadAction"></param>
        ''' <remarks></remarks>
    Public Sub Set_____onDayChanged___OccurringAction(pRecurringThreadAction As RecurringThreadAction)
        Me._____onDayChanged___OccurringAction = pRecurringThreadAction

    End Sub



        ''' <summary>
        ''' Use to set at which seconds you want this class to raise Time Left. 
        ''' Also, you can set what action the thread will take using Set_____onTimeLeft___OccurringAction
        ''' </summary>
        ''' <param name="p____CustomSecsLeftAlert"></param>
        ''' <remarks></remarks>
    Public Sub Set____CustomSecsLeftAlert__Intervals(ParamArray p____CustomSecsLeftAlert As Long())
        If Me._____thrMainSecCounter IsNot Nothing Then Return '   Started Already

        If p____CustomSecsLeftAlert IsNot Nothing Then
            Me.____CustomSecsLeftAlert = p____CustomSecsLeftAlert.Distinct().ToList()
        End If

    End Sub











    Private Sub Run___onSecondIncrement()
            Try
                If IsDebugMode Then Program.Logger.Print("Run___onSecondIncrement: - Me.SecsLeft: " & Me.SecsLeft)
                RaiseEvent onSecondIncrement(Me)

                Debug.Print("Run___onSecondIncrement() Event Forced Call") 'This line is just to make sure the raiseevent calls
                If IsDebugMode Then Program.Logger.Print("Run___onSecondIncrement: - After Me.SecsLeft: " & Me.SecsLeft)

            Catch ex As Exception
                If IsDebugMode Then Program.Logger.Print("Run___onSecondIncrement", ex)
            Finally
                Me._____onSecondIncrement___Thread_____IsAlive = False
            End Try
    End Sub

    Private Sub Invoke___onSecondIncrement()
        Try

                If IsDebugMode Then Program.Logger.Print("Trying to Call _____onSecondIncrement___Thread_____IsAlive: " &
                    _____onSecondIncrement___Thread_____IsAlive)

                Try

                    If Me._____onSecondIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                        If Me._____thr___onSecondIncrement_Informer IsNot Nothing AndAlso Me._____onSecondIncrement___Thread_____IsAlive Then Return

                    ElseIf Me._____onSecondIncrement___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                        If Me._____thr___onSecondIncrement_Informer IsNot Nothing AndAlso Me._____onSecondIncrement___Thread_____IsAlive Then _____thr___onSecondIncrement_Informer.Abort()
                        _____thr___onSecondIncrement_Informer = Nothing


                    End If


                Catch ex As Exception
                    '
                    '   Am not expecting error here, however if there is an error, do not proceed to next step
                    '
                    Program.Logger.Print(ex)
                    Return
                End Try


                Me._____thr___onSecondIncrement_Informer = New Thread(AddressOf Me.Run___onSecondIncrement) With {.Name = "_____thr___onSecondIncrement_Informer", .IsBackground = True}
                _____thr___onSecondIncrement_Informer.SetApartmentState(ApartmentState.MTA)
                _____thr___onSecondIncrement_Informer.Start()
                Me._____onSecondIncrement___Thread_____IsAlive = True
                If IsDebugMode Then Program.Logger.Print("_____onSecondIncrement___Thread_____IsAlive: Invoked")
            Catch ex As Exception
                Program.Logger.Print("Invoke___onSecondIncrement", ex)
            End Try
    End Sub



    Private Sub Run___onMinuteIncrement()
        Try
                RaiseEvent onMinuteIncrement(Me)
                Debug.Print("onMinuteIncrement() Event Forced Call") 'This line is just to make sure the raiseevent calls

        Catch ex As Exception
            Program.Logger.Print("Run___onMinuteIncrement", ex)
        End Try
    End Sub

    Private Sub Invoke___onMinuteIncrement()
        Try

            If Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                If Me._____thr___onMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onMinuteIncrement_Informer.IsAlive Then Return

            ElseIf Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                If Me._____thr___onMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onMinuteIncrement_Informer.IsAlive Then _____thr___onMinuteIncrement_Informer.Abort()
                _____thr___onMinuteIncrement_Informer = Nothing


            End If

            Me._____thr___onMinuteIncrement_Informer = New Thread(AddressOf Me.Run___onMinuteIncrement) With {.Name = "_____thr___onMinuteIncrement_Informer", .IsBackground = True}
            _____thr___onMinuteIncrement_Informer.SetApartmentState(ApartmentState.MTA)
            _____thr___onMinuteIncrement_Informer.Start()


        Catch ex As Exception
            Program.Logger.Print("Invoke___onMinuteIncrement", ex)
        End Try
    End Sub





        Private Sub Run___onHalfMinuteIncrement()
            Try
                RaiseEvent onHalfMinuteIncrement(Me)
                Debug.Print("onHalfMinuteIncrement() Event Forced Call") 'This line is just to make sure the raiseevent calls

            Catch ex As Exception
                Program.Logger.Print("Run___onHalfMinuteIncrement", ex)
            End Try
        End Sub

        Private Sub Invoke___onHalfMinuteIncrement()
            Try

                If Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                    If Me._____thr___onHalfMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onHalfMinuteIncrement_Informer.IsAlive Then Return

                ElseIf Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                    If Me._____thr___onHalfMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onHalfMinuteIncrement_Informer.IsAlive Then _____thr___onHalfMinuteIncrement_Informer.Abort()
                    _____thr___onHalfMinuteIncrement_Informer = Nothing


                End If

                Me._____thr___onHalfMinuteIncrement_Informer = New Thread(AddressOf Me.Run___onHalfMinuteIncrement) With {.Name = "_____thr___onHalfMinuteIncrement_Informer", .IsBackground = True}
                _____thr___onHalfMinuteIncrement_Informer.SetApartmentState(ApartmentState.MTA)
                _____thr___onHalfMinuteIncrement_Informer.Start()


            Catch ex As Exception
                Program.Logger.Print("Invoke___onHalfMinuteIncrement", ex)
            End Try
        End Sub






        Private Sub Run___onQuarterMinuteIncrement()
            Try
                RaiseEvent onQuarterMinuteIncrement(Me)
                Debug.Print("onQuarterMinuteIncrement() Event Forced Call") 'This line is just to make sure the raiseevent calls

            Catch ex As Exception
                Program.Logger.Print("Run___onQuarterMinuteIncrement", ex)
            End Try
        End Sub

        Private Sub Invoke___onQuarterMinuteIncrement()
            Try

                If Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                    If Me._____thr___onQuarterMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onQuarterMinuteIncrement_Informer.IsAlive Then Return

                ElseIf Me._____onMinuteIncrement___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                    If Me._____thr___onQuarterMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onQuarterMinuteIncrement_Informer.IsAlive Then _____thr___onQuarterMinuteIncrement_Informer.Abort()
                    _____thr___onQuarterMinuteIncrement_Informer = Nothing


                End If

                Me._____thr___onQuarterMinuteIncrement_Informer = New Thread(AddressOf Me.Run___onQuarterMinuteIncrement) With {.Name = "_____thr___onQuarterMinuteIncrement_Informer", .IsBackground = True}
                _____thr___onQuarterMinuteIncrement_Informer.SetApartmentState(ApartmentState.MTA)
                _____thr___onQuarterMinuteIncrement_Informer.Start()


            Catch ex As Exception
                Program.Logger.Print("Invoke___onQuarterMinuteIncrement", ex)
            End Try
        End Sub







    Private Sub Run___onTimeLeft()
        Try
                RaiseEvent onTimeLeft(Me)
                Debug.Print("Run___onTimeLeft() Event Forced Call") 'This line is just to make sure the raiseevent calls

        Catch ex As Exception
            Program.Logger.Print("Run___onTimeLeft", ex)
        End Try
    End Sub

    Private Sub Invoke___onTimeLeft()
        Try


            If Me._____onTimeLeft___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                If Me._____thr___onTimeLeft_Informer IsNot Nothing AndAlso Me._____thr___onTimeLeft_Informer.IsAlive Then Return

            ElseIf Me._____onTimeLeft___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                If Me._____thr___onTimeLeft_Informer IsNot Nothing AndAlso Me._____thr___onTimeLeft_Informer.IsAlive Then _____thr___onTimeLeft_Informer.Abort()
                _____thr___onTimeLeft_Informer = Nothing


            End If

            Me._____thr___onTimeLeft_Informer = New Thread(AddressOf Me.Run___onTimeLeft) With {.Name = "_____thr___onTimeLeft_Informer", .IsBackground = True}
            _____thr___onTimeLeft_Informer.SetApartmentState(ApartmentState.MTA)
            _____thr___onTimeLeft_Informer.Start()


        Catch ex As Exception
            Program.Logger.Print("Invoke___onTimeLeft", ex)
        End Try
    End Sub





    Private Sub Run___onDayChanged()
        Try
                RaiseEvent onDayChanged(Me)
                Debug.Print("Run___onDayChanged() Event Forced Call") 'This line is just to make sure the raiseevent calls

        Catch ex As Exception
            Program.Logger.Print("Run___onDayChanged", ex)
        End Try
    End Sub

    Private Sub Invoke___onDayChanged()
        Try

            If Me._____onDayChanged___OccurringAction = RecurringThreadAction.IGNORE___IF_IN_USE Then

                If Me._____thr___onDayChanged_Informer IsNot Nothing AndAlso Me._____thr___onDayChanged_Informer.IsAlive Then Return

            ElseIf Me._____onDayChanged___OccurringAction = RecurringThreadAction.ABORT___IF_IN_USE Then

                If Me._____thr___onDayChanged_Informer IsNot Nothing AndAlso Me._____thr___onDayChanged_Informer.IsAlive Then _____thr___onDayChanged_Informer.Abort()
                _____thr___onDayChanged_Informer = Nothing


            End If

            Me._____thr___onDayChanged_Informer = New Thread(AddressOf Me.Run___onDayChanged) With {.Name = "_____thr___onDayChanged_Informer", .IsBackground = True}
            _____thr___onDayChanged_Informer.SetApartmentState(ApartmentState.MTA)
            _____thr___onDayChanged_Informer.Start()



        Catch ex As Exception
            Program.Logger.Print("Invoke___onDayChanged", ex)
        End Try
    End Sub







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
                ' Handle Each Thread Individually incase user dispose thread from inside same thread
                ' TODO: dispose managed state (managed objects).
                Try


                    If Me._____thr___onDayChanged_Informer IsNot Nothing AndAlso Me._____thr___onDayChanged_Informer.IsAlive Then Me._____thr___onDayChanged_Informer.Abort()
                    Me._____thr___onDayChanged_Informer = Nothing

                Catch ex As Exception
                    Program.Logger.Print(ex)
                End Try

                Try


                    If Me._____thr___onMinuteIncrement_Informer IsNot Nothing AndAlso Me._____thr___onMinuteIncrement_Informer.IsAlive Then Me._____thr___onMinuteIncrement_Informer.Abort()
                    Me._____thr___onMinuteIncrement_Informer = Nothing

                Catch ex As Exception
                    Program.Logger.Print(ex)
                End Try


                Try


                    If Me._____thr___onSecondIncrement_Informer IsNot Nothing AndAlso Me._____thr___onSecondIncrement_Informer.IsAlive Then Me._____thr___onSecondIncrement_Informer.Abort()
                    Me._____thr___onSecondIncrement_Informer = Nothing

                Catch ex As Exception
                    Program.Logger.Print(ex)
                End Try


                Try


                    If Me._____thr___onTimeLeft_Informer IsNot Nothing AndAlso Me._____thr___onTimeLeft_Informer.IsAlive Then Me._____thr___onTimeLeft_Informer.Abort()
                    Me._____thr___onTimeLeft_Informer = Nothing

                Catch ex As Exception
                    Program.Logger.Print(ex)
                End Try



                Try


                    If Me._____thrMainSecCounter IsNot Nothing AndAlso Me._____thrMainSecCounter.IsAlive Then Me._____thrMainSecCounter.Abort()
                    Me._____thrMainSecCounter = Nothing

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




End Namespace
