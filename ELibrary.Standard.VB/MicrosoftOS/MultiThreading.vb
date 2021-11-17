Imports System.Threading

Imports CODERiT.Logger.Standard.VB.Exceptions
Imports ELibrary.Standard.VB.Types.EDelegates

Namespace MicrosoftOS
    ''' <summary>
    ''' This Class Helps Multi-Tasking
    ''' </summary>
    ''' <author>
    ''' Epro Company
    ''' </author>
    ''' <version>1.0.0</version>
    ''' <remarks></remarks>
    Public Class MultiThreading

#Region "Properties"

        ''' <summary>
        ''' This the threads container that holds all threads
        ''' </summary>
        ''' <remarks></remarks>
        Protected ThreadsAccumulated As List(Of Thread)

        ''' <summary>
        ''' Only Enables this class to exit threads without aborting them during execution
        ''' </summary>
        ''' <remarks></remarks>
        Private AwaitingExitThreads As List(Of Thread)



        '' ''' <summary>
        '' ''' Use to Manage Critical Sections for multi-threading
        '' ''' </summary>
        '' ''' <remarks></remarks>
        ''Public Structure CRITICAL_SECTION

        ''    REM Since the new thread coming is Holding down the process
        ''    REM Change of Algorithm
        ''    REM If a new thread comes and hit another thread ... Then Set the is waiting =true
        ''    REM.. .If it is time Up Sequence .. Quit current  thread and call the method again
        ''    REM Else Finish normally and call the method again





        ''    REM Implement a thread waiting protocol 
        ''    REM Like once another person comes the person currently using the place has to leave
        ''    Public Property EnableTimeUp As Boolean


        ''    ''' <summary>
        ''    ''' Indicates another thread is waiting to enter the critical section
        ''    ''' </summary>
        ''    ''' <value></value>
        ''    ''' <returns></returns>
        ''    ''' <remarks></remarks>
        ''    Public Property aThreadisWaiting As Boolean



        ''    ''' <summary>
        ''    ''' Indicate if a thread is Using Critical Section
        ''    ''' </summary>
        ''    ''' <value></value>
        ''    ''' <returns></returns>
        ''    ''' <remarks></remarks>
        ''    Private Property isInUse As Boolean
        ''    '' ''' <summary>
        ''    '' ''' Keep Count of Threads Waiting to Enter Critical Section
        ''    '' ''' </summary>
        ''    '' ''' <remarks></remarks>
        ''    ''Public WaitingThreads As Int16


        ''    ''' <summary>
        ''    ''' Enter Section and Lock for Entrance
        ''    ''' </summary>
        ''    ''' <remarks></remarks>
        ''    Public Function EnterSection(ByVal ThreadID As Int32) As Boolean

        ''        Me.aThreadisWaiting = True

        ''        If Me.isInUse Then Return False
        ''        Me.isInUse = True
        ''        REM Once the thread has been allowed in then it is no more waiting
        ''        Me.aThreadisWaiting = False
        ''        ''Me.WaitingThreads -= 1

        ''        ' Debug.Print("I just Entered {0}", ThreadID)

        ''        Return Me.isInUse

        ''    End Function

        ''    ''' <summary>
        ''    ''' Unlock For other threads to use
        ''    ''' </summary>
        ''    ''' <remarks></remarks>
        ''    Public Sub LeaveSection(ByVal ThreadID As Int32)
        ''        Me.isInUse = False

        ''        'Debug.Print("I am leaving {0}", ThreadID)

        ''        Me.aThreadisWaiting = False


        ''    End Sub

        ''End Structure

        ''' <summary>
        ''' Get the numbers of thread in this class
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property ThreadsCount As Long
            Get
                Return Me.ThreadsAccumulated.Count
            End Get
        End Property

#End Region


#Region "Constructors"
        Sub New()
            ThreadsAccumulated = New List(Of Thread)
            AwaitingExitThreads = New List(Of Thread)
        End Sub

#End Region

#Region "Methods"

#Region "Public"

        'Public Sub PlaceThreadInInfiniteLoop(
        '                                   ByVal OwnerForm As Form,
        '                                   ByVal CallBackFunc As Action(Of Thread),
        '                                   Optional ByVal SleepTimeMilliseconds As Long = 1000)
        '    Dim Sender As Thread = Nothing
        '    Sender = New Thread(
        '                New ThreadStart(
        '                                Sub() RunAlways(Sender:=Sender,
        '                                                CallBackFunc:=CallBackFunc,
        '                                                OwnerForm:=OwnerForm,
        '                                                SleepTimeMilliseconds:=SleepTimeMilliseconds
        '                                                )
        '                                 )
        '                         ) With {.IsBackground = True}



        '    ThreadsAccumulated.Add(Sender)

        '    Sender.Start()

        'End Sub

#Region "Sending the calling Thread"

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="SendThread">Dummy Overload</param>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Public Sub PlaceThreadInInfiniteLoop(
                                         ByVal SendThread As Boolean,
                                         ByVal CallBackFunc As Action(Of Thread),
                                         Optional ByVal SleepTimeMilliseconds As Long = 1000)


            Dim Sender As Thread = Nothing
            Sender = New Thread(
                        New ThreadStart(
                                        Sub() RunAlways(Sender:=Sender,
                                                        CallBackFunc:=CallBackFunc,
                                                        SleepTimeMilliseconds:=SleepTimeMilliseconds
                                                        )
                                         )
                                 ) With {.IsBackground = True}



            ThreadsAccumulated.Add(Sender)

            Sender.Start()

        End Sub


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="SendThread">Dummy Overload</param>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Public Sub PlaceThreadInInfiniteLoop(
                                         ByVal SendThread As Boolean,
                                          ByVal CallBackFunc As delegateSubBoolThread,
                                         ByVal BoolVal As Boolean,
                                         Optional ByVal SleepTimeMilliseconds As Long = 1000)


            Dim Sender As Thread = Nothing
            Sender = New Thread(
                        New ThreadStart(
                                        Sub() RunAlways(Sender:=Sender,
                                                        BoolVal:=BoolVal,
                                                        CallBackFunc:=CallBackFunc,
                                                        SleepTimeMilliseconds:=SleepTimeMilliseconds
                                                        )
                                         )
                                 ) With {.IsBackground = True}



            ThreadsAccumulated.Add(Sender)

            Sender.Start()

        End Sub


#End Region


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="BoolVal" >Boolean Value</param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Public Sub PlaceThreadInInfiniteLoop(
                                         ByVal CallBackFunc As delegateSubBool,
                                         ByVal BoolVal As Boolean,
                                         Optional ByVal SleepTimeMilliseconds As Long = 1000)


            Dim Sender As Thread = New Thread(
                        New ThreadStart(
                                        Sub() RunAlways(
                                                        CallBackFunc:=CallBackFunc, BoolVal:=BoolVal,
                                                        SleepTimeMilliseconds:=SleepTimeMilliseconds
                                                        )
                                         )
                                 ) With {.IsBackground = True}



            ThreadsAccumulated.Add(Sender)

            Sender.Start()

        End Sub


        ''' <summary>
        ''' Place a thread in inFiniteLoop
        ''' </summary>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Public Function PlaceThreadInInfiniteLoop(
                                           ByVal CallBackFunc As delegateNoParam,
                                           Optional ByVal SleepTimeMilliseconds As Long = 1000) As Thread
            Dim Sender As Thread = New Thread(
                        New ThreadStart(
                                        Sub() RunAlways(
                                                        CallBackFunc:=CallBackFunc,
                                                        SleepTimeMilliseconds:=SleepTimeMilliseconds
                                                        )
                                         )
                                 ) With {.IsBackground = True}

            ThreadsAccumulated.Add(Sender)

            Sender.Start()

            Return Sender

        End Function

#Region "Run Once"


        ''' <summary>
        ''' Run Thread Once
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RunOnce(
                            ByVal CallBackFunc As Action)
            Dim ThreadTemp As New Thread(New ThreadStart(Sub() CallBackFunc())) With {.IsBackground = True}
            ThreadTemp.Start()
        End Sub

        'Public Sub RunOnce(ByVal OwnerForm As Form,
        '                  ByVal CallBackFunc As Action(Of Boolean),
        '                        Optional ByVal Param As Boolean = False)
        '    Try
        '        Dim ThreadTemp As New Thread(New ThreadStart(
        '                                       Sub() InvokeRun(OwnerForm, CallBackFunc, Param)
        '                                    )
        '                                 ) With {.IsBackground = True}
        '        ThreadTemp.Start()

        '    Catch ex As Exception

        '    End Try

        'End Sub

#End Region


        Public Sub Abort(ByVal Thr As Thread)
            Dim ind As Long = ThreadsAccumulated.IndexOf(Thr)
            If ind >= 0 And ThreadsAccumulated.Count > 0 Then
                REM Remove from awaiting too if it is there
                If Me.AwaitingExitThreads.Contains(Thr) Then Me.AwaitingExitThreads.Remove(Thr)

                ThreadsAccumulated.Item(CInt(ind)).Abort()
                ThreadsAccumulated.RemoveAt(CInt(ind))

            End If
        End Sub

        Public Sub AbortAll()
            If ThreadsAccumulated.Count <= 0 Then Return
            Try
                For Each Thr As Thread In ThreadsAccumulated
                    If Thr.IsAlive Then Thr.Abort()
                    REM Remove from awaiting too if it is there
                    If Me.AwaitingExitThreads.Contains(Thr) Then Me.AwaitingExitThreads.Remove(Thr)
                Next
                ThreadsAccumulated.RemoveRange(
                    ThreadsAccumulated.IndexOf(ThreadsAccumulated.First),
                    ThreadsAccumulated.Count)

            Catch ex As Exception
                Debug.Print(ex.Message)
            End Try
        End Sub



        ''' <summary>
        ''' Add to threads that are waiting to be removed
        ''' </summary>
        ''' <param name="ThreadID"></param>
        ''' <remarks></remarks>
        Public Sub ExitAsyncThread(ByVal ThreadID As Thread)
            If Not Me.AwaitingExitThreads.Contains(ThreadID) Then
                Me.AwaitingExitThreads.Add(ThreadID)
            End If

        End Sub


        ''' <summary>
        ''' Returns true if this process succeeds on fetching this lock else the lock with same name is already in use. ReleaseMutex when you finish using it
        ''' </summary>
        ''' <param name="Key"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getOSLockUsing(ByRef AppMutex As System.Threading.Mutex,
                                              ByVal Key As String) As Boolean
            REM Dim AppMutex As System.Threading.Mutex = Nothing
            REM Lock using mutex
            AppMutex = New System.Threading.Mutex(False, Key)
            Return AppMutex.WaitOne(TimeSpan.Zero, False)

        End Function


#End Region

#Region "Private"

        'Private Sub RunAlways(ByVal Sender As Thread,
        '                     ByVal OwnerForm As Form,
        '                     ByVal CallBackFunc As Action(Of Thread),
        '                     Optional ByVal SleepTimeMilliseconds As Long = 1000)
        '    Try
        '        While True

        '            If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited before Execution Without Abortion Error.") : Exit While

        '            OwnerForm.Invoke(Sub() CallBackFunc(Sender))

        '            If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited after Execution Without Abortion Error.") : Exit While

        '            Thread.Sleep(CInt(SleepTimeMilliseconds))
        '        End While

        '    Catch ex As System.Threading.ThreadAbortException
        '        REM Only Expecting Error Here
        '        MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways")

        '    Catch ex As Exception
        '        If Not ex Is Nothing Then
        '            MyLogFile.Log(New EException("Thread Aborted While Sleeping Under RunAlways", ex))
        '        End If
        '    End Try
        'End Sub

        Private Sub RunAlways(ByVal Sender As Thread,
                            ByVal CallBackFunc As Action(Of Thread),
                            Optional ByVal SleepTimeMilliseconds As Long = 1000)
            Try
                While True

                    If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited before Execution Without Abortion Error.") : Exit While

                    Call CallBackFunc(Sender)

                    If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited after Execution Without Abortion Error.") : Exit While

                    Thread.Sleep(CInt(SleepTimeMilliseconds))
                End While

            Catch ex As System.Threading.ThreadAbortException
                REM Only Expecting Error Here
                MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways")

            Catch ex As Exception
                If Not ex Is Nothing Then
                    MyLogFile.Log(New EException("Thread Aborted While Sleeping Under RunAlways", ex))
                End If
            End Try
        End Sub

        Private Sub RunAlways(ByVal Sender As Thread,
                              ByVal BoolVal As Boolean,
                            ByVal CallBackFunc As delegateSubBoolThread,
                            Optional ByVal SleepTimeMilliseconds As Long = 1000)
            Try
                While True

                    If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited before Execution Without Abortion Error.") : Exit While

                    Call CallBackFunc(BoolVal, Sender)

                    If Me.AwaitingExitThreads.Contains(Sender) Then MyLogFile.Log("Thread Exited after Execution Without Abortion Error.") : Exit While

                    Thread.Sleep(CInt(SleepTimeMilliseconds))
                End While

            Catch ex As System.Threading.ThreadAbortException
                REM Only Expecting Error Here
                MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways")

            Catch ex As Exception
                If Not ex Is Nothing Then
                    MyLogFile.Log(New EException("Thread Aborted While Sleeping Under RunAlways", ex))
                End If
            End Try
        End Sub


        ''' <summary>
        ''' Note: You will need to check for invokeRequired
        ''' </summary>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Private Sub RunAlways(ByVal CallBackFunc As delegateSubBool,
                              ByVal BoolVal As Boolean,
                            Optional ByVal SleepTimeMilliseconds As Long = 1000)
            Try
                While True

                    Call CallBackFunc(BoolVal)
                    Thread.Sleep(CInt(SleepTimeMilliseconds))
                End While

            Catch ex As System.Threading.ThreadAbortException
                REM Only Expecting Error Here
                MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways")

            Catch ex As Exception
                If Not ex Is Nothing Then
                    MyLogFile.Log(New EException("Thread Aborted While Sleeping Under RunAlways", ex))
                End If
            End Try
        End Sub


        ''' <summary>
        ''' Keeps Running Thread
        ''' </summary>
        ''' <param name="CallBackFunc"></param>
        ''' <param name="SleepTimeMilliseconds"></param>
        ''' <remarks></remarks>
        Private Sub RunAlways(
                             ByVal CallBackFunc As delegateNoParam,
                             Optional ByVal SleepTimeMilliseconds As Long = 1000)
            Try
                While True
                    REM Enable Close Async"
                    'If Thread.CurrentThread.Name = "CloseAsync" Then Debug.Print("Thread Closed Easily") : Exit While

                    Call CallBackFunc()
                    Thread.Sleep(CInt(SleepTimeMilliseconds))
                End While

            Catch ex As System.Threading.ThreadAbortException
                REM Only Expecting Error Here
                MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways")

            Catch ex As Exception
                If Not ex Is Nothing Then
                    MyLogFile.Log(New EException("Thread Aborted While Sleeping Under RunAlways", ex))
                End If
            End Try
        End Sub



        ''Private Sub InvokeRun(ByVal OwnerForm As Form,
        ''                     ByVal CallBackFunc As Action
        ''                     )
        ''    OwnerForm.Invoke(CallBackFunc)
        ''End Sub
        Private Sub InvokeRun(ByVal OwnerForm As Form,
                          ByVal CallBackFunc As Action(Of Boolean),
                                Optional ByVal Param As Boolean = False
                         )
            OwnerForm.Invoke(Sub() CallBackFunc(Param))
        End Sub

#End Region

#End Region



    End Class

End Namespace