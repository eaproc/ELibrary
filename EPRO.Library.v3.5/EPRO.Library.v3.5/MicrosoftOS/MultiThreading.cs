using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using CODERiT.Logger.v._3._5.Exceptions;
using ELibrary.Standard.Modules;
using static ELibrary.Standard.Types.EDelegates;

namespace ELibrary.Standard.MicrosoftOS
{
    /// <summary>
    /// This Class Helps Multi-Tasking
    /// </summary>
    /// <author>
    /// Epro Company
    /// </author>
    /// <version>1.0.0</version>
    /// <remarks></remarks>
    public class MultiThreading
    {

        #region Properties

        /// <summary>
        /// This the threads container that holds all threads
        /// </summary>
        /// <remarks></remarks>
        protected List<Thread> ThreadsAccumulated;

        /// <summary>
        /// Only Enables this class to exit threads without aborting them during execution
        /// </summary>
        /// <remarks></remarks>
        private List<Thread> AwaitingExitThreads;



        // ' ''' <summary>
        // ' ''' Use to Manage Critical Sections for multi-threading
        // ' ''' </summary>
        // ' ''' <remarks></remarks>
        // 'Public Structure CRITICAL_SECTION

        // '    REM Since the new thread coming is Holding down the process
        // '    REM Change of Algorithm
        // '    REM If a new thread comes and hit another thread ... Then Set the is waiting =true
        // '    REM.. .If it is time Up Sequence .. Quit current  thread and call the method again
        // '    REM Else Finish normally and call the method again





        // '    REM Implement a thread waiting protocol 
        // '    REM Like once another person comes the person currently using the place has to leave
        // '    Public Property EnableTimeUp As Boolean


        // '    ''' <summary>
        // '    ''' Indicates another thread is waiting to enter the critical section
        // '    ''' </summary>
        // '    ''' <value></value>
        // '    ''' <returns></returns>
        // '    ''' <remarks></remarks>
        // '    Public Property aThreadisWaiting As Boolean



        // '    ''' <summary>
        // '    ''' Indicate if a thread is Using Critical Section
        // '    ''' </summary>
        // '    ''' <value></value>
        // '    ''' <returns></returns>
        // '    ''' <remarks></remarks>
        // '    Private Property isInUse As Boolean
        // '    '' ''' <summary>
        // '    '' ''' Keep Count of Threads Waiting to Enter Critical Section
        // '    '' ''' </summary>
        // '    '' ''' <remarks></remarks>
        // '    ''Public WaitingThreads As Int16


        // '    ''' <summary>
        // '    ''' Enter Section and Lock for Entrance
        // '    ''' </summary>
        // '    ''' <remarks></remarks>
        // '    Public Function EnterSection(ByVal ThreadID As Int32) As Boolean

        // '        Me.aThreadisWaiting = True

        // '        If Me.isInUse Then Return False
        // '        Me.isInUse = True
        // '        REM Once the thread has been allowed in then it is no more waiting
        // '        Me.aThreadisWaiting = False
        // '        ''Me.WaitingThreads -= 1

        // '        ' Debug.Print("I just Entered {0}", ThreadID)

        // '        Return Me.isInUse

        // '    End Function

        // '    ''' <summary>
        // '    ''' Unlock For other threads to use
        // '    ''' </summary>
        // '    ''' <remarks></remarks>
        // '    Public Sub LeaveSection(ByVal ThreadID As Int32)
        // '        Me.isInUse = False

        // '        'Debug.Print("I am leaving {0}", ThreadID)

        // '        Me.aThreadisWaiting = False


        // '    End Sub

        // 'End Structure

        /// <summary>
        /// Get the numbers of thread in this class
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public long ThreadsCount
        {
            get
            {
                return ThreadsAccumulated.Count;
            }
        }

        #endregion


        #region Constructors
        public MultiThreading()
        {
            ThreadsAccumulated = new List<Thread>();
            AwaitingExitThreads = new List<Thread>();
        }

        #endregion

        #region Methods

        #region Public

        public void PlaceThreadInInfiniteLoop(Form OwnerForm, Action<Thread> CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            Thread Sender = null;
            Sender = new Thread(new ThreadStart(() => RunAlways(Sender: Sender, CallBackFunc: CallBackFunc, OwnerForm: OwnerForm, SleepTimeMilliseconds: SleepTimeMilliseconds))) { IsBackground = true };
            ThreadsAccumulated.Add(Sender);
            Sender.Start();
        }

        #region Sending the calling Thread

        /// <summary>
        /// 
        /// </summary>
        /// <param name="SendThread">Dummy Overload</param>
        /// <param name="CallBackFunc"></param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        public void PlaceThreadInInfiniteLoop(bool SendThread, Action<Thread> CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            Thread Sender = null;
            Sender = new Thread(new ThreadStart(() => RunAlways(Sender: Sender, CallBackFunc: CallBackFunc, SleepTimeMilliseconds: SleepTimeMilliseconds))) { IsBackground = true };
            ThreadsAccumulated.Add(Sender);
            Sender.Start();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="SendThread">Dummy Overload</param>
        /// <param name="CallBackFunc"></param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        public void PlaceThreadInInfiniteLoop(bool SendThread, delegateSubBoolThread CallBackFunc, bool BoolVal, long SleepTimeMilliseconds = 1000L)
        {
            Thread Sender = null;
            Sender = new Thread(new ThreadStart(() => RunAlways(Sender: Sender, BoolVal: BoolVal, CallBackFunc: CallBackFunc, SleepTimeMilliseconds: SleepTimeMilliseconds))) { IsBackground = true };
            ThreadsAccumulated.Add(Sender);
            Sender.Start();
        }


        #endregion


        /// <summary>
        /// 
        /// </summary>
        /// <param name="CallBackFunc"></param>
        /// <param name="BoolVal" >Boolean Value</param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        public void PlaceThreadInInfiniteLoop(delegateSubBool CallBackFunc, bool BoolVal, long SleepTimeMilliseconds = 1000L)
        {
            var Sender = new Thread(new ThreadStart(() => RunAlways(CallBackFunc: CallBackFunc, BoolVal: BoolVal, SleepTimeMilliseconds: SleepTimeMilliseconds))) { IsBackground = true };
            ThreadsAccumulated.Add(Sender);
            Sender.Start();
        }


        /// <summary>
        /// Place a thread in inFiniteLoop
        /// </summary>
        /// <param name="CallBackFunc"></param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        public Thread PlaceThreadInInfiniteLoop(delegateNoParam CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            var Sender = new Thread(new ThreadStart(() => RunAlways(CallBackFunc: CallBackFunc, SleepTimeMilliseconds: SleepTimeMilliseconds))) { IsBackground = true };
            ThreadsAccumulated.Add(Sender);
            Sender.Start();
            return Sender;
        }

        #region Run Once


        /// <summary>
        /// Run Thread Once
        /// </summary>
        /// <remarks></remarks>
        public void RunOnce(Action CallBackFunc)
        {
            var ThreadTemp = new Thread(new ThreadStart(() => CallBackFunc())) { IsBackground = true };
            ThreadTemp.Start();
        }

        public void RunOnce(Form OwnerForm, Action<bool> CallBackFunc, bool Param = false)
        {
            try
            {
                var ThreadTemp = new Thread(new ThreadStart(() => InvokeRun(OwnerForm, CallBackFunc, Param))) { IsBackground = true };
                ThreadTemp.Start();
            }
            catch (Exception ex)
            {
            }
        }

        #endregion


        public void Abort(Thread Thr)
        {
            long ind = ThreadsAccumulated.IndexOf(Thr);
            if (ind >= 0L & ThreadsAccumulated.Count > 0)
            {
                // REM Remove from awaiting too if it is there
                if (AwaitingExitThreads.Contains(Thr))
                    AwaitingExitThreads.Remove(Thr);
                ThreadsAccumulated[(int)ind].Abort();
                ThreadsAccumulated.RemoveAt((int)ind);
            }
        }

        public void AbortAll()
        {
            if (ThreadsAccumulated.Count <= 0)
                return;
            try
            {
                foreach (Thread Thr in ThreadsAccumulated)
                {
                    if (Thr.IsAlive)
                        Thr.Abort();
                    // REM Remove from awaiting too if it is there
                    if (AwaitingExitThreads.Contains(Thr))
                        AwaitingExitThreads.Remove(Thr);
                }

                ThreadsAccumulated.RemoveRange(ThreadsAccumulated.IndexOf(ThreadsAccumulated.First()), ThreadsAccumulated.Count);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }



        /// <summary>
        /// Add to threads that are waiting to be removed
        /// </summary>
        /// <param name="ThreadID"></param>
        /// <remarks></remarks>
        public void ExitAsyncThread(Thread ThreadID)
        {
            if (!AwaitingExitThreads.Contains(ThreadID))
            {
                AwaitingExitThreads.Add(ThreadID);
            }
        }


        /// <summary>
        /// Returns true if this process succeeds on fetching this lock else the lock with same name is already in use. ReleaseMutex when you finish using it
        /// </summary>
        /// <param name="Key"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool getOSLockUsing(ref Mutex AppMutex, string Key)
        {
            // REM Dim AppMutex As System.Threading.Mutex = Nothing
            // REM Lock using mutex
            AppMutex = new Mutex(false, Key);
            return AppMutex.WaitOne(TimeSpan.Zero, false);
        }


        #endregion

        #region Private

        private void RunAlways(Thread Sender, Form OwnerForm, Action<Thread> CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            try
            {
                while (true)
                {
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited before Execution Without Abortion Error.");
                        break;
                    }

                    OwnerForm.Invoke(new Action(() => CallBackFunc(Sender)));
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited after Execution Without Abortion Error.");
                        break;
                    }

                    Thread.Sleep((int)SleepTimeMilliseconds);
                }
            }
            catch (ThreadAbortException ex)
            {
                // REM Only Expecting Error Here
                basMain.MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways");
            }
            catch (Exception ex)
            {
                if (ex is object)
                {
                    basMain.MyLogFile.Log(new EException("Thread Aborted While Sleeping Under RunAlways", ex));
                }
            }
        }

        private void RunAlways(Thread Sender, Action<Thread> CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            try
            {
                while (true)
                {
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited before Execution Without Abortion Error.");
                        break;
                    }

                    CallBackFunc(Sender);
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited after Execution Without Abortion Error.");
                        break;
                    }

                    Thread.Sleep((int)SleepTimeMilliseconds);
                }
            }
            catch (ThreadAbortException ex)
            {
                // REM Only Expecting Error Here
                basMain.MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways");
            }
            catch (Exception ex)
            {
                if (ex is object)
                {
                    basMain.MyLogFile.Log(new EException("Thread Aborted While Sleeping Under RunAlways", ex));
                }
            }
        }

        private void RunAlways(Thread Sender, bool BoolVal, delegateSubBoolThread CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            try
            {
                while (true)
                {
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited before Execution Without Abortion Error.");
                        break;
                    }

                    CallBackFunc(BoolVal, Sender);
                    if (AwaitingExitThreads.Contains(Sender))
                    {
                        basMain.MyLogFile.Log("Thread Exited after Execution Without Abortion Error.");
                        break;
                    }

                    Thread.Sleep((int)SleepTimeMilliseconds);
                }
            }
            catch (ThreadAbortException ex)
            {
                // REM Only Expecting Error Here
                basMain.MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways");
            }
            catch (Exception ex)
            {
                if (ex is object)
                {
                    basMain.MyLogFile.Log(new EException("Thread Aborted While Sleeping Under RunAlways", ex));
                }
            }
        }


        /// <summary>
        /// Note: You will need to check for invokeRequired
        /// </summary>
        /// <param name="CallBackFunc"></param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        private void RunAlways(delegateSubBool CallBackFunc, bool BoolVal, long SleepTimeMilliseconds = 1000L)
        {
            try
            {
                while (true)
                {
                    CallBackFunc(BoolVal);
                    Thread.Sleep((int)SleepTimeMilliseconds);
                }
            }
            catch (ThreadAbortException ex)
            {
                // REM Only Expecting Error Here
                basMain.MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways");
            }
            catch (Exception ex)
            {
                if (ex is object)
                {
                    basMain.MyLogFile.Log(new EException("Thread Aborted While Sleeping Under RunAlways", ex));
                }
            }
        }


        /// <summary>
        /// Keeps Running Thread
        /// </summary>
        /// <param name="CallBackFunc"></param>
        /// <param name="SleepTimeMilliseconds"></param>
        /// <remarks></remarks>
        private void RunAlways(delegateNoParam CallBackFunc, long SleepTimeMilliseconds = 1000L)
        {
            try
            {
                while (true)
                {
                    // REM Enable Close Async"
                    // If Thread.CurrentThread.Name = "CloseAsync" Then Debug.Print("Thread Closed Easily") : Exit While

                    CallBackFunc();
                    Thread.Sleep((int)SleepTimeMilliseconds);
                }
            }
            catch (ThreadAbortException ex)
            {
                // REM Only Expecting Error Here
                basMain.MyLogFile.Log("Thread Aborted While Sleeping Under RunAlways");
            }
            catch (Exception ex)
            {
                if (ex is object)
                {
                    basMain.MyLogFile.Log(new EException("Thread Aborted While Sleeping Under RunAlways", ex));
                }
            }
        }



        // 'Private Sub InvokeRun(ByVal OwnerForm As Form,
        // '                     ByVal CallBackFunc As Action
        // '                     )
        // '    OwnerForm.Invoke(CallBackFunc)
        // 'End Sub
        private void InvokeRun(Form OwnerForm, Action<bool> CallBackFunc, bool Param = false)
        {
            OwnerForm.Invoke(new Action(() => CallBackFunc(Param)));
        }

        #endregion

        #endregion



    }
}