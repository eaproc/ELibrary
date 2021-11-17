using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;

namespace ELibrary.Standard.SplashCoordinator
{

    // Events runs on Co-ordinator Main thread
    // passed in actions runs on a separate background thread


    /// <summary>
    /// Initiates Coordinator on Constructor. Show your splash first before calling coordinator. Events runs on Main thread
    /// </summary>
    /// <remarks></remarks>
    public class ProcessCoordinator<Initiator> : IDisposable where Initiator : IProgressEvent, IProcessCoordinator, Control
    {


        #region Properties

        private Thread tmrCoordinate;
        private List<Action> ProcessesToRun;
        // REM         Private ExitAction As Action

        private ushort TotalProcessToRunCount;
        private ushort TotalProcessesRan;
        private ushort CountUpTo;
        private int CurrentCount;
        private int NextIntervalMarkPoint;
        private ushort NormalInterval;
        private Initiator ProgressHandler;

        public Initiator Initializer
        {
            get
            {
                return ProgressHandler;
            }
        }

        // REM Private waitingForSignal As Boolean

        #endregion




        #region Constructors

        /// <summary>
        /// Organizes the way processes are run under splash screens. Events runs on Main thread
        /// </summary>
        /// <param name="__CountUpTo">Must be greater than total process times 5</param>
        /// <param name="__ProcessesToRun">Processes you wish to execute under splash screens</param>
        /// <remarks></remarks>
        public ProcessCoordinator(ushort __CountUpTo, Initiator __ProgressHandler, params Action[] __ProcessesToRun)
        {
            ProcessesToRun = new List<Action>();
            if (__ProcessesToRun is object)
                ProcessesToRun.AddRange(__ProcessesToRun);
            // REM If __ExitAction Is Nothing Then Throw New Exception("Exit Action Can not be Null")
            // REM Me.ExitAction = __ExitAction
            ProgressHandler = __ProgressHandler;
            CountUpTo = __CountUpTo;
            TotalProcessToRunCount = (ushort)ProcessesToRun.Count;
            if (TotalProcessToRunCount > 0)
            {
                if (CountUpTo < TotalProcessToRunCount * 5)
                    throw new Exception("Count Up To Parameter must be greater than  or equals to processes count * 5");
                NormalInterval = (ushort)Math.Round(CountUpTo / (double)TotalProcessToRunCount);
                CountUpTo += NormalInterval;  // REM Extend Counter
                SetNextInterval();
            }

            // REM Run a start thread here to let the splash form show
            tmrCoordinate = new Thread(runThread);
            {
                var withBlock = tmrCoordinate;
                // REM AddHandler .Tick, AddressOf Me.tmrCoordinate_Tick
                // REM .Interval = 5
                withBlock.IsBackground = true;
                withBlock.Start();
            }
        }

        #endregion



        #region Methods

        private bool hasAnyProcessLeft
        {
            get
            {
                return NextIntervalMarkPoint > 0;
            }
        }

        private void SetNextInterval()
        {
            switch (ProcessesToRun.Count)
            {
                case var @case when @case > 0:
                    {

                        // 'Case Is > 1
                        // '    Me.NextIntervalMarkPoint = (Me.NormalInterval * Me.TotalProcessesRan) + Me.NormalInterval

                        // 'Case Is = 1
                        // '    Me.NextIntervalMarkPoint = (Me.NormalInterval * Me.TotalProcessesRan) + Me.NormalInterval

                        NextIntervalMarkPoint = (ushort)(NormalInterval * TotalProcessesRan) + NormalInterval; // REM Is = 0
                        break;
                    }

                default:
                    {
                        NextIntervalMarkPoint = 0;
                        break;
                    }
            }
        }

        private Action PopNextAction()
        {
            var ac = ProcessesToRun[0];
            ProcessesToRun.RemoveAt(0);
            return ac;
        }

        private bool isProgressHandlerValid()
        {
            return ProgressHandler is object && ProgressHandler.IsHandleCreated && !ProgressHandler.IsDisposed;
        }

        private void runThread()
        {
            try
            {
                while (!IsDisposed)
                {
                    tmrCoordinate_Tick(null, null);
                    Thread.Sleep(5);
                }
            }
            catch (Exception ex)
            {
                // REM Abort or finished
            }
        }

        private void tmrCoordinate_Tick(object sender, EventArgs e)
        {

            // 'REM Since this is multi Apartment
            // 'REM wait for signal
            // 'While Me.waitingForSignal

            // '    Application.DoEvents()
            // 'End While

            if (IsDisposed)
            {
                // REM If Me.tmrCoordinate IsNot Nothing Then Me.tmrCoordinate.Stop()
                throw new Exception("that is enough");
            }
            else
            {
                CurrentCount += 1;
                if (CurrentCount == CountUpTo)
                {
                    if (isProgressHandlerValid())
                        ProgressHandler.Invoke(new Action(() => ProgressHandler.onExitAction()));
                    disposedValue = true; // REM Me.tmrCoordinate.Stop()
                    return;
                }
                else if (CurrentCount == NextIntervalMarkPoint)
                {
                    // REM This forks the thread on some occassion
                    // REM Me.waitingForSignal = True
                    PopNextAction().Invoke();
                    TotalProcessesRan = (ushort)(TotalProcessesRan + 1);
                    SetNextInterval();
                }

                if (isProgressHandlerValid())
                    ProgressHandler.Invoke(new Action(() => ProgressHandler.ProgressChanged(CountUpTo, CurrentCount)));
            }
        }


        /// <summary>
        /// Terminate this class quitely. It will call the onTerminateAction
        /// </summary>
        /// <remarks></remarks>
        public void Terminate()
        {
            if (!IsDisposed)
            {
                disposedValue = true;
                try
                {
                    if (isProgressHandlerValid())
                        ProgressHandler.Invoke(new Action(() => ProgressHandler.onTerminateAction()));
                }
                catch (Exception ex)
                {
                    // REM Ignore any error during invoke
                }
            }
        }


        #endregion


        #region IDisposable Support
        private bool disposedValue; // To detect redundant calls

        public bool IsDisposed
        {
            get
            {
                return disposedValue;
            }
        }
        // IDisposable
        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                    if (tmrCoordinate is object && tmrCoordinate.IsAlive)
                        tmrCoordinate.Abort();
                    tmrCoordinate = null;
                }

                // TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                // TODO: set large fields to null.
            }

            disposedValue = true;
        }

        // TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
        // Protected Overrides Sub Finalize()
        // ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
        // Dispose(False)
        // MyBase.Finalize()
        // End Sub

        // This code added by Visual Basic to correctly implement the disposable pattern.
        /// <summary>
        /// This will abort the thread and will not return to your next line of call. Also, if you use application.exit .. 
        /// That might not stop this thread for some time before it finally stops. So use terminate when necessary
        /// </summary>
        /// <remarks></remarks>
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion

    }
}