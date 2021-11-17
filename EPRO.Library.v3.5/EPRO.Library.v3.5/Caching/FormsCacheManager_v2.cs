using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using CODERiT.Logger.v._3._5;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.Caching
{
    public class FormsCacheManager_v2 : IDisposable
    {


        #region Constructors

        /// <summary>
        /// 'Careat a cach manager with your own initializer if you want
        /// </summary>
        /// <param name="___InitializeFormsSiliently"></param>
        /// <param name="pFormLoader"></param>
        /// <param name="pRunResolutionMonitor"></param>
        /// <remarks></remarks>
        public FormsCacheManager_v2(bool ___InitializeFormsSiliently, Control pFormLoader, CacheFormLoaderMode pLoaderMode, bool pRunResolutionMonitor = true)
        {
            Logger = new Log1(GetType()).Load(Log1.Modes.FILE, true);
            LOCK__AddingNewCacheType = new object();
            LOCK__GivingOutCacheType = new object();
            CachedForms = new Dictionary<Type, CacheFormLoads>();
            FormLoader = pFormLoader;
            InitializeFormsSiliently = ___InitializeFormsSiliently;
            if (pLoaderMode == CacheFormLoaderMode.NONE || !InitializeFormsSiliently)
            {
                if (pFormLoader is object && !pFormLoader.IsDisposed)
                    pFormLoader.Dispose();
                pFormLoader = null;
                if (IsFormLoaderValid())
                    FormLoader.Dispose();
                FormLoader = null;
                pFormLoader = null;
                InitializeFormsSiliently = false;
            }

            ____CacheFormLoaderMode = pLoaderMode;
            LastReadResolution = My.MyProject.Computer.Screen.Bounds;
            if (pRunResolutionMonitor)
            {
                thrResolutionMonitor = new Thread(runResolutionMonitor);
                {
                    var withBlock = thrResolutionMonitor;
                    withBlock.Name = "thrResolutionMonitor";
                    withBlock.IsBackground = true;
                    withBlock.Start();
                }
            }
        }

        /// <summary>
        /// indicate if the show event should be called. NB: A form is created on the thread that calls the show event if silentInitialization is set to true AND it is created on the
        /// thread that is initializing this constructor.
        /// </summary>
        /// <param name="___InitializeFormsSiliently"></param>
        /// <param name="pLoaderMode">Indicates if the Loader if set, will always load forms silently or not</param>
        /// <remarks></remarks>
        public FormsCacheManager_v2(bool ___InitializeFormsSiliently, CacheFormLoaderMode pLoaderMode = CacheFormLoaderMode.NONE) : this(___InitializeFormsSiliently, new SampleInitializedControl(), pLoaderMode)
        {
        }

        // REM Of CacheableForm As {Form, New, ICacheableForm}

        #endregion

        #region Properties

        private Dictionary<Type, CacheFormLoads> CachedForms;
        /// <summary>
        /// Use for indicating a new form is getting cached so don't remove from the list
        /// </summary>
        /// <remarks></remarks>
        private object LOCK__AddingNewCacheType;
        private object LOCK__GivingOutCacheType;
        private Log1 Logger;
        private bool InitializeFormsSiliently;
        private Control FormLoader;
        private CacheFormLoaderMode ____CacheFormLoaderMode;
        private Thread thrResolutionMonitor;
        private Rectangle LastReadResolution;

        public static bool IsDebugMode { get; set; }

        public int CachedFormCount
        {
            get
            {
                return CachedForms.Count;
            }
        }

        public List<object> CachedObjects
        {
            get
            {
                return (from d in CachedForms.Values.ToList()
                        select d.CachedForm).ToList();
            }
        }



        #endregion


        #region Methods


        private void runResolutionMonitor()
        {
            try
            {
                while (!IsDisposed)
                {
                    var cRes = My.MyProject.Computer.Screen.Bounds;
                    if (cRes.Width != LastReadResolution.Width || cRes.Height != LastReadResolution.Height)
                    {
                        if (IsDebugMode)
                            Logger.Print("Resolution Changed. Am recaching forms now");
                        lock (LOCK__GivingOutCacheType)
                        {
                            Interaction.MsgBox(string.Format("The resolution of this device has been changed!!!. {0}Please restart your application to fit to this new resolution.", Environment.NewLine), Constants.vbSystemModal, "Resolution Changed.");
                            break;

                            // 'Dim aCopy As Dictionary(Of String, CacheFormLoads) =
                            // '    New Dictionary(Of String, CacheFormLoads)(Me.CachedForms)
                            // 'Me.CachedForms.Clear()

                            // 'Logger.Print("Me.CachedForms.Count: " & Me.CachedForms.Count.ToString())
                            // 'Logger.Print("aCopy.Count: " & aCopy.Count.ToString())

                            // 'For Each f As KeyValuePair(Of String, CacheFormLoads) In aCopy
                            // '    REM Dim ctl As Control = CType(f.CachedForm, Control)

                            // '    Me.FillUpCacheThreaded(f.Value)



                            // 'Next


                            // 'Try


                            // '    Logger.Print("Am disposing old forms now")
                            // '    REM Clear Received Copy
                            // '    For Each f As CacheFormLoads In aCopy.Values
                            // '        If Not IsNothing(f) AndAlso Not f.CachedForm.IsDisposed Then
                            // '            If TypeOf f.CachedForm Is Control Then
                            // '                Dim fName As String = CType(f.CachedForm, Control).Name
                            // '                If CType(f.CachedForm, Control).IsHandleCreated Then
                            // '                    CType(f.CachedForm, Control).Invoke(Sub() f.CachedForm.Dispose())
                            // '                    Logger.Print("Disposed: " & fName)
                            // '                Else
                            // '                    f.CachedForm = Nothing
                            // '                    Logger.Print("NULLED: " & fName)
                            // '                End If
                            // '            End If
                            // '        End If
                            // '    Next

                            // 'Catch ex As Exception
                            // '    Logger.Print("Error While Disposing:", ex)

                            // 'End Try

                            // 'aCopy.Clear()


                            // ' ''Logger.Print("Me.CachedForms.Count: " & Me.CachedForms.Count.ToString())
                            // ' ''Logger.Print("aCopy.Count: " & aCopy.Count.ToString())

                            // 'aCopy = Nothing



                        }

                        LastReadResolution = cRes;
                        if (IsDebugMode)
                            Logger.Print("Finished recaching forms");
                    }

                    Thread.Sleep(RESOLUTION__CHECKER__THREAD__SLEEP_TIME);
                }
            }
            catch (ThreadAbortException ex)
            {
            }
            catch (Exception ex)
            {
                Logger.Print(ex);
            }

            if (IsDebugMode)
                Logger.Print(Thread.CurrentThread.Name + " is quiting ...");
        }

        public bool AddForm<T>() where T : ICacheableControl, Control, new()
        {
            if (IsDisposed)
                throw new ObjectDisposedException(GetType().Name);
            try
            {
                if (!CachedForms.ContainsKey(typeof(T)))
                {
                    // REM                    Me.FillUpCacheWithFullInitialization(Of T)()
                    FillUpCache<T>();
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exceptions.InvalidCacheFormException();
            }

            return false;
        }


        // 'Public Function AddForm(Of T As ICacheableControl)(obj As T) As Boolean
        // '    Return Me.AddForm(obj, obj.GetType().FullName)
        // 'End Function

        // 'Public Function AddForm(Of T As {Control, New, ICacheableControl})(obj As T) As Boolean
        // '    Return Me.AddForm(Of T)()
        // 'End Function

        // ' ''' <summary>
        // ' ''' Thread Safe. Returns False if Form already exist or an unexpected error occurred
        // ' ''' </summary>
        // ' ''' <returns></returns>
        // ' ''' <remarks></remarks>
        // 'Public Function AddForm(Of T As {Control, New, ICacheableControl})() As Boolean

        // '    If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)


        // '    Try

        // '        If Not Me.CachedForms.ContainsKey(GetType(T)) Then
        // '            Me.FillUpCache(Of T)(InitializeFormsSiliently)
        // '            Return True
        // '        End If

        // '    Catch ex As Exception
        // '        Throw New Exceptions.InvalidCacheFormException()
        // '    End Try

        // '    Return False
        // 'End Function

        /// <summary>
        /// Checks if user will be able to call getForm successfully on this thread
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool canFetchControl<T>() where T : Control, ICacheableControl, new()
        {
            if (!IsFormCached<T>())
                return false;
            var rst = CachedForms[typeof(T)];
            return !rst.IsCaching;
        }

        /// <summary>
        /// gets a cached copy if available. Throws Exception
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <remarks></remarks>
        public T getForm<T>(ApartmentState ThreadAptState = ApartmentState.MTA) where T : Control, ICacheableControl, new()
        {
            if (!IsFormCached<T>())
                throw new Exceptions.NotCachedException(typeof(T).FullName);
            if (InitializeFormsSiliently && IsFormLoaderValid())
            {
                var rst = CachedForms[typeof(T)];
                if (rst.IsCaching)
                {
                    throw new Exception("Please wait, the form you have requested for is still being cached. Try again in 2secs later.");
                }
            }

            lock (LOCK__GivingOutCacheType)
            {
                lock (LOCK__AddingNewCacheType)
                {
                    try
                    {
                        if (____CacheFormLoaderMode == CacheFormLoaderMode.ONCE && InitializeFormsSiliently)
                        {
                            // REM That is ok
                            // REM Dispose and set to false
                            ____CacheFormLoaderMode = CacheFormLoaderMode.NONE;
                            // If Me.IsFormLoaderValid() Then Me.FormLoader.Dispose()
                            // Just set pointer to null
                            FormLoader = null;
                            InitializeFormsSiliently = false;
                        }

                        var rst = CachedForms[typeof(T)];
                        while (rst.IsCaching)
                        {
                            Thread.Sleep(5);
                            Application.DoEvents();
                        }

                        T retVal = (T)rst.CachedForm;
                        rst.IsCaching = true;
                        if (retVal is null)
                            throw new Exceptions.InvalidCodeException(typeof(T).FullName, Thread.CurrentThread.GetApartmentState().ToString());
                        CachedForms[typeof(T)] = rst;
                        var thrFillUpCache = new Thread(() => FillUpCache<T>());
                        thrFillUpCache.IsBackground = true;
                        thrFillUpCache.Name = "Fillup Thread";
                        thrFillUpCache.SetApartmentState(ThreadAptState);
                        thrFillUpCache.Start();
                        return retVal;
                    }
                    catch (Exception ex)
                    {
                        Logger.Print(ex.Message);
                        return null;
                    }
                }
            }
        }


        // ' ''' <summary>
        // ' ''' Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        // ' ''' </summary>
        // ' ''' <remarks></remarks>
        // 'Private Sub FillUpCacheThreaded(ByVal oldControl As CacheFormLoads)
        // '    Me.FillUpCache(oldControl.CachedForm.getNew(), oldControl.TypeFullName)
        // 'End Sub




        /// <summary>
        /// Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <remarks></remarks>
        private void FillUpCache<T>() where T : ICacheableControl, Control, new()
        {
            lock (LOCK__AddingNewCacheType)
            {
                // REM Some controls are not working on cross threading so
                // REM Do atomic Initialize all on the Main thread or initialize none

                if (InitializeFormsSiliently && IsFormLoaderValid())
                {
                    FormLoader.Invoke(new Action(() => TryInitializeThis<T>()));
                }
                else
                {
                    T CachedValue = null;
                    try
                    {
                        CachedValue = new T();
                        CachedValue.ClearControls();
                    }
                    catch (Exception ex)
                    {
                        Logger.Print("FillUpCache()", ex);
                    }

                    if (!IsFormCached<T>())
                    {
                        CachedForms.Add(typeof(T), new CacheFormLoads(CachedValue));
                        if (IsDebugMode)
                            Logger.Print("Just Cached: " + typeof(T).FullName);
                    }
                    else
                    {
                        CachedForms[typeof(T)] = new CacheFormLoads(CachedValue);
                        if (IsDebugMode)
                            Logger.Print("Just REPLACED Cache: " + typeof(T).FullName);
                    }
                }
            }
        }




        // ' ''' <summary>
        // ' ''' Only used for initial loading on splash screen
        // ' ''' </summary>
        // ' ''' <typeparam name="T"></typeparam>
        // ' ''' <param name="CallShowEvent"></param>
        // ' ''' <remarks></remarks>
        // 'Private Sub FillUpCacheWithFullInitialization(Of T As {New, ICacheableControl, Control})()
        // '    Me.FillUpCache(Of T)()




        // '    If Me.IsFormCached(Of T)() Then

        // '        REM A form is created on the thread that calls it show() method
        // '        REM Also calling show will call Form_Load event which is not good
        // '        Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)  REM CType(Me.CachedForms(GetType(T)).CachedForm, T)
        // '        Dim frm As Form = TryCast(CachedValue, Form)
        // '        If frm IsNot Nothing Then

        // '            Dim IniOpacity As Double = frm.Opacity
        // '            frm.Opacity = 0
        // '            frm.Show()
        // '            frm.Hide()
        // '            frm.Opacity = IniOpacity

        // '        End If

        // '    End If
        // 'End Sub



        private void TryInitializeThis<T>() where T : ICacheableControl, Control, new()
        {
            T CachedValue = null;
            try
            {
                CachedValue = new T();
                CachedValue.ClearControls();
            }
            catch (Exception ex)
            {
                Logger.Print("FillUpCache()", ex);
            }

            if (!IsFormCached<T>())
            {
                CachedForms.Add(typeof(T), new CacheFormLoads(CachedValue, true));
                if (IsDebugMode)
                    Logger.Print("Just Cached: " + typeof(T).FullName);
            }
            else
            {
                CachedForms[typeof(T)] = new CacheFormLoads(CachedValue, true);
                if (IsDebugMode)
                    Logger.Print("Just REPLACED Cache: " + typeof(T).FullName);
            }

            if (IsFormCached<T>())
            {

                // REM A form is created on the thread that calls it show() method
                // REM Also calling show will call Form_Load event which is not good
                // REM Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)  REM CType(Me.CachedForms(GetType(T)).CachedForm, T)
                Form frm = CachedValue as Form;
                if (frm is object)
                {
                    double IniOpacity = frm.Opacity;
                    frm.Opacity = 0d;
                    frm.Show();
                    frm.Hide();
                    frm.Opacity = IniOpacity;
                }
            }

            var pCac = CachedForms[typeof(T)];
            pCac.IsCaching = false;
            CachedForms[typeof(T)] = pCac;
            Debug.Print("I just Initialized stuff on Main Thread.: " + Objects.EStrings.valueOf(Thread.CurrentThread.Name));
        }



        // ' ''' <summary>
        // ' ''' Only used for initial loading on splash screen
        // ' ''' </summary>
        // ' ''' <typeparam name="T"></typeparam>
        // ' ''' <param name="CallShowEvent"></param>
        // ' ''' <remarks></remarks>
        // 'Private Sub FillUpCache(Of T As {Control, New, ICacheableControl})(ByVal CallShowEvent As Boolean)
        // '    Me.FillUpCache(Of T)()
        // '    If Not CallShowEvent Then Return

        // '    If Me.IsFormCached(Of T)() Then

        // '        REM A form is created on the thread that calls it show() method
        // '        REM Also calling show will call Form_Load event which is not good
        // '        Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)

        // '        If TryCast(CachedValue, Form) IsNot Nothing Then

        // '            Dim IniOpacity As Double = TryCast(CachedValue, Form).Opacity
        // '            TryCast(CachedValue, Form).Opacity = 0
        // '            CachedValue.Show()
        // '            CachedValue.Hide()
        // '            TryCast(CachedValue, Form).Opacity = IniOpacity

        // '        End If

        // '    End If
        // 'End Sub

        // 'Public Function IsFormCached(Of T As {ICacheableControl})() As Boolean
        // '    If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)

        // '    Return Me.CachedForms.ContainsKey(GetType(T).FullName)

        // 'End Function

        public bool IsFormCached<T>() where T : ICacheableControl, Control, new()
        {
            if (IsDisposed)
                throw new ObjectDisposedException(typeof(T).FullName);
            return CachedForms.ContainsKey(typeof(T));
        }

        private bool IsFormLoaderValid()
        {
            return FormLoader is object && !FormLoader.IsDisposed && FormLoader.IsHandleCreated;
        }

        #endregion

        #region Enums and Consts

        private const int RESOLUTION__CHECKER__THREAD__SLEEP_TIME = 1000;

        public enum CacheFormLoaderMode
        {
            NONE,
            ONCE,
            ALWAYS
        }
        #endregion

        #region Structures


        private struct CacheFormLoads
        {
            public CacheFormLoads(object f) : this(f, false)
            {
            }

            public CacheFormLoads(object f, bool pIsCaching)
            {
                CachedForm = f;
                IsCaching = pIsCaching;
            }

            public object CachedForm { get; set; }
            public bool IsCaching { get; set; }
        }
        #endregion


        #region IDisposable Support
        private bool disposedValue; // To detect redundant calls

        /// <summary>
        /// Indicates if this class has been disposed
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
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

                    try
                    {

                        // REM If user call this method directly, it will call it on the user thread.
                        // REM 

                        // REM You shouldn't be disposing forms while adding forms at the same time
                        lock (LOCK__GivingOutCacheType)
                        {
                            lock (LOCK__AddingNewCacheType)
                            {
                                // Try Loading on both threads
                                foreach (CacheFormLoads f in CachedForms.Values)
                                {
                                    if (!Information.IsNothing(f) && !((Control)f.CachedForm).IsDisposed && ((Control)f.CachedForm).IsHandleCreated)
                                    {
                                        if (IsFormLoaderValid())
                                        {
                                            if (IsDebugMode)
                                                Logger.Print("Am calling with the loader thread");
                                            if (IsDebugMode)
                                                Logger.Print("If you have cross thread error here, it means you loaded the forms here with handle on cache thread" + " not app thread and you are trying to close on app thread");
                                            // Possibilities, the background thread that created the loader is not accessible again
                                            // First solution is passing in the COORDINATOR Into the Initialization of the Cachmanager
                                            // Second solution is invoking the TryFormCaching under the main thread
                                            FormLoader.Invoke(new Action(() => ((IDisposable)f.CachedForm).Dispose()));
                                        }
                                        else
                                        {
                                            if (IsDebugMode)
                                                Logger.Print("Am calling with thread calling this method");
                                            ((IDisposable)f.CachedForm).Dispose();
                                        }
                                    }
                                }

                                CachedForms.Clear();
                                CachedForms = null;
                            }
                        }

                        disposedValue = true;
                        if (thrResolutionMonitor is object && thrResolutionMonitor.IsAlive)
                            thrResolutionMonitor.Abort();
                        thrResolutionMonitor = null;
                    }
                    catch (Exception ex)
                    {
                        Logger.Print("Disposing FormCacheManager: ", ex);
                    }
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
        /// NOTE: this will dispose your initialization form if available
        /// </summary>
        /// <remarks></remarks>
        public void Dispose()
        {
            // Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            DisposeAllUnmanagedResources(true);
        }

        /// <summary>
        /// You can pass in false if you don't want to dispose your passed in variables like cache form loader. Thread Safe
        /// </summary>
        /// <param name="pValDisposePassedInVariables"></param>
        /// <remarks></remarks>
        public void DisposeAllUnmanagedResources(bool pValDisposePassedInVariables)
        {
            Dispose(true);
            if (pValDisposePassedInVariables)
            {
                if (IsFormLoaderValid())
                    FormLoader.Dispose();
                FormLoader = null;
            }

            GC.SuppressFinalize(this);
        }
        #endregion

    }
}