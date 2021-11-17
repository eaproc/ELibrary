Imports CODERiT.Logger.v._3._5



Imports System.Windows.Forms
Imports System.Threading
Imports System.Drawing

Namespace Caching
    Public Class FormsCacheManager_v2
        Implements IDisposable


#Region "Constructors"

        ''' <summary>
        ''' 'Careat a cach manager with your own initializer if you want
        ''' </summary>
        ''' <param name="___InitializeFormsSiliently"></param>
        ''' <param name="pFormLoader"></param>
        ''' <param name="pRunResolutionMonitor"></param>
        ''' <remarks></remarks>
        Sub New(___InitializeFormsSiliently As Boolean,
                pFormLoader As Control,
                pLoaderMode As CacheFormLoaderMode,
                Optional pRunResolutionMonitor As Boolean = True
                                                            )
            Me.Logger = New Log1(Me.GetType()).Load(Log1.Modes.FILE, True)
            Me.LOCK__AddingNewCacheType = New Object()
            Me.LOCK__GivingOutCacheType = New Object()
            Me.CachedForms = New Dictionary(Of Type, CacheFormLoads)()
            Me.FormLoader = pFormLoader
            Me.InitializeFormsSiliently = ___InitializeFormsSiliently

            If pLoaderMode = CacheFormLoaderMode.NONE OrElse Not InitializeFormsSiliently Then
                If pFormLoader IsNot Nothing AndAlso Not pFormLoader.IsDisposed Then pFormLoader.Dispose()
                pFormLoader = Nothing
                If Me.IsFormLoaderValid() Then Me.FormLoader.Dispose()
                Me.FormLoader = Nothing
                pFormLoader = Nothing
                Me.InitializeFormsSiliently = False
            End If

            Me.____CacheFormLoaderMode = pLoaderMode


            Me.LastReadResolution = My.Computer.Screen.Bounds


            If pRunResolutionMonitor Then
                Me.thrResolutionMonitor = New Thread(AddressOf Me.runResolutionMonitor)
                With Me.thrResolutionMonitor
                    .Name = "thrResolutionMonitor"
                    .IsBackground = True
                    .Start()

                End With
            End If

        End Sub

        ''' <summary>
        ''' indicate if the show event should be called. NB: A form is created on the thread that calls the show event if silentInitialization is set to true AND it is created on the
        ''' thread that is initializing this constructor.
        ''' </summary>
        ''' <param name="___InitializeFormsSiliently"></param>
        ''' <param name="pLoaderMode">Indicates if the Loader if set, will always load forms silently or not</param>
        ''' <remarks></remarks>
        Sub New(___InitializeFormsSiliently As Boolean, Optional pLoaderMode As CacheFormLoaderMode = CacheFormLoaderMode.NONE)
            Me.New(___InitializeFormsSiliently, New SampleInitializedControl(), pLoaderMode)
        End Sub

        REM Of CacheableForm As {Form, New, ICacheableForm}

#End Region

#Region "Properties"

        Private CachedForms As Dictionary(Of Type, CacheFormLoads)
        ''' <summary>
        ''' Use for indicating a new form is getting cached so don't remove from the list
        ''' </summary>
        ''' <remarks></remarks>
        Private LOCK__AddingNewCacheType As Object
        Private LOCK__GivingOutCacheType As Object

        Private Logger As Log1

        Private InitializeFormsSiliently As Boolean
        Private FormLoader As Control

        Private ____CacheFormLoaderMode As CacheFormLoaderMode

        Private thrResolutionMonitor As Thread
        Private LastReadResolution As Rectangle
        Public Shared Property IsDebugMode As Boolean



        Public ReadOnly Property CachedFormCount As Int32
            Get
                Return Me.CachedForms.Count
            End Get
        End Property

        Public ReadOnly Property CachedObjects As List(Of Object)
            Get
                Return (
                    From d In Me.CachedForms.Values.ToList()
                    Select d.CachedForm
                    ).ToList()
            End Get
        End Property



#End Region


#Region "Methods"


        Private Sub runResolutionMonitor()
            Try
                While Not Me.IsDisposed

                    Dim cRes As Rectangle = My.Computer.Screen.Bounds
                    If cRes.Width <> Me.LastReadResolution.Width OrElse cRes.Height <> Me.LastReadResolution.Height Then

                        If IsDebugMode Then Logger.Print("Resolution Changed. Am recaching forms now")

                        SyncLock LOCK__GivingOutCacheType

                            MsgBox(
                                String.Format(
                                    "The resolution of this device has been changed!!!. {0}Please restart your application to fit to this new resolution.",
                                    Environment.NewLine()
                                    ), vbSystemModal, "Resolution Changed."
                                )
                            Exit While

                            ''Dim aCopy As Dictionary(Of String, CacheFormLoads) =
                            ''    New Dictionary(Of String, CacheFormLoads)(Me.CachedForms)
                            ''Me.CachedForms.Clear()

                            ''Logger.Print("Me.CachedForms.Count: " & Me.CachedForms.Count.ToString())
                            ''Logger.Print("aCopy.Count: " & aCopy.Count.ToString())

                            ''For Each f As KeyValuePair(Of String, CacheFormLoads) In aCopy
                            ''    REM Dim ctl As Control = CType(f.CachedForm, Control)

                            ''    Me.FillUpCacheThreaded(f.Value)



                            ''Next


                            ''Try


                            ''    Logger.Print("Am disposing old forms now")
                            ''    REM Clear Received Copy
                            ''    For Each f As CacheFormLoads In aCopy.Values
                            ''        If Not IsNothing(f) AndAlso Not f.CachedForm.IsDisposed Then
                            ''            If TypeOf f.CachedForm Is Control Then
                            ''                Dim fName As String = CType(f.CachedForm, Control).Name
                            ''                If CType(f.CachedForm, Control).IsHandleCreated Then
                            ''                    CType(f.CachedForm, Control).Invoke(Sub() f.CachedForm.Dispose())
                            ''                    Logger.Print("Disposed: " & fName)
                            ''                Else
                            ''                    f.CachedForm = Nothing
                            ''                    Logger.Print("NULLED: " & fName)
                            ''                End If
                            ''            End If
                            ''        End If
                            ''    Next

                            ''Catch ex As Exception
                            ''    Logger.Print("Error While Disposing:", ex)

                            ''End Try

                            ''aCopy.Clear()


                            '' ''Logger.Print("Me.CachedForms.Count: " & Me.CachedForms.Count.ToString())
                            '' ''Logger.Print("aCopy.Count: " & aCopy.Count.ToString())

                            ''aCopy = Nothing



                        End SyncLock

                        Me.LastReadResolution = cRes

                        If IsDebugMode Then Logger.Print("Finished recaching forms")
                    End If


                    Thread.Sleep(RESOLUTION__CHECKER__THREAD__SLEEP_TIME)

                End While
            Catch ex As ThreadAbortException
            Catch ex As Exception
                Logger.Print(ex)
            End Try

            If IsDebugMode Then Logger.Print(Thread.CurrentThread.Name & " is quiting ...")
        End Sub




        Public Function AddForm(Of T As {New, ICacheableControl, Control})() As Boolean
            If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)


            Try

                If Not Me.CachedForms.ContainsKey(GetType(T)) Then
                    REM                    Me.FillUpCacheWithFullInitialization(Of T)()
                    Me.FillUpCache(Of T)()
                    Return True
                End If

            Catch ex As Exception
                Throw New Exceptions.InvalidCacheFormException()
            End Try

            Return False
        End Function


        ''Public Function AddForm(Of T As ICacheableControl)(obj As T) As Boolean
        ''    Return Me.AddForm(obj, obj.GetType().FullName)
        ''End Function

        ''Public Function AddForm(Of T As {Control, New, ICacheableControl})(obj As T) As Boolean
        ''    Return Me.AddForm(Of T)()
        ''End Function

        '' ''' <summary>
        '' ''' Thread Safe. Returns False if Form already exist or an unexpected error occurred
        '' ''' </summary>
        '' ''' <returns></returns>
        '' ''' <remarks></remarks>
        ''Public Function AddForm(Of T As {Control, New, ICacheableControl})() As Boolean

        ''    If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)


        ''    Try

        ''        If Not Me.CachedForms.ContainsKey(GetType(T)) Then
        ''            Me.FillUpCache(Of T)(InitializeFormsSiliently)
        ''            Return True
        ''        End If

        ''    Catch ex As Exception
        ''        Throw New Exceptions.InvalidCacheFormException()
        ''    End Try

        ''    Return False
        ''End Function

        ''' <summary>
        ''' Checks if user will be able to call getForm successfully on this thread
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function canFetchControl(Of T As {New, Control, ICacheableControl})() As Boolean
            If Not Me.IsFormCached(Of T)() Then Return False

            Dim rst As CacheFormLoads = Me.CachedForms(GetType(T))
            Return Not rst.IsCaching
        End Function

        ''' <summary>
        ''' gets a cached copy if available. Throws Exception
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getForm(Of T As {New, Control, ICacheableControl})(Optional ThreadAptState As Threading.ApartmentState = Threading.ApartmentState.MTA) As T
            If Not Me.IsFormCached(Of T)() Then Throw New Exceptions.NotCachedException(GetType(T).FullName)

            If Me.InitializeFormsSiliently AndAlso Me.IsFormLoaderValid Then
                Dim rst As CacheFormLoads = Me.CachedForms(GetType(T))
                If rst.IsCaching Then
                    Throw New Exception("Please wait, the form you have requested for is still being cached. Try again in 2secs later.")
                End If
            End If


            SyncLock LOCK__GivingOutCacheType
                SyncLock LOCK__AddingNewCacheType

                    Try


                        If Me.____CacheFormLoaderMode = CacheFormLoaderMode.ONCE AndAlso Me.InitializeFormsSiliently Then
                            REM That is ok
                            REM Dispose and set to false
                            Me.____CacheFormLoaderMode = CacheFormLoaderMode.NONE
                            ' If Me.IsFormLoaderValid() Then Me.FormLoader.Dispose()
                            ' Just set pointer to null
                            Me.FormLoader = Nothing
                            Me.InitializeFormsSiliently = False

                        End If




                        Dim rst As CacheFormLoads = Me.CachedForms(GetType(T))
                        Do While rst.IsCaching
                            Threading.Thread.Sleep(5)
                            Application.DoEvents()
                        Loop

                        Dim retVal As T = CType(rst.CachedForm, T)

                        rst.IsCaching = True

                        If retVal Is Nothing Then Throw New Exceptions.InvalidCodeException(GetType(T).FullName, Threading.Thread.CurrentThread.GetApartmentState().ToString())

                        Me.CachedForms(GetType(T)) = rst

                        Dim thrFillUpCache As New Threading.Thread(Sub() Me.FillUpCache(Of T)())
                        thrFillUpCache.IsBackground = True
                        thrFillUpCache.Name = "Fillup Thread"
                        thrFillUpCache.SetApartmentState(ThreadAptState)
                        thrFillUpCache.Start()
                        Return retVal

                    Catch ex As Exception
                        Logger.Print(ex.Message)
                        Return Nothing
                    End Try

                End SyncLock

            End SyncLock

        End Function


        '' ''' <summary>
        '' ''' Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        '' ''' </summary>
        '' ''' <remarks></remarks>
        ''Private Sub FillUpCacheThreaded(ByVal oldControl As CacheFormLoads)
        ''    Me.FillUpCache(oldControl.CachedForm.getNew(), oldControl.TypeFullName)
        ''End Sub




        ''' <summary>
        ''' Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <remarks></remarks>
        Private Sub FillUpCache(Of T As {New, ICacheableControl, Control})()
            SyncLock LOCK__AddingNewCacheType
                REM Some controls are not working on cross threading so
                REM Do atomic Initialize all on the Main thread or initialize none

                If Me.InitializeFormsSiliently AndAlso Me.IsFormLoaderValid Then
                    Me.FormLoader.Invoke(Sub() Me.TryInitializeThis(Of T)())

                Else

                    Dim CachedValue As T = Nothing

                    Try

                        CachedValue = New T()
                        CachedValue.ClearControls()

                    Catch ex As Exception
                        Logger.Print("FillUpCache()", ex)
                    End Try

                    If Not IsFormCached(Of T)() Then
                        Me.CachedForms.Add(GetType(T), New CacheFormLoads(CachedValue))
                        If IsDebugMode Then Logger.Print("Just Cached: " & GetType(T).FullName)
                    Else
                        Me.CachedForms(GetType(T)) = New CacheFormLoads(CachedValue)
                        If IsDebugMode Then Logger.Print("Just REPLACED Cache: " & GetType(T).FullName)
                    End If

                End If

            End SyncLock
        End Sub




        '' ''' <summary>
        '' ''' Only used for initial loading on splash screen
        '' ''' </summary>
        '' ''' <typeparam name="T"></typeparam>
        '' ''' <param name="CallShowEvent"></param>
        '' ''' <remarks></remarks>
        ''Private Sub FillUpCacheWithFullInitialization(Of T As {New, ICacheableControl, Control})()
        ''    Me.FillUpCache(Of T)()




        ''    If Me.IsFormCached(Of T)() Then

        ''        REM A form is created on the thread that calls it show() method
        ''        REM Also calling show will call Form_Load event which is not good
        ''        Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)  REM CType(Me.CachedForms(GetType(T)).CachedForm, T)
        ''        Dim frm As Form = TryCast(CachedValue, Form)
        ''        If frm IsNot Nothing Then

        ''            Dim IniOpacity As Double = frm.Opacity
        ''            frm.Opacity = 0
        ''            frm.Show()
        ''            frm.Hide()
        ''            frm.Opacity = IniOpacity

        ''        End If

        ''    End If
        ''End Sub



        Private Sub TryInitializeThis(Of T As {New, ICacheableControl, Control})()

            Dim CachedValue As T = Nothing

            Try

                CachedValue = New T()
                CachedValue.ClearControls()

            Catch ex As Exception
                Logger.Print("FillUpCache()", ex)
            End Try

            If Not IsFormCached(Of T)() Then
                Me.CachedForms.Add(GetType(T), New CacheFormLoads(CachedValue, True))
                If IsDebugMode Then Logger.Print("Just Cached: " & GetType(T).FullName)
            Else
                Me.CachedForms(GetType(T)) = New CacheFormLoads(CachedValue, True)
                If IsDebugMode Then Logger.Print("Just REPLACED Cache: " & GetType(T).FullName)
            End If


            If Me.IsFormCached(Of T)() Then

                REM A form is created on the thread that calls it show() method
                REM Also calling show will call Form_Load event which is not good
                REM Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)  REM CType(Me.CachedForms(GetType(T)).CachedForm, T)
                Dim frm As Form = TryCast(CachedValue, Form)
                If frm IsNot Nothing Then

                    Dim IniOpacity As Double = frm.Opacity
                    frm.Opacity = 0
                    frm.Show()
                    frm.Hide()
                    frm.Opacity = IniOpacity

                End If

            End If

            Dim pCac As CacheFormLoads = Me.CachedForms(GetType(T))
            pCac.IsCaching = False
            Me.CachedForms(GetType(T)) = pCac
            Debug.Print("I just Initialized stuff on Main Thread.: " & Objects.EStrings.valueOf(Thread.CurrentThread.Name))

        End Sub



        '' ''' <summary>
        '' ''' Only used for initial loading on splash screen
        '' ''' </summary>
        '' ''' <typeparam name="T"></typeparam>
        '' ''' <param name="CallShowEvent"></param>
        '' ''' <remarks></remarks>
        ''Private Sub FillUpCache(Of T As {Control, New, ICacheableControl})(ByVal CallShowEvent As Boolean)
        ''    Me.FillUpCache(Of T)()
        ''    If Not CallShowEvent Then Return

        ''    If Me.IsFormCached(Of T)() Then

        ''        REM A form is created on the thread that calls it show() method
        ''        REM Also calling show will call Form_Load event which is not good
        ''        Dim CachedValue As T = CType(Me.CachedForms(GetType(T)).CachedForm, T)

        ''        If TryCast(CachedValue, Form) IsNot Nothing Then

        ''            Dim IniOpacity As Double = TryCast(CachedValue, Form).Opacity
        ''            TryCast(CachedValue, Form).Opacity = 0
        ''            CachedValue.Show()
        ''            CachedValue.Hide()
        ''            TryCast(CachedValue, Form).Opacity = IniOpacity

        ''        End If

        ''    End If
        ''End Sub

        ''Public Function IsFormCached(Of T As {ICacheableControl})() As Boolean
        ''    If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)

        ''    Return Me.CachedForms.ContainsKey(GetType(T).FullName)

        ''End Function

        Public Function IsFormCached(Of T As {New, ICacheableControl, Control})() As Boolean
            If Me.IsDisposed Then Throw New ObjectDisposedException(GetType(T).FullName)

            Return Me.CachedForms.ContainsKey(GetType(T))

        End Function


        Private Function IsFormLoaderValid() As Boolean
            Return Me.FormLoader IsNot Nothing AndAlso Not Me.FormLoader.IsDisposed AndAlso Me.FormLoader.IsHandleCreated
        End Function

#End Region

#Region "Enums and Consts"

        Private Const RESOLUTION__CHECKER__THREAD__SLEEP_TIME As Int32 = 1000

        Public Enum CacheFormLoaderMode
            NONE
            ONCE
            ALWAYS
        End Enum
#End Region

#Region "Structures"


        Private Structure CacheFormLoads
            Sub New(ByVal f As Object)
                Me.New(f, False)
            End Sub

            Sub New(ByVal f As Object, pIsCaching As Boolean)
                Me.CachedForm = f
                Me.IsCaching = pIsCaching

            End Sub


            Public Property CachedForm As Object


            Public Property IsCaching As Boolean


        End Structure
#End Region


#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ''' <summary>
        ''' Indicates if this class has been disposed
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
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

                    Try

                        REM If user call this method directly, it will call it on the user thread.
                        REM 

                        REM You shouldn't be disposing forms while adding forms at the same time
                        SyncLock LOCK__GivingOutCacheType
                            SyncLock LOCK__AddingNewCacheType
                                '   Try Loading on both threads
                                For Each f As CacheFormLoads In Me.CachedForms.Values
                                    If Not IsNothing(f) AndAlso Not CType(f.CachedForm, Control).IsDisposed AndAlso CType(f.CachedForm, Control).IsHandleCreated Then
                                        If Me.IsFormLoaderValid() Then
                                            If IsDebugMode Then Logger.Print("Am calling with the loader thread")
                                            If IsDebugMode Then Logger.Print("If you have cross thread error here, it means you loaded the forms here with handle on cache thread" &
                                                         " not app thread and you are trying to close on app thread")
                                            '   Possibilities, the background thread that created the loader is not accessible again
                                            '   First solution is passing in the COORDINATOR Into the Initialization of the Cachmanager
                                            '   Second solution is invoking the TryFormCaching under the main thread
                                            Me.FormLoader.Invoke(Sub() CType(f.CachedForm, IDisposable).Dispose())
                                        Else
                                            If IsDebugMode Then Logger.Print("Am calling with thread calling this method")
                                            CType(f.CachedForm, IDisposable).Dispose()
                                        End If

                                    End If

                                Next

                                Me.CachedForms.Clear()
                                Me.CachedForms = Nothing

                            End SyncLock
                        End SyncLock


                        Me.disposedValue = True
                        If Me.thrResolutionMonitor IsNot Nothing AndAlso Me.thrResolutionMonitor.IsAlive Then Me.thrResolutionMonitor.Abort()
                        Me.thrResolutionMonitor = Nothing

                    Catch ex As Exception
                        Logger.Print("Disposing FormCacheManager: ", ex)
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
        ''' <summary>
        ''' NOTE: this will dispose your initialization form if available
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Me.DisposeAllUnmanagedResources(True)
        End Sub

        ''' <summary>
        ''' You can pass in false if you don't want to dispose your passed in variables like cache form loader. Thread Safe
        ''' </summary>
        ''' <param name="pValDisposePassedInVariables"></param>
        ''' <remarks></remarks>
        Public Sub DisposeAllUnmanagedResources(pValDisposePassedInVariables As Boolean)
            Me.Dispose(True)

            If pValDisposePassedInVariables Then
                If Me.IsFormLoaderValid Then Me.FormLoader.Dispose()
                Me.FormLoader = Nothing
            End If

            GC.SuppressFinalize(Me)

        End Sub
#End Region

    End Class

End Namespace