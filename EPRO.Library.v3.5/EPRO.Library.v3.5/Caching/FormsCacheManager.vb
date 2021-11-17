Imports CODERiT.Logger.v._3._5



Imports System.Windows.Forms
Imports System.Threading
Imports System.Drawing

Namespace Caching
    Public Class FormsCacheManager
        Implements IDisposable


#Region "Constructors"

        ''' <summary>
        ''' indicate if the show event should be called. NB: A form is created on the thread that calls the show event
        ''' </summary>
        ''' <param name="___InitializeFormsSiliently"></param>
        ''' <remarks></remarks>
        Sub New(___InitializeFormsSiliently As Boolean)
            Me.Logger = New Log1(Me.GetType()).Load(Log1.Modes.FILE, True)
            Me.LOCK__AddingNewCacheType = New Object()
            Me.LOCK__GivingOutCacheType = New Object()
            Me.CachedForms = New Dictionary(Of String, CacheFormLoads)()
            Me.InitializeFormsSiliently = ___InitializeFormsSiliently
            Me.LastReadResolution = My.Computer.Screen.Bounds

            Me.thrResolutionMonitor = New Thread(AddressOf Me.runResolutionMonitor)
            With Me.thrResolutionMonitor
                .Name = "thrResolutionMonitor"
                .IsBackground = True
                .Start()

            End With
        End Sub

        REM Of CacheableForm As {Form, New, ICacheableForm}

#End Region

#Region "Properties"

        Private CachedForms As Dictionary(Of String, CacheFormLoads)
        ''' <summary>
        ''' Use for indicating a new form is getting cached so don't remove from the list
        ''' </summary>
        ''' <remarks></remarks>
        Private LOCK__AddingNewCacheType As Object
        Private LOCK__GivingOutCacheType As Object

        Private Logger As Log1

        Private InitializeFormsSiliently As Boolean

        Private Const RESOLUTION__CHECKER__THREAD__SLEEP_TIME As Int32 = 1000

        Private thrResolutionMonitor As Thread
        Private LastReadResolution As Rectangle

#End Region


#Region "Methods"


        Private Sub runResolutionMonitor()
            Try
                While Not Me.IsDisposed

                    Dim cRes As Rectangle = My.Computer.Screen.Bounds
                    If cRes.Width <> Me.LastReadResolution.Width OrElse cRes.Height <> Me.LastReadResolution.Height Then

                        Logger.Print("Resolution Changed. Am recaching forms now")

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

                        Logger.Print("Finished recaching forms")
                    End If


                    Thread.Sleep(RESOLUTION__CHECKER__THREAD__SLEEP_TIME)

                End While
            Catch ex As ThreadAbortException
            Catch ex As Exception
                Logger.Print(ex)
            End Try

            Logger.Print(Thread.CurrentThread.Name & " is quiting ...")
        End Sub




        Public Function AddForm(Of T As ICacheableControl)(obj As T, ByVal CacheAsTypeFullName As String) As Boolean
            If Me.IsDisposed Then Throw New ObjectDisposedException(Me.GetType().Name)


            Try

                If Not Me.CachedForms.ContainsKey(CacheAsTypeFullName) Then
                    Me.FillUpCache(Of T)(obj, InitializeFormsSiliently, CacheAsTypeFullName)
                    Return True
                End If

            Catch ex As Exception
                Throw New Exceptions.InvalidCacheFormException()
            End Try

            Return False
        End Function


        Public Function AddForm(Of T As ICacheableControl)(obj As T) As Boolean
            Return Me.AddForm(obj, obj.GetType().FullName)
        End Function

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
        ''' gets a cached copy if available. Throws Exception
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getForm(Of T As {ICacheableControl})(Optional ThreadAptState As Threading.ApartmentState = Threading.ApartmentState.MTA) As T

            SyncLock LOCK__GivingOutCacheType
                SyncLock LOCK__AddingNewCacheType

                    If Not Me.IsFormCached(GetType(T).FullName) Then Throw New Exceptions.NotCachedException(GetType(T).FullName)

                    Try


                        Dim rst As CacheFormLoads = CType(Me.CachedForms(GetType(T).FullName), Global.EPRO.Library.v3._5.Caching.FormsCacheManager.CacheFormLoads)
                        Do While rst.IsCaching
                            Threading.Thread.Sleep(5)
                            Application.DoEvents()
                        Loop

                        Dim retVal As T = CType(rst.CachedForm, T)

                        rst.IsCaching = True

                        If retVal Is Nothing Then Throw New Exceptions.InvalidCodeException(GetType(T).FullName, Threading.Thread.CurrentThread.GetApartmentState().ToString())

                        Dim thrFillUpCache As New Threading.Thread(Sub() Me.FillUpCacheThreaded(rst))
                        thrFillUpCache.IsBackground = True
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


        ''' <summary>
        ''' Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub FillUpCacheThreaded(ByVal oldControl As CacheFormLoads)
            Me.FillUpCache(oldControl.CachedForm.getNew(), oldControl.TypeFullName)
        End Sub




        ''' <summary>
        ''' Thread Safe. Only One thread is allow to add a form at a time and disposing is blocked while adding or setting
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <remarks></remarks>
        Private Sub FillUpCache(Of T As {ICacheableControl})(ByVal obj As T,
                                                             ByVal CacheAsTypeFullName As String
                                                             )
            SyncLock LOCK__AddingNewCacheType
                Dim CachedValue As T = Nothing

                Try

                    ''CachedValue = New T()
                    CachedValue = obj
                    CachedValue.ClearControls()

                Catch ex As Exception
                    Logger.Print("FillUpCache()", ex)
                End Try

                If Not IsFormCached(CacheAsTypeFullName) Then
                    Me.CachedForms.Add(CacheAsTypeFullName, New CacheFormLoads(CachedValue, CacheAsTypeFullName))
                    Logger.Print("Just Cached: " & CacheAsTypeFullName)
                Else
                    Me.CachedForms(CacheAsTypeFullName) = New CacheFormLoads(CachedValue, CacheAsTypeFullName)
                    Logger.Print("Just REPLACED Cache: " & CacheAsTypeFullName)
                End If



            End SyncLock
        End Sub




        ''' <summary>
        ''' Only used for initial loading on splash screen
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="CallShowEvent"></param>
        ''' <remarks></remarks>
        Private Sub FillUpCache(Of T As {ICacheableControl})(ByVal obj As T, ByVal CallShowEvent As Boolean,
                                                              ByVal CacheAsTypeFullName As String)
            Me.FillUpCache(Of T)(obj, CacheAsTypeFullName)
            If Not CallShowEvent Then Return

            If Me.IsFormCached(CacheAsTypeFullName) Then

                REM A form is created on the thread that calls it show() method
                REM Also calling show will call Form_Load event which is not good
                Dim CachedValue As T = CType(Me.CachedForms(CacheAsTypeFullName).CachedForm, T)  REM CType(Me.CachedForms(GetType(T)).CachedForm, T)
                Dim frm As Form = TryCast(CachedValue, Form)
                If frm IsNot Nothing Then

                    Dim IniOpacity As Double = frm.Opacity
                    frm.Opacity = 0
                    frm.Show()
                    frm.Hide()
                    frm.Opacity = IniOpacity

                End If

            End If
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

        Public Function IsFormCached(CacheTypeFullName As String) As Boolean
            If Me.IsDisposed Then Throw New ObjectDisposedException(CacheTypeFullName)

            Return Me.CachedForms.ContainsKey(CacheTypeFullName)

        End Function


#End Region


#Region "Structures"


        Private Structure CacheFormLoads
            Sub New(ByVal f As ICacheableControl, ByVal CacheAsTypeFullName As String)
                Me.CachedForm = f
                Me.IsCaching = False
                Me.__TypeFullName = CacheAsTypeFullName
            End Sub

            Public Property CachedForm As ICacheableControl

            Private __TypeFullName As String
            Public ReadOnly Property TypeFullName As String
                Get
                    Return Me.__TypeFullName
                End Get
            End Property

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

                        REM You shouldn't be disposing forms while adding forms at the same time
                        SyncLock LOCK__GivingOutCacheType
                            SyncLock LOCK__AddingNewCacheType

                                For Each f As CacheFormLoads In Me.CachedForms.Values
                                    If Not IsNothing(f) AndAlso Not CType(f.CachedForm, Form).IsDisposed Then CType(f.CachedForm, IDisposable).Dispose()
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
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
#End Region

    End Class

End Namespace