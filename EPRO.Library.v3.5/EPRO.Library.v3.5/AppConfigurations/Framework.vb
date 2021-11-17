Option Strict On

Imports EPRO.Library.v3._5.Objects.EForm
Imports EPRO.Library.v3._5.Modules

Imports System.Windows.Forms
Imports CODERiT.Logger.v._3._5.Exceptions
Imports System.Runtime.InteropServices
Imports System.Security.Principal

Namespace AppConfigurations
    Public NotInheritable Class Framework


#Region "Showing Window"
        Private Enum ShowWindowEnum
            Hide = 0
            ShowNormal = 1
            ShowMinimized = 2
            ShowMaximized = 3
            Maximize = 3
            ShowNormalNoActivate = 4
            Show = 5
            Minimize = 6
            ShowMinNoActivate = 7
            ShowNoActivate = 8
            Restore = 9
            ShowDefault = 10
            ForceMinimized = 11
        End Enum

        <System.Runtime.InteropServices.DllImport("user32.dll")> _
        Private Shared Function ShowWindow(hWnd As IntPtr, flags As ShowWindowEnum) As <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)> Boolean
        End Function


        Private Structure Windowplacement
            Public length As Integer
            Public flags As Integer
            Public showCmd As Integer
            Public ptMinPosition As System.Drawing.Point
            Public ptMaxPosition As System.Drawing.Point
            Public rcNormalPosition As System.Drawing.Rectangle
        End Structure

        <DllImport("user32.dll")> _
        Private Shared Function GetWindowPlacement(hWnd As IntPtr, ByRef lpwndpl As Windowplacement) As <MarshalAs(UnmanagedType.Bool)> Boolean
        End Function


#End Region





#Region "Constructors"


        Sub New(ByVal StartUp_Form As Form,
               Optional ByVal ShutDownMode As ShutdownModes = ShutdownModes.CLOSE_WHEN_LAST_FORM_CLOSES,
               Optional ByVal EnableWinXp_VisualStyles As Boolean = True,
               Optional ByVal SingleInstanceApplication As Boolean = False,
               Optional ByVal ForceRunAsAdministrator As Boolean = False)



            If ForceRunAsAdministrator Then

                REM If not XP
                If (Environment.OSVersion.Version.Major <> 5 AndAlso Not IsRunAsAdministrator()) Then
                    MsgBox("Please, start application as Administrator",
                           CType(MsgBoxStyle.Exclamation + MsgBoxStyle.SystemModal, MsgBoxStyle), "ERROR")

                    Return
                End If


            End If


            If Not isValid_Form(StartUp_Form) Then MyLogFile.Log("FrameWork Cant run because of Invalid Startup Form") : Exit Sub REM This will not work because handle is not created until you call form.show

            Dim AppMutex As System.Threading.Mutex = Nothing

            If SingleInstanceApplication Then
                REM Lock using mutex
                Dim GUID As String = Me.get_guid_from_ObjectType(StartUp_Form.GetType())
                If GUID <> String.Empty Then

                    AppMutex = New System.Threading.Mutex(False, GUID)
                    If Not AppMutex.WaitOne(TimeSpan.Zero, False) Then
                        Try
                            '
                            '   Try to bring to focus
                            '
                            Dim proc = Process.GetProcessesByName(EIO.getFileNameWithoutExtension(Application.ExecutablePath)).FirstOrDefault()
                            If proc IsNot Nothing Then
                                If IsDebugMode Then
                                    Program.Logger.Print("proc.MainWindowHandle.ToInt64(): " & proc.MainWindowHandle.ToInt64())
                                    Program.Logger.Print("proc.MainWindowHandle.ToInt32(): " & proc.MainWindowHandle.ToInt32())
                                    Program.Logger.Print("proc.Handle.ToInt32(): " & proc.Handle.ToInt32())

                                End If

                                'Try
                                '    ' ''
                                '    ' ''   If minimized then restore
                                '    ' ''
                                '    ''Dim placement As Windowplacement = New Windowplacement()
                                '    ''GetWindowPlacement(proc.MainWindowHandle, placement)

                                '    ''If placement.showCmd = 2 Then _

                                '    ShowWindow(proc.Handle, ShowWindowEnum.Restore) 'Restore Window

                                'Catch ex As Exception
                                '    ' Ignore if it isnt minimized
                                '    Program.Logger.Print(ex)
                                'End Try

                                Try
                                    ' ''
                                    ' ''   If minimized then restore
                                    ' ''
                                    ''Dim placement As Windowplacement = New Windowplacement()
                                    ''GetWindowPlacement(proc.MainWindowHandle, placement)

                                    ''If placement.showCmd = 2 Then _
                                    Debug.Print("Slowing ... down ...restore ")
                                    ShowWindow(proc.MainWindowHandle, ShowWindowEnum.Restore) 'Restore Window

                                Catch ex As Exception
                                    ' Ignore if it isnt minimized
                                    Program.Logger.Print(ex)
                                End Try

                                'Try
                                '    ' ''
                                '    ' ''   If minimized then restore
                                '    ' ''
                                '    ''Dim placement As Windowplacement = New Windowplacement()
                                '    ''GetWindowPlacement(proc.MainWindowHandle, placement)

                                '    ''If placement.showCmd = 2 Then _

                                '    ShowWindow(proc.MainWindowHandle, ShowWindowEnum.ShowMaximized) 'Restore Window

                                'Catch ex As Exception
                                '    ' Ignore if it isnt minimized
                                '    Program.Logger.Print(ex)
                                'End Try
                              

                                '
                                '   Set Infront
                                '
                                Dim p As System.Windows.Automation.AutomationElement = Windows.Automation.AutomationElement.FromHandle(proc.MainWindowHandle)
                                p.SetFocus()
                            Else
                                Throw New Exception("Could not get this application's MainWindowHandle")
                            End If
                        Catch ex As Exception
                            Program.Logger.Print(ex)
                            MyLogFile.Log("An Instance of this Application is already running.")
                            MsgBox("An Instance of this Application is already running.")
                        End Try

                        Exit Sub
                    End If
                End If
            End If



            If EnableWinXp_VisualStyles Then Application.EnableVisualStyles()

            REM Optional ByVal SetCompatibleTextRenderingDefault As Boolean = False,
            REM This will not work because it must be set before any form can be created
            REM Application.SetCompatibleTextRenderingDefault(SetCompatibleTextRenderingDefault)


            REM MyLogFile.Log(Objects.EForm.get_guid_from_ObjectType(StartUp_Form.GetType()))

            Select Case ShutDownMode
                Case Is = ShutdownModes.CLOSE_WHEN_START_UP_FORM_CLOSES
                    Application.Run(StartUp_Form)

                Case Is = ShutdownModes.CLOSE_WHEN_LAST_FORM_CLOSES




                    REM NOTE: This Form Event Handler will be on the second or last on
                    REM       the form's handler list .. meaning our decision on e.Cancel is final :)
                    Me.CloseApp_FormHandler = StartUp_Form
                    AddHandler StartUp_Form.FormClosing, AddressOf Me.PreventClosing
                    AddHandler StartUp_Form.Load, AddressOf Initiate_StartUp_Params

                    REM     This line can be released if
                    REM     1.) This form is closed
                    REM     2.) This form is hidden which will execute on form closing event
                    REM         Not form closed
                    REM     3.) This form is disposed
                    REM 
                    REM     A good way to know if user is closing or hidding a form is
                    REM     e.CloseReason=None on hidden and disposing
                    REM     e.CloseReason=OtherValues on Closing


                    '' ''on FormClosing Event
                    '' ''a close form
                    '' ''sender:         itself()
                    '' ''Text:           valid()
                    '' ''closereason:    valid()
                    '' ''isDisposed: false
                    '' ''isAccessible:false


                    '' ''a hidden form
                    '' ''sender:         itself()
                    '' ''Text:           valid()
                    '' ''closereason:    none()
                    '' ''isDisposed: false
                    '' ''isAccessible:false


                    '' ''A disposed form
                    '' ''sender:         itself()
                    '' ''Text:           null()
                    '' ''closereason:    none()
                    '' ''isDisposed: true
                    '' ''isAccessible:false

                    REM DIfference between disposing and hidden
                    REM isDisposed =True on disposing while on hidden it is false

                    StartUp_Form.ShowDialog()


                    ''You can still hide your start up form it doesnt affect it
                    ''When you call hide() 

                    ''The form is popped from application.OpenForms list

                    ''and form closing event is call which apparently can not be closed from external thread
                    REM Normally this place shouldnt run unless it is closed. If you hide it and didnt open any form before hiding it. That is bad
                    If Framework.IsDebugMode Then MyLogFile.Log("Frame Work Exited After Show Dialog closed. Me.IsDisposed: " & Me.isDisposed.ToString())
                    If Not Me.isDisposed Then MyLogFile.Log("It is NOT a good practice to hide start up form. To get better performance, just close it.")
                    While Not Me.isDisposed
                        REM MyLogFile.Log("Waiting for thread to close:")
                        Application.DoEvents()
                        Threading.Thread.Sleep(1)
                    End While

                    REM Thread realeased due to some reasons
                    REM Wait for the asyc thread to terminate
                    ''While Not Me.isDisposed
                    ''    MyLogFile.Log("Waiting for thread to close:")
                    ''    Threading.Thread.Sleep(1000)
                    ''End While

                    REM The loop is hanging the whole application because it is the main thread
                    REM So if it is released and not disposed reshow it

                    REM Incase this was a call from application.exit then inform thread that is enough
                    REM If Not Me.isDisposed Then Me.isDisposed = True


            End Select



            If AppMutex IsNot Nothing Then AppMutex.ReleaseMutex() : AppMutex = Nothing


        End Sub



#End Region


#Region "Property"

        ''' <summary>
        ''' The form that serves as a pointer for us to close app
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property CloseApp_FormHandler As Form = Nothing

        ''' <summary>
        ''' Indicates the application is closed
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Property isDisposed As Boolean = False

        ''' <summary>
        ''' The title of the handler form
        ''' </summary>
        ''' <remarks></remarks>
        Private Const HANDLER_TITLE As String = "DO_NOT_TOUCH_HANDLER"


        ''' <summary>
        ''' Used to monitor app
        ''' </summary>
        ''' <remarks></remarks>
        Private TChecker As Threading.Thread = Nothing

        Public Shared Property IsDebugMode As Boolean = False

#End Region


#Region "Enumerations"

        Public Enum ShutdownModes
            CLOSE_WHEN_START_UP_FORM_CLOSES
            CLOSE_WHEN_LAST_FORM_CLOSES
        End Enum

#End Region

#Region "Methods"


        Private Function IsRunAsAdministrator() As Boolean
            Dim wi = WindowsIdentity.GetCurrent()
            Dim wp = New WindowsPrincipal(wi)
            Return wp.IsInRole(WindowsBuiltInRole.Administrator)


            'Dim processInfo = New ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase)
            ''        processInfo.UseShellExecute = True
            ''        processInfo.Verb = "runas"
            ''        Try
            ''            Process.Start(processInfo)
            ''        Catch ex As Exception
            ''            MsgBox("This application needs to be run as Administrator. Please restart Navigate as Administrator")
            ''        End Try
            ''        Me.Close()
        End Function



        Private Sub Run()

            Try

                While Not Me.isDisposed
                    ''Try

                    ''    REM Debugging
                    ''    MyLogFile.Log("Application.OpenForms.Count: " & Application.OpenForms.Count)
                    ''    For Each F As Form In Application.OpenForms
                    ''        MyLogFile.Log("Form.Name : " & F.Name)
                    ''    Next

                    ''Catch ex As Exception
                    ''    REM A form closed while we are running for each loop
                    ''    MyLogFile.Log(New EException(ex))
                    ''End Try


                    REM I discovered none of the application forms refers to this startup form i kept
                    REM It is absolutely disconnected
                    If Framework.IsDebugMode Then MyLogFile.Log("Thread is running with Openforms.count = : " & Application.OpenForms.Count)
                    ''If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 Then MyLogFile.Log("First Form: " & Application.OpenForms(0).Name)
                    ''If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 1 Then MyLogFile.Log("Second Form: " & Application.OpenForms(1).Name)


                    If Application.OpenForms.Count = 0 OrElse
                        (Application.OpenForms.Count = 1 AndAlso Application.OpenForms(0).Name = HANDLER_TITLE) Then

JUST_CLOSE_THE_DAMN_STUFF:
                        If Framework.IsDebugMode Then MyLogFile.Log("Entering the loop because count = : " & Application.OpenForms.Count)
                        If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 Then MyLogFile.Log("First Form: " & Application.OpenForms(0).Name)

                        REM Indicate we are actually closing the app
                        Me.isDisposed = True


                        REM If isValid_Form(Me.CloseApp_FormHandler) AndAlso Me.CloseApp_FormHandler.IsHandleCreated Then
                        REM Me.CloseApp_FormHandler.Close()
                        REM Or I can just use dispose
                        If isValid_Form(Me.CloseApp_FormHandler) Then
                            If Me.CloseApp_FormHandler.IsHandleCreated Then
                                Me.CloseApp_FormHandler.Invoke(Sub() Me.CloseApp_FormHandler.Dispose())
                            Else
                                MyLogFile.Log("You are writing your program in a really wrong way. You hid the startup form")
                                REM Application.Exit()  REM this is really bad . . it will still close because it will return the foreground thread
                            End If
                            REM Once that form is disposed the remaining part doesnt run
                            If Framework.IsDebugMode Then MyLogFile.Log("Closed Application Form Handler from thread")
                        End If


                        Exit While

                    ElseIf Application.OpenForms.Count = 1 AndAlso Application.OpenForms(0).Name = String.Empty AndAlso
                        Application.OpenForms(0).Visible = False Then
                        GoTo JUST_CLOSE_THE_DAMN_STUFF
                        REM The name just couldn't set
                    ElseIf Application.OpenForms.Count = 1 Then
                        REM Finds it difficult to set properties on a disposing form
                        REM for debugging
                        If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 Then MyLogFile.Log("First Form: " & Application.OpenForms(0).Name)
                        If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 AndAlso Application.OpenForms(0).Tag IsNot Nothing Then MyLogFile.Log("First Form Tag: " & Application.OpenForms(0).Tag.ToString())
                        If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 Then MyLogFile.Log("First Form Visibility: " & Application.OpenForms(0).Visible.ToString())
                    End If

                    Threading.Thread.Sleep(2000)
                End While

            Catch ex As Threading.ThreadAbortException

                MyLogFile.Log(New EException(ex))

            Catch ex As Exception

                MyLogFile.Log(New EException(ex))

            End Try

            REM When you run a foreground thread, Try to dispose it
            REM If Me.TChecker.IsAlive Then Me.TChecker.Abort() you cant exit a thread in same thread

            If Framework.IsDebugMode Then MyLogFile.Log("TChecker Thread Exiting!")
        End Sub




        Private Sub PreventClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)


            e.Cancel = Not Me.isDisposed

            REM Check if it is really closing
            REM e.CloseReason <> CloseReason.None == This eliminates disposing and hidden
            If Framework.IsDebugMode Then MyLogFile.Log(e.CloseReason.ToString())

            If e.Cancel AndAlso e.CloseReason <> CloseReason.None Then
                With Me.CloseApp_FormHandler
                    .Visible = True REM This will invoke another prevent close with e.Reason=None
                    .Opacity = 0
                    .WindowState = FormWindowState.Minimized
                    .ShowInTaskbar = False
                    .Name = HANDLER_TITLE       REM having problem setting the name
                    CType(sender, Form).Text = HANDLER_TITLE
                    .Tag = HANDLER_TITLE
                    If Framework.IsDebugMode Then MyLogFile.Log("Name changed to: " & .Name)

                End With

            End If

            If Framework.IsDebugMode Then MyLogFile.Log("PreventClosing: Frame Work Cancel Close: " & e.Cancel.ToString())

        End Sub




        ''' <summary>
        ''' Fetch GUID from an obj.getType
        ''' </summary>
        ''' <param name="objType"></param>
        ''' <returns></returns>
        ''' <remarks>Actually I can fetch the GUID from the application directly</remarks>
        Private Function get_guid_from_ObjectType(ByVal objType As System.Type) As String
            Try


                Return New Guid(
                                CType(
                                    objType.Assembly.GetCustomAttributes(GetType(GuidAttribute), False)(0), 
                                    GuidAttribute).Value()
                                   ).ToString
            Catch ex As Exception
                Return String.Empty
            End Try

        End Function



        REM C# Closes because Application.Run() Closes all forms when startup app closes automatically
        REM Only the startup form run as foreground thread. The rest runs as background which do not prevent an app from terminating

        Private Sub Initiate_StartUp_Params(ByVal sender As Object, ByVal e As EventArgs)
            REM  MyLogFile.Log("Start Up OnLoad Event.")
            REM Run StartUp form on a different thread and sleep this one
            REM until no form is on
            If Me.TChecker Is Nothing Then
                Me.TChecker = New Threading.Thread(AddressOf Me.Run) With {.IsBackground = True}
                With Me.TChecker
                    .SetApartmentState(Threading.ApartmentState.MTA) REM Doesnt work if you need OLE Calls
                    .Start()
                End With
            End If
        End Sub




#End Region







    End Class

End Namespace
