using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Windows.Forms;
using CODERiT.Logger.v._3._5.Exceptions;
using ELibrary.Standard.Modules;
using static ELibrary.Standard.Objects.EForm;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.AppConfigurations
{
    public sealed class Framework
    {


        #region Showing Window
        private enum ShowWindowEnum
        {
            Hide = 0,
            ShowNormal = 1,
            ShowMinimized = 2,
            ShowMaximized = 3,
            Maximize = 3,
            ShowNormalNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActivate = 7,
            ShowNoActivate = 8,
            Restore = 9,
            ShowDefault = 10,
            ForceMinimized = 11
        }

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, ShowWindowEnum flags);

        private struct Windowplacement
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref Windowplacement lpwndpl);


        #endregion





        #region Constructors


        public Framework(Form StartUp_Form, ShutdownModes ShutDownMode = ShutdownModes.CLOSE_WHEN_LAST_FORM_CLOSES, bool EnableWinXp_VisualStyles = true, bool SingleInstanceApplication = false, bool ForceRunAsAdministrator = false)
        {
            if (ForceRunAsAdministrator)
            {

                // REM If not XP
                if (Environment.OSVersion.Version.Major != 5 && !IsRunAsAdministrator())
                {
                    Interaction.MsgBox("Please, start application as Administrator", (MsgBoxStyle)Conversions.ToInteger((int)MsgBoxStyle.Exclamation + (int)MsgBoxStyle.SystemModal), "ERROR");
                    return;
                }
            }

            if (!isValid_Form(StartUp_Form))
            {
                basMain.MyLogFile.Log("FrameWork Cant run because of Invalid Startup Form");
                return;
            } // REM This will not work because handle is not created until you call form.show

            System.Threading.Mutex AppMutex = null;
            if (SingleInstanceApplication)
            {
                // REM Lock using mutex
                string GUID = get_guid_from_ObjectType(StartUp_Form.GetType());
                if ((GUID ?? "") != (string.Empty ?? ""))
                {
                    AppMutex = new System.Threading.Mutex(false, GUID);
                    if (!AppMutex.WaitOne(TimeSpan.Zero, false))
                    {
                        try
                        {
                            // 
                            // Try to bring to focus
                            // 
                            var proc = Process.GetProcessesByName(EIO.getFileNameWithoutExtension(Application.ExecutablePath)).FirstOrDefault();
                            if (proc is object)
                            {
                                if (IsDebugMode)
                                {
                                    Program.Logger.Print("proc.MainWindowHandle.ToInt64(): " + proc.MainWindowHandle.ToInt64());
                                    Program.Logger.Print("proc.MainWindowHandle.ToInt32(): " + proc.MainWindowHandle.ToInt32());
                                    Program.Logger.Print("proc.Handle.ToInt32(): " + proc.Handle.ToInt32());
                                }

                                // Try
                                // ' ''
                                // ' ''   If minimized then restore
                                // ' ''
                                // ''Dim placement As Windowplacement = New Windowplacement()
                                // ''GetWindowPlacement(proc.MainWindowHandle, placement)

                                // ''If placement.showCmd = 2 Then _

                                // ShowWindow(proc.Handle, ShowWindowEnum.Restore) 'Restore Window

                                // Catch ex As Exception
                                // ' Ignore if it isnt minimized
                                // Program.Logger.Print(ex)
                                // End Try

                                try
                                {
                                    // ''
                                    // ''   If minimized then restore
                                    // ''
                                    // 'Dim placement As Windowplacement = New Windowplacement()
                                    // 'GetWindowPlacement(proc.MainWindowHandle, placement)

                                    // 'If placement.showCmd = 2 Then _
                                    Debug.Print("Slowing ... down ...restore ");
                                    ShowWindow(proc.MainWindowHandle, ShowWindowEnum.Restore); // Restore Window
                                }
                                catch (Exception ex)
                                {
                                    // Ignore if it isnt minimized
                                    Program.Logger.Print(ex);
                                }

                                // Try
                                // ' ''
                                // ' ''   If minimized then restore
                                // ' ''
                                // ''Dim placement As Windowplacement = New Windowplacement()
                                // ''GetWindowPlacement(proc.MainWindowHandle, placement)

                                // ''If placement.showCmd = 2 Then _

                                // ShowWindow(proc.MainWindowHandle, ShowWindowEnum.ShowMaximized) 'Restore Window

                                // Catch ex As Exception
                                // ' Ignore if it isnt minimized
                                // Program.Logger.Print(ex)
                                // End Try


                                // 
                                // Set Infront
                                // 
                                var p = System.Windows.Automation.AutomationElement.FromHandle(proc.MainWindowHandle);
                                p.SetFocus();
                            }
                            else
                            {
                                throw new Exception("Could not get this application's MainWindowHandle");
                            }
                        }
                        catch (Exception ex)
                        {
                            Program.Logger.Print(ex);
                            basMain.MyLogFile.Log("An Instance of this Application is already running.");
                            Interaction.MsgBox("An Instance of this Application is already running.");
                        }

                        return;
                    }
                }
            }

            if (EnableWinXp_VisualStyles)
                Application.EnableVisualStyles();

            // REM Optional ByVal SetCompatibleTextRenderingDefault As Boolean = False,
            // REM This will not work because it must be set before any form can be created
            // REM Application.SetCompatibleTextRenderingDefault(SetCompatibleTextRenderingDefault)


            // REM MyLogFile.Log(Objects.EForm.get_guid_from_ObjectType(StartUp_Form.GetType()))

            switch (ShutDownMode)
            {
                case var @case when @case == ShutdownModes.CLOSE_WHEN_START_UP_FORM_CLOSES:
                    {
                        Application.Run(StartUp_Form);
                        break;
                    }

                case var case1 when case1 == ShutdownModes.CLOSE_WHEN_LAST_FORM_CLOSES:
                    {




                        // REM NOTE: This Form Event Handler will be on the second or last on
                        // REM       the form's handler list .. meaning our decision on e.Cancel is final :)
                        CloseApp_FormHandler = StartUp_Form;
                        StartUp_Form.FormClosing += PreventClosing;
                        StartUp_Form.Load += Initiate_StartUp_Params;

                        // REM     This line can be released if
                        // REM     1.) This form is closed
                        // REM     2.) This form is hidden which will execute on form closing event
                        // REM         Not form closed
                        // REM     3.) This form is disposed
                        // REM 
                        // REM     A good way to know if user is closing or hidding a form is
                        // REM     e.CloseReason=None on hidden and disposing
                        // REM     e.CloseReason=OtherValues on Closing


                        // ' ''on FormClosing Event
                        // ' ''a close form
                        // ' ''sender:         itself()
                        // ' ''Text:           valid()
                        // ' ''closereason:    valid()
                        // ' ''isDisposed: false
                        // ' ''isAccessible:false


                        // ' ''a hidden form
                        // ' ''sender:         itself()
                        // ' ''Text:           valid()
                        // ' ''closereason:    none()
                        // ' ''isDisposed: false
                        // ' ''isAccessible:false


                        // ' ''A disposed form
                        // ' ''sender:         itself()
                        // ' ''Text:           null()
                        // ' ''closereason:    none()
                        // ' ''isDisposed: true
                        // ' ''isAccessible:false

                        // REM DIfference between disposing and hidden
                        // REM isDisposed =True on disposing while on hidden it is false

                        StartUp_Form.ShowDialog();


                        // 'You can still hide your start up form it doesnt affect it
                        // 'When you call hide() 

                        // 'The form is popped from application.OpenForms list

                        // 'and form closing event is call which apparently can not be closed from external thread
                        // REM Normally this place shouldnt run unless it is closed. If you hide it and didnt open any form before hiding it. That is bad
                        if (IsDebugMode)
                            basMain.MyLogFile.Log("Frame Work Exited After Show Dialog closed. Me.IsDisposed: " + isDisposed.ToString());
                        if (!isDisposed)
                            basMain.MyLogFile.Log("It is NOT a good practice to hide start up form. To get better performance, just close it.");
                        while (!isDisposed)
                        {
                            // REM MyLogFile.Log("Waiting for thread to close:")
                            Application.DoEvents();
                            System.Threading.Thread.Sleep(1);
                        }

                        break;
                    }

                    // REM Thread realeased due to some reasons
                    // REM Wait for the asyc thread to terminate
                    // 'While Not Me.isDisposed
                    // '    MyLogFile.Log("Waiting for thread to close:")
                    // '    Threading.Thread.Sleep(1000)
                    // 'End While

                    // REM The loop is hanging the whole application because it is the main thread
                    // REM So if it is released and not disposed reshow it

                    // REM Incase this was a call from application.exit then inform thread that is enough
                    // REM If Not Me.isDisposed Then Me.isDisposed = True


            }

            if (AppMutex is object)
            {
                AppMutex.ReleaseMutex();
                AppMutex = null;
            }
        }



        #endregion


        #region Property

        /// <summary>
        /// The form that serves as a pointer for us to close app
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private Form CloseApp_FormHandler { get; set; } = null;

        /// <summary>
        /// Indicates the application is closed
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool isDisposed { get; set; } = false;

        /// <summary>
        /// The title of the handler form
        /// </summary>
        /// <remarks></remarks>
        private const string HANDLER_TITLE = "DO_NOT_TOUCH_HANDLER";


        /// <summary>
        /// Used to monitor app
        /// </summary>
        /// <remarks></remarks>
        private System.Threading.Thread TChecker = null;

        public static bool IsDebugMode { get; set; } = false;

        #endregion


        #region Enumerations

        public enum ShutdownModes
        {
            CLOSE_WHEN_START_UP_FORM_CLOSES,
            CLOSE_WHEN_LAST_FORM_CLOSES
        }

        #endregion

        #region Methods


        private bool IsRunAsAdministrator()
        {
            var wi = WindowsIdentity.GetCurrent();
            var wp = new WindowsPrincipal(wi);
            return wp.IsInRole(WindowsBuiltInRole.Administrator);


            // Dim processInfo = New ProcessStartInfo(Assembly.GetExecutingAssembly().CodeBase)
            // '        processInfo.UseShellExecute = True
            // '        processInfo.Verb = "runas"
            // '        Try
            // '            Process.Start(processInfo)
            // '        Catch ex As Exception
            // '            MsgBox("This application needs to be run as Administrator. Please restart Navigate as Administrator")
            // '        End Try
            // '        Me.Close()
        }

        private void Run()
        {
            try
            {
                while (!isDisposed)
                {
                    // 'Try

                    // '    REM Debugging
                    // '    MyLogFile.Log("Application.OpenForms.Count: " & Application.OpenForms.Count)
                    // '    For Each F As Form In Application.OpenForms
                    // '        MyLogFile.Log("Form.Name : " & F.Name)
                    // '    Next

                    // 'Catch ex As Exception
                    // '    REM A form closed while we are running for each loop
                    // '    MyLogFile.Log(New EException(ex))
                    // 'End Try


                    // REM I discovered none of the application forms refers to this startup form i kept
                    // REM It is absolutely disconnected
                    if (IsDebugMode)
                        basMain.MyLogFile.Log("Thread is running with Openforms.count = : " + Application.OpenForms.Count);
                    // 'If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 0 Then MyLogFile.Log("First Form: " & Application.OpenForms(0).Name)
                    // 'If Framework.IsDebugMode AndAlso Application.OpenForms.Count > 1 Then MyLogFile.Log("Second Form: " & Application.OpenForms(1).Name)


                    if (Application.OpenForms.Count == 0 || Application.OpenForms.Count == 1 && (Application.OpenForms[0].Name ?? "") == HANDLER_TITLE)
                    {
                    JUST_CLOSE_THE_DAMN_STUFF:
                        ;
                        if (IsDebugMode)
                            basMain.MyLogFile.Log("Entering the loop because count = : " + Application.OpenForms.Count);
                        if (IsDebugMode && Application.OpenForms.Count > 0)
                            basMain.MyLogFile.Log("First Form: " + Application.OpenForms[0].Name);

                        // REM Indicate we are actually closing the app
                        isDisposed = true;


                        // REM If isValid_Form(Me.CloseApp_FormHandler) AndAlso Me.CloseApp_FormHandler.IsHandleCreated Then
                        // REM Me.CloseApp_FormHandler.Close()
                        // REM Or I can just use dispose
                        if (isValid_Form(CloseApp_FormHandler))
                        {
                            if (CloseApp_FormHandler.IsHandleCreated)
                            {
                                CloseApp_FormHandler.Invoke(new Action(() => CloseApp_FormHandler.Dispose()));
                            }
                            else
                            {
                                basMain.MyLogFile.Log("You are writing your program in a really wrong way. You hid the startup form");
                                // REM Application.Exit()  REM this is really bad . . it will still close because it will return the foreground thread
                            }
                            // REM Once that form is disposed the remaining part doesnt run
                            if (IsDebugMode)
                                basMain.MyLogFile.Log("Closed Application Form Handler from thread");
                        }

                        break;
                    }
                    else if (Application.OpenForms.Count == 1 && (Application.OpenForms[0].Name ?? "") == (string.Empty ?? "") && Application.OpenForms[0].Visible == false)
                    {
                        goto JUST_CLOSE_THE_DAMN_STUFF;
                    }
                    // REM The name just couldn't set
                    else if (Application.OpenForms.Count == 1)
                    {
                        // REM Finds it difficult to set properties on a disposing form
                        // REM for debugging
                        if (IsDebugMode && Application.OpenForms.Count > 0)
                            basMain.MyLogFile.Log("First Form: " + Application.OpenForms[0].Name);
                        if (IsDebugMode && Application.OpenForms.Count > 0 && Application.OpenForms[0].Tag is object)
                            basMain.MyLogFile.Log("First Form Tag: " + Application.OpenForms[0].Tag.ToString());
                        if (IsDebugMode && Application.OpenForms.Count > 0)
                            basMain.MyLogFile.Log("First Form Visibility: " + Application.OpenForms[0].Visible.ToString());
                    }

                    System.Threading.Thread.Sleep(2000);
                }
            }
            catch (System.Threading.ThreadAbortException ex)
            {
                basMain.MyLogFile.Log(new EException(ex));
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Log(new EException(ex));
            }

            // REM When you run a foreground thread, Try to dispose it
            // REM If Me.TChecker.IsAlive Then Me.TChecker.Abort() you cant exit a thread in same thread

            if (IsDebugMode)
                basMain.MyLogFile.Log("TChecker Thread Exiting!");
        }

        private void PreventClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = !isDisposed;

            // REM Check if it is really closing
            // REM e.CloseReason <> CloseReason.None == This eliminates disposing and hidden
            if (IsDebugMode)
                basMain.MyLogFile.Log(e.CloseReason.ToString());
            if (e.Cancel && e.CloseReason != CloseReason.None)
            {
                {
                    var withBlock = CloseApp_FormHandler;
                    withBlock.Visible = true; // REM This will invoke another prevent close with e.Reason=None
                    withBlock.Opacity = 0d;
                    withBlock.WindowState = FormWindowState.Minimized;
                    withBlock.ShowInTaskbar = false;
                    withBlock.Name = HANDLER_TITLE;       // REM having problem setting the name
                    ((Form)sender).Text = HANDLER_TITLE;
                    withBlock.Tag = HANDLER_TITLE;
                    if (IsDebugMode)
                        basMain.MyLogFile.Log("Name changed to: " + withBlock.Name);
                }
            }

            if (IsDebugMode)
                basMain.MyLogFile.Log("PreventClosing: Frame Work Cancel Close: " + e.Cancel.ToString());
        }




        /// <summary>
        /// Fetch GUID from an obj.getType
        /// </summary>
        /// <param name="objType"></param>
        /// <returns></returns>
        /// <remarks>Actually I can fetch the GUID from the application directly</remarks>
        private string get_guid_from_ObjectType(Type objType)
        {
            try
            {
                return new Guid(((GuidAttribute)objType.Assembly.GetCustomAttributes(typeof(GuidAttribute), false)[0]).Value).ToString();
            }
            catch (Exception ex)
            {
                return string.Empty;
            }
        }



        // REM C# Closes because Application.Run() Closes all forms when startup app closes automatically
        // REM Only the startup form run as foreground thread. The rest runs as background which do not prevent an app from terminating

        private void Initiate_StartUp_Params(object sender, EventArgs e)
        {
            // REM  MyLogFile.Log("Start Up OnLoad Event.")
            // REM Run StartUp form on a different thread and sleep this one
            // REM until no form is on
            if (TChecker is null)
            {
                TChecker = new System.Threading.Thread(Run) { IsBackground = true };
                {
                    var withBlock = TChecker;
                    withBlock.SetApartmentState(System.Threading.ApartmentState.MTA); // REM Doesnt work if you need OLE Calls
                    withBlock.Start();
                }
            }
        }




        #endregion







    }
}