using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.MicrosoftOS
{
    /// <summary>
    /// handles request to cmd.exe
    /// </summary>
    /// <remarks>Make sure cmd.exe is available and enabled by administrator</remarks>
    public class CommandPrompt
    {

        #region Constructors

        /// <summary>
        /// NOTE: if set to pipe result, It doesnt affect the command even if it contains a pipe. It will only override it
        /// </summary>
        /// <param name="PipeResults">If it is not true, Result will always return empty string</param>
        /// <remarks></remarks>
        public CommandPrompt(bool PipeResults = true, bool Visible = false)
        {
            _PipeResults = PipeResults;
            _Visible = Visible;
        }


        #endregion



        public const int WAIT__TILL__INFINITY = -1;


        #region Properties

        private bool _PipeResults = false;

        public bool PipeResults
        {
            get
            {
                return _PipeResults;
            }
        }

        private bool _Visible = false;

        public bool Visible
        {
            get
            {
                return _Visible;
            }
        }


        // 'Public ReadOnly Property Result As String
        // '    Get
        // '        If Me.PipeResults AndAlso File.Exists(Me.ResultFile) Then
        // '            Try

        // '                Return File.ReadAllText(Me.ResultFile)

        // '            Catch ex As Exception
        // '                Me._LastError = ex
        // '                Debug.Print(ex.Message)
        // '            End Try

        // '        End If
        // '        Return String.Empty
        // '    End Get
        // 'End Property



        // 'Public ReadOnly Property ResultFile As String
        // '    Get
        // '        Return String.Format("{0}\{1}.cmdOutput", Application.StartupPath, get_guid_from_ObjectType(Me.GetType()))
        // '    End Get
        // 'End Property
        private string _Result = string.Empty;

        public string Result
        {
            get
            {
                return _Result;
            }
        }

        private Exception _LastError = new Exception();

        public Exception LastError
        {
            get
            {
                return _LastError;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Execute a Command Prompt Code Like dir * /w
        /// </summary>
        /// <param name="args">dir * /w</param>
        /// <param name="WorkingDirectory">The folder you want the command prompt to start from</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Execute(string args, string WorkingDirectory)
        {
            var p = new Process();
            var StartInfo = new ProcessStartInfo();
            if (!Visible)
            {
                StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            }
            else
            {
                StartInfo.WindowStyle = ProcessWindowStyle.Normal;
            }

            StartInfo.WorkingDirectory = WorkingDirectory;
            StartInfo.FileName = "cmd.exe";
            // REM StartInfo.Arguments = "/C " & args
            StartInfo.Arguments = string.Format("/C \"{0}\"", args);
            // REM /K wont matter here any more because if the output is redirected
            // REM it will automatically close after execution
            // REM Also it wont display a thing on the console.. it is just similar to piping

            // REM I am using this since I will only be executing an exe which is cmd.exe
            StartInfo.RedirectStandardOutput = PipeResults;  // REM So I can read from standard output
            StartInfo.UseShellExecute = !PipeResults; // REM If this is false redirection will work and also i can run only .exe file with this process only
            // REM If ShellExecute is false it will always be visible

            // 'StartInfo.RedirectStandardOutput = False REM So I can read from standard output
            // 'StartInfo.UseShellExecute = True   REM If this is false redirection will work and also i can run only .exe file with this process only


            // ' If Me.PipeResults Then StartInfo.Arguments &= " >" & Me.ResultFile

            p.StartInfo = StartInfo;
            try
            {
                p.Start();
                if (PipeResults)
                    _Result = p.StandardOutput.ReadToEnd();
                p.WaitForExit();
                while (!p.HasExited)
                    Application.DoEvents();
            }
            catch (Exception ex)
            {
                _LastError = ex;
                return false;
            }

            return true;
        }


        /// <summary>
        /// Execute a Command Prompt Code Like dir * /w
        /// </summary>
        /// <param name="args">dir * /w</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Execute(string args)
        {
            return Execute(args, Application.StartupPath);
        }






        /// <summary>
        /// Executes commands using Interaction.Shell, uses cmd.exe from system32
        /// </summary>
        /// <param name="pCommand"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool Execute(string pCommand, int pTimeoutMilliSeconds)
        {
            var vMode = AppWinStyle.NormalFocus;
            if (!Visible)
                vMode = AppWinStyle.Hide;
            try
            {
                // REM cmd.exe /K "Color 47"
                // REM cmd.exe /C "Color 47"
                var pFullCommand = new StringBuilder(@"""C:\Windows\System32\cmd.exe"" /C");
                pFullCommand.Append(pCommand);
                string pResultFile = string.Format(@"{0}\ShellExcuteGenerated.log", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData));
                if (PipeResults)
                    pFullCommand.Append(string.Format(" >\"{0}\"", pResultFile));
                int procID = Interaction.Shell(pFullCommand.ToString(), vMode, PipeResults, pTimeoutMilliSeconds);
                Modules.basMain.MyLogFile.Print("ShellExecute: - Generated Process ID: " + procID);
                if (PipeResults && System.IO.File.Exists(pResultFile))
                    _Result = System.IO.File.ReadAllText(pResultFile);
                EIO.DeleteFileIfExists(pResultFile);
                return true;
            }
            catch (Exception ex)
            {
                _LastError = ex;
                Modules.basMain.MyLogFile.Print("ShellExecute: ", ex);
                return false;
            }
        }




        /// <summary>
        /// Executes commands using Interaction.Shell, uses cmd.exe from system32 and uses temporary file batch
        /// </summary>
        /// <param name="pCommand"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public bool ForceExecute(string pCommand, int pTimeoutMilliSeconds = WAIT__TILL__INFINITY)
        {
            var vMode = AppWinStyle.NormalFocus;
            if (!Visible)
                vMode = AppWinStyle.Hide;
            try
            {
                string pResultFile = string.Format(@"{0}\ShellExcuteGenerated.log", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData));
                string pBatchFile = string.Format(@"{0}\ShellExcuteTemp.bat", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData));
                // REM cmd.exe /K "Color 47"
                // REM cmd.exe /C "Color 47"
                var FullCommand = new StringBuilder(@"""C:\Windows\System32\cmd.exe"" /C");
                FullCommand.Append(string.Format(" \"{0}\"", pBatchFile));
                var UserCommand = new StringBuilder(pCommand);
                if (PipeResults)
                    UserCommand.Append(string.Format(" >\"{0}\"", pResultFile));
                System.IO.File.WriteAllText(pBatchFile, UserCommand.ToString());
                int procID = Interaction.Shell(FullCommand.ToString(), vMode, PipeResults, pTimeoutMilliSeconds);
                Modules.basMain.MyLogFile.Print("ShellExecute: - Generated Process ID: " + procID);
                if (PipeResults && System.IO.File.Exists(pResultFile))
                    _Result = System.IO.File.ReadAllText(pResultFile);
                EIO.DeleteFileIfExists(pResultFile);
                EIO.DeleteFileIfExists(pBatchFile);
                return true;
            }
            catch (Exception ex)
            {
                _LastError = ex;
                Modules.basMain.MyLogFile.Print("ShellExecute: ", ex);
                return false;
            }
        }













        /// <summary>
        /// Returns the Last result
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public override string ToString()
        {
            return Result;
        }


        #endregion

    }
}