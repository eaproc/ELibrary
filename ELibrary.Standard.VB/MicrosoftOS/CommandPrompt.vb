Imports System.Windows.Forms
Imports System.Text


Namespace MicrosoftOS
    ''' <summary>
    ''' handles request to cmd.exe
    ''' </summary>
    ''' <remarks>Make sure cmd.exe is available and enabled by administrator</remarks>
    Public Class CommandPrompt

#Region "Constructors"

        ''' <summary>
        ''' NOTE: if set to pipe result, It doesnt affect the command even if it contains a pipe. It will only override it
        ''' </summary>
        ''' <param name="PipeResults">If it is not true, Result will always return empty string</param>
        ''' <remarks></remarks>
        Sub New(Optional ByVal PipeResults As Boolean = True,
                Optional ByVal Visible As Boolean = False)

            Me._PipeResults = PipeResults
            Me._Visible = Visible

        End Sub


#End Region



        Public Const WAIT__TILL__INFINITY As Int32 = -1


#Region "Properties"

        Private _PipeResults As Boolean = False
        Public ReadOnly Property PipeResults As Boolean
            Get
                Return Me._PipeResults
            End Get
        End Property


        Private _Visible As Boolean = False
        Public ReadOnly Property Visible As Boolean
            Get
                Return Me._Visible
            End Get
        End Property


        ''Public ReadOnly Property Result As String
        ''    Get
        ''        If Me.PipeResults AndAlso File.Exists(Me.ResultFile) Then
        ''            Try

        ''                Return File.ReadAllText(Me.ResultFile)

        ''            Catch ex As Exception
        ''                Me._LastError = ex
        ''                Debug.Print(ex.Message)
        ''            End Try

        ''        End If
        ''        Return String.Empty
        ''    End Get
        ''End Property



        ''Public ReadOnly Property ResultFile As String
        ''    Get
        ''        Return String.Format("{0}\{1}.cmdOutput", Application.StartupPath, get_guid_from_ObjectType(Me.GetType()))
        ''    End Get
        ''End Property
        Private _Result As String = String.Empty
        Public ReadOnly Property Result As String
            Get
                Return Me._Result
            End Get
        End Property


        Private _LastError As Exception = New Exception()
        Public ReadOnly Property LastError As Exception
            Get
                Return Me._LastError
            End Get
        End Property

#End Region

#Region "Methods"

        ''' <summary>
        ''' Execute a Command Prompt Code Like dir * /w
        ''' </summary>
        ''' <param name="args">dir * /w</param>
        ''' <param name="WorkingDirectory">The folder you want the command prompt to start from</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(ByVal args As String, ByVal WorkingDirectory As String) As Boolean

            Dim p As Process = New Process()
            Dim StartInfo As ProcessStartInfo = New ProcessStartInfo()
            If Not Me.Visible Then
                StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            Else
                StartInfo.WindowStyle = ProcessWindowStyle.Normal
            End If

            StartInfo.WorkingDirectory = WorkingDirectory
            StartInfo.FileName = "cmd.exe"
            REM StartInfo.Arguments = "/C " & args
            StartInfo.Arguments = String.Format("/C ""{0}""", args)
            REM /K wont matter here any more because if the output is redirected
            REM it will automatically close after execution
            REM Also it wont display a thing on the console.. it is just similar to piping

            REM I am using this since I will only be executing an exe which is cmd.exe
            StartInfo.RedirectStandardOutput = Me.PipeResults  REM So I can read from standard output
            StartInfo.UseShellExecute = Not Me.PipeResults REM If this is false redirection will work and also i can run only .exe file with this process only
            REM If ShellExecute is false it will always be visible

            ''StartInfo.RedirectStandardOutput = False REM So I can read from standard output
            ''StartInfo.UseShellExecute = True   REM If this is false redirection will work and also i can run only .exe file with this process only


            '' If Me.PipeResults Then StartInfo.Arguments &= " >" & Me.ResultFile

            p.StartInfo = StartInfo

            Try


                p.Start()

                If Me.PipeResults Then Me._Result = p.StandardOutput.ReadToEnd()
                p.WaitForExit()
                While Not p.HasExited

                    Application.DoEvents()
                End While

            Catch ex As Exception

                Me._LastError = ex

                Return False
            End Try

            Return True
        End Function


        ''' <summary>
        ''' Execute a Command Prompt Code Like dir * /w
        ''' </summary>
        ''' <param name="args">dir * /w</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(ByVal args As String) As Boolean

            Return Me.Execute(args, Application.StartupPath)

        End Function






        ''' <summary>
        ''' Executes commands using Interaction.Shell, uses cmd.exe from system32
        ''' </summary>
        ''' <param name="pCommand"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Execute(ByVal pCommand As String,
                                ByVal pTimeoutMilliSeconds As Int32) As Boolean

            Dim vMode As AppWinStyle = AppWinStyle.NormalFocus
            If Not Me.Visible Then vMode = AppWinStyle.Hide
            Try
                REM cmd.exe /K "Color 47"
                REM cmd.exe /C "Color 47"
                Dim pFullCommand As New StringBuilder("""C:\Windows\System32\cmd.exe"" /C")
                pFullCommand.Append(pCommand)
                Dim pResultFile As String = String.Format("{0}\ShellExcuteGenerated.log", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
                If Me.PipeResults Then pFullCommand.Append(String.Format(" >""{0}""", pResultFile))

                Dim procID As Int32 = Interaction.Shell(pFullCommand.ToString(), vMode, Me.PipeResults, pTimeoutMilliSeconds)
                Modules.basMain.MyLogFile.Print("ShellExecute: - Generated Process ID: " & procID)

                If Me.PipeResults AndAlso IO.File.Exists(pResultFile) Then Me._Result = IO.File.ReadAllText(pResultFile)
                EIO.DeleteFileIfExists(pResultFile)

                Return True
            Catch ex As Exception
                Me._LastError = ex
                Modules.basMain.MyLogFile.Print("ShellExecute: ", ex)
                Return False
            End Try
        End Function




        ''' <summary>
        ''' Executes commands using Interaction.Shell, uses cmd.exe from system32 and uses temporary file batch
        ''' </summary>
        ''' <param name="pCommand"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ForceExecute(ByVal pCommand As String,
                                Optional ByVal pTimeoutMilliSeconds As Int32 = WAIT__TILL__INFINITY) As Boolean

            Dim vMode As AppWinStyle = AppWinStyle.NormalFocus
            If Not Me.Visible Then vMode = AppWinStyle.Hide
            Try
                Dim pResultFile As String = String.Format("{0}\ShellExcuteGenerated.log", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
                Dim pBatchFile As String = String.Format("{0}\ShellExcuteTemp.bat", Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
                REM cmd.exe /K "Color 47"
                REM cmd.exe /C "Color 47"
                Dim FullCommand As New StringBuilder("""C:\Windows\System32\cmd.exe"" /C")
                FullCommand.Append(String.Format(" ""{0}""", pBatchFile))

                Dim UserCommand As New StringBuilder(pCommand)
                If Me.PipeResults Then UserCommand.Append(String.Format(" >""{0}""", pResultFile))

                IO.File.WriteAllText(pBatchFile, UserCommand.ToString())

                Dim procID As Int32 = Interaction.Shell(FullCommand.ToString(), vMode, Me.PipeResults, pTimeoutMilliSeconds)
                Modules.basMain.MyLogFile.Print("ShellExecute: - Generated Process ID: " & procID)

                If Me.PipeResults AndAlso IO.File.Exists(pResultFile) Then Me._Result = IO.File.ReadAllText(pResultFile)
                EIO.DeleteFileIfExists(pResultFile)
                EIO.DeleteFileIfExists(pBatchFile)

                Return True
            Catch ex As Exception
                Me._LastError = ex
                Modules.basMain.MyLogFile.Print("ShellExecute: ", ex)
                Return False
            End Try
        End Function













        ''' <summary>
        ''' Returns the Last result
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function ToString() As String
            Return Me.Result
        End Function


#End Region

    End Class

End Namespace