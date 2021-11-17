Imports System.IO
Imports System.Threading
Imports System.Windows.Forms

Public Class ShellWait

    REM My Shell and wait process is failing
    REM First you should know not all process supports cmd
    REM If you start some executables .. once the launch is completed the handle is returned
    REM So first this should be used when the handle is held until the process completes

    ''' <summary>
    ''' Using Command Prompts Title to Pick Process that will return only on exit to command prompt
    ''' </summary>
    ''' <param name="filePath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WaitProcessToExit(ByVal filePath As String,
                                              ByVal Style As AppWinStyle) As Boolean

        Try
            Dim tmpBatFile As String = String.Format(
                                                        "{0}\ShellAndWait.bat",
                                                        FileIO.FileSystem.GetParentPath(filePath)
                                                        )


            Randomize()
            Dim sCaption As String = "Please Wait -" & CInt(Rnd() * 1000)

            REM Create a Title Window
            File.WriteAllText(tmpBatFile, String.Format("TITLE {0}{1}", sCaption, vbCrLf))

            REM Launches the previous bat in this bat
            File.AppendAllText(tmpBatFile, String.Format("""{0}""", filePath))


            Interaction.Shell(tmpBatFile, Style)
            Application.DoEvents()

            Thread.Sleep(200)

            Do While True

                For Each Proc As Process In Process.GetProcessesByName("cmd")


                    If Proc.MainWindowTitle <> vbNullString AndAlso Proc.MainWindowTitle.IndexOf(sCaption) >= 0 Then
                        GoTo CONTINUE_LOOP
                    End If

                Next

                Exit Do
CONTINUE_LOOP:

                Thread.Sleep(50)
            Loop

        Catch ex As Exception
            Return False
        End Try


        REM Just tarry a little longer
        Thread.Sleep(200)

        Return True

    End Function



    '' '' ''#Region "Others"

    '' '' ''    ''' <summary>
    '' '' ''    ''' Shell a bat File and return once the execution finishes [Tested ... More Efficient ]
    '' '' ''    ''' </summary>
    '' '' ''    ''' <param name="filePath"></param>
    '' '' ''    ''' <returns >Returns true on success </returns>
    '' '' ''    ''' <remarks></remarks>
    '' '' ''    Public Shared Function ShellAndWait(ByVal filePath As String,
    '' '' ''                                        ByVal Style As AppWinStyle) As Boolean
    '' '' ''        Try
    '' '' ''            Dim tmpBatFile As String = String.Format(
    '' '' ''                                                        "{0}\ShellAndWait.bat",
    '' '' ''                                                        FileIO.FileSystem.GetParentPath(filePath)
    '' '' ''                                                        )


    '' '' ''            Randomize()
    '' '' ''            Dim sCaption As String = "Please Wait -" & CInt(Rnd() * 1000)

    '' '' ''            REM Create a Title Window
    '' '' ''            File.WriteAllText(tmpBatFile, String.Format("TITLE {0}{1}", sCaption, vbCrLf))

    '' '' ''            REM Launches the previous bat in this bat
    '' '' ''            File.AppendAllText(tmpBatFile, String.Format("""{0}""", filePath))


    '' '' ''            Shell(tmpBatFile, Style)
    '' '' ''            Application.DoEvents()

    '' '' ''            Thread.Sleep(100)

    '' '' ''            Do While GetWindowHandle__UsingCaption("Administrator:  " & sCaption).ToInt32 <> 0 Or
    '' '' ''                GetWindowHandle__UsingCaption(sCaption).ToInt32 <> 0




    '' '' ''                Application.DoEvents()
    '' '' ''                'Thread.Sleep(50)

    '' '' ''            Loop


    '' '' ''            REM Wait more for the process to really close
    '' '' ''            Thread.Sleep(1000)

    '' '' ''            Return True


    '' '' ''        Catch ex As Exception

    '' '' ''        End Try


    '' '' ''        Return False

    '' '' ''    End Function


    '' '' ''    ''' <summary>
    '' '' ''    ''' Shell a bat File and return once the execution finishes
    '' '' ''    ''' </summary>
    '' '' ''    ''' <param name="filePath"></param>
    '' '' ''    ''' <returns >Returns the sCaption Window You should wait for</returns>
    '' '' ''    ''' <remarks></remarks>
    '' '' ''    Public Shared Function ShellAndWait(ByVal filePath As String) As String
    '' '' ''        Dim tmpBatFile As String = String.Format(
    '' '' ''                                                    "{0}\ShellAndWait.bat",
    '' '' ''                                                    FileIO.FileSystem.GetParentPath(filePath)
    '' '' ''                                                    )

    '' '' ''        Randomize()
    '' '' ''        Dim sCaption As String = "Please Wait -" & CInt(Rnd() * 1000)

    '' '' ''        REM Create a Title Window
    '' '' ''        File.WriteAllText(tmpBatFile, String.Format("TITLE {0}{1}", sCaption, vbCrLf))

    '' '' ''        REM Launches the previous bat in this bat
    '' '' ''        File.AppendAllText(tmpBatFile, String.Format("""{0}""", filePath))


    '' '' ''        Dim proc As Process = Process.Start(tmpBatFile)
    '' '' ''        proc.WaitForExit(1000)
    '' '' ''        'Dim hwnd As Int32 = Shell(tmpBatFile)

    '' '' ''        'Debug.Print(GetWindowCaptionText(proc.MainWindowHandle))
    '' '' ''        Try
    '' '' ''            Return proc.MainWindowTitle
    '' '' ''        Catch ex As Exception

    '' '' ''        End Try


    '' '' ''        Return vbNullString

    '' '' ''    End Function


    '' '' ''    Dim clsShellProWaitThread As Thread
    '' '' ''    Dim clsShellProWaitThreadMutex As Boolean = False
    '' '' ''    Public Sub WaitProcessToExitAsync(ByVal procCaption As String,
    '' '' ''                                      ByVal addressToCallOnExit As clsDelegates.delegateNoParam
    '' '' ''                                      )

    '' '' ''        REM Class is in use
    '' '' ''        If clsShellProWaitThreadMutex Then Return

    '' '' ''        clsShellProWaitThread = New Thread(New ThreadStart(Sub() WaitProcessToExitAsyncThread(procCaption, addressToCallOnExit))) With {.IsBackground = True}
    '' '' ''        clsShellProWaitThread.Start()
    '' '' ''        Me.clsShellProWaitThreadMutex = True

    '' '' ''    End Sub

    '' '' ''    Public Sub WaitProcessToExitAsync(ByVal proc As Process,
    '' '' ''                                      ByVal addressToCallOnExit As clsDelegates.delegateNoParam
    '' '' ''                                      )

    '' '' ''        REM Class is in use
    '' '' ''        If clsShellProWaitThreadMutex Then Return

    '' '' ''        clsShellProWaitThread = New Thread(New ThreadStart(Sub() WaitProcessToExitAsyncThread(proc, addressToCallOnExit))) With {.IsBackground = True}
    '' '' ''        clsShellProWaitThread.Start()

    '' '' ''        Me.clsShellProWaitThreadMutex = True

    '' '' ''    End Sub

    '' '' ''    Private Sub WaitProcessToExitAsyncThread(ByVal proc As Process,
    '' '' ''                                              ByVal addressToCallOnExit As clsDelegates.delegateNoParam)

    '' '' ''        Try
    '' '' ''            Do Until proc.HasExited

    '' '' ''                Thread.Sleep(200)
    '' '' ''            Loop

    '' '' ''            Call addressToCallOnExit()

    '' '' ''        Catch ex As Exception

    '' '' ''        End Try

    '' '' ''        Me.clsShellProWaitThreadMutex = False
    '' '' ''    End Sub




    '' '' ''    Private Sub WaitProcessToExitAsyncThread(ByVal procCaption As String,
    '' '' ''                                             ByVal addressToCallOnExit As clsDelegates.delegateNoParam)

    '' '' ''        Try
    '' '' ''            Do While isWindowOn(procCaption)

    '' '' ''                Thread.Sleep(1000)
    '' '' ''            Loop

    '' '' ''            Call addressToCallOnExit()

    '' '' ''        Catch ex As Exception

    '' '' ''        End Try

    '' '' ''        Me.clsShellProWaitThreadMutex = False

    '' '' ''    End Sub

    '' '' ''#End Region


End Class
