Imports System.Net
Imports EPRO.WaitAsyncMgr

''' <summary>
''' Handles Posting of Information to a URL
''' </summary>
''' <remarks></remarks>
Public Class EWebClient
    Implements IDisposable

    ''' <summary>
    ''' Holds a Format that is readable and usable for this class
    ''' </summary>
    ''' <remarks></remarks>
    Public Class WebClientURLFormat

        Private _URL As String = vbNullString
        Private _Parameters As Specialized.NameValueCollection = Nothing

        Public ReadOnly Property URL As String
            Get
                Return _URL
            End Get
        End Property

        Public ReadOnly Property URL_to_URi As Uri
            Get
                Return New Uri(URL)
            End Get
        End Property

        Public ReadOnly Property Parameters As Specialized.NameValueCollection
            Get
                Return _Parameters
            End Get
        End Property

        ''' <summary>
        ''' Get the URL and Parameters on a single string
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property CompressedFormat As String
            Get
                Dim Result As String = Me.URL & "?"

                For Each paramKey As String In Me.Parameters.AllKeys

                    Result &= String.Format("{0}={1}&", paramKey, Me.Parameters(paramKey))


                Next

                Return Result

            End Get
        End Property


        Public Sub New(ByVal URL As String)

            Dim SplitURL() As String = Split(URL, "?")
            If SplitURL IsNot Nothing Then
                Me._URL = SplitURL(0)
                If SplitURL.Count = 2 Then
                    REM There is parameters
                    Dim SplitParameters() As String = Split(SplitURL(1), "&")
                    If SplitParameters IsNot Nothing Then

                        If SplitParameters.Count > 0 Then
                            Dim retParam As Specialized.NameValueCollection = New Specialized.NameValueCollection(20)

                            For Each sptParam As String In SplitParameters
                                REM Since they go in pairs
                                Dim Param As String = vbNullString
                                Dim ParamValue As String = vbNullString
                                Dim Param_Val() As String = Split(sptParam, "=")
                                If Param_Val IsNot Nothing Then
                                    If Param_Val.Count > 0 Then Param = Param_Val(0)
                                    If Param_Val.Count > 1 Then ParamValue = Param_Val(1)

                                    retParam.Add(Param, ParamValue)


                                End If

                            Next

                            Me._Parameters = retParam

                        End If

                    End If

                End If



            End If




        End Sub

        Public Sub New(ByVal URL As String, ByVal Parameters As Specialized.NameValueCollection)
            Me._URL = URL
            Me._Parameters = Parameters

        End Sub




    End Class


    REM A special effect of this class is once it is connected to an address
    REM fetching any url from that address spins really
    REM Another advantage is that it doesnt hang when the pc is offline

    ''' <summary>
    ''' Maintain the waiting Process
    ''' </summary>
    ''' <remarks></remarks>
    Dim MyWait As WaitAsync = Nothing


    ''' <summary>
    ''' Use Locally ... and Monitor
    ''' </summary>
    ''' <remarks></remarks>
    Dim WithEvents client As New WebClient

    ''' <summary>
    ''' Last Request Start Time
    ''' </summary>
    ''' <remarks></remarks>
    Dim StartTime As Date = Now

    ''' <summary>
    ''' Keeps records of how many secs it takes to complete a request
    ''' </summary>
    ''' <remarks></remarks>
    Dim LastRequestUsedSecs As Long


    Public ReadOnly Property LastRequestUsedSeconds As Long
        Get
            Return Me.LastRequestUsedSecs
        End Get
    End Property


    ''Dim reqparm As New Specialized.NameValueCollection
    ''            reqparm.Add("qry", "version_server")

    ''' <summary>
    ''' User should have a copy of this to use that method
    ''' </summary>
    ''' <param name="ResponseText"></param>
    ''' <param name="ErrorOccurred"></param>
    ''' <remarks></remarks>
    Public Delegate Sub ReplyUploadValues(ByVal ResponseText As String, ByVal ErrorOccurred As Boolean)
    Dim UploadValueCallBackTarget As ReplyUploadValues

#Region "Dispose"
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If MyWait IsNot Nothing Then
                MyWait.Dispose()
                MyWait = Nothing
            End If
            If client IsNot Nothing Then
                client.Dispose()
                client = Nothing
            End If
        End If
    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub
#End Region


    ''' <summary>
    ''' Returns Only through here
    ''' </summary>
    ''' <param name="ResponseText"></param>
    ''' <param name="ErrorOccurred"></param>
    ''' <remarks></remarks>
    Private Sub UploadValuesReturning(ByVal ResponseText As String, ByVal ErrorOccurred As Boolean)

        Me.MyWait.endWaiting()

        Me.LastRequestUsedSecs = DateDiff(DateInterval.Second, StartTime, Now)

        Call UploadValueCallBackTarget(ResponseText, ErrorOccurred)

    End Sub

    Private Sub client_DownloadFileCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.AsyncCompletedEventArgs) Handles client.DownloadFileCompleted

    End Sub

    Private Sub client_DownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs) Handles client.DownloadProgressChanged
        Debug.Print(e.ProgressPercentage.ToString())
    End Sub

    Private Sub client_UploadProgressChanged(ByVal sender As Object, ByVal e As System.Net.UploadProgressChangedEventArgs) Handles client.UploadProgressChanged
        'Debug.Print(e.ProgressPercentage)
    End Sub



    Private Sub client_UploadValuesCompleted(ByVal sender As Object, ByVal e As System.Net.UploadValuesCompletedEventArgs) Handles client.UploadValuesCompleted
        REM You can be doing hand over for methods that calls async process and you want them to act as sync
        REM They will terminate and once the asycn has been answered , it will call back the main method indicating the step
        REM it stops.. The main method can use goto command to implement it

        Try


            If e.Cancelled Then UploadValuesReturning(vbNullString, False) : Return

            Call UploadValuesReturning(
                (New Text.UTF8Encoding).GetString(e.Result), False
                )
        Catch ex As Exception
            REM Error Occurred
            UploadValuesReturning(vbNullString, True)
        End Try


    End Sub





    ''' <summary>
    ''' Uses Post method to fetch URL. NOTE: You are responsible for empty result
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="CallBackFunc"></param>
    ''' <param name="reqparm"></param>
    ''' <remarks>NB: Simulating Sync for this Async Sub will not work for long reply request</remarks>
    Public Sub UploadValues(ByVal URL As String,
                            ByVal CallBackFunc As ReplyUploadValues,
                            Optional ByVal reqparm As Specialized.NameValueCollection = Nothing)


        REM Load Wait
        If Me.MyWait Is Nothing Then
            Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

        Else

            Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

        End If



        REM IF reqparm is nothing create a dummy data
        If IsNothing(reqparm) Then
            reqparm = New Specialized.NameValueCollection
            reqparm.Add("q", "dummy")
        End If

        Me.StartTime = Now
        Me.UploadValueCallBackTarget = CallBackFunc

        If client.IsBusy Or Not Network.CanConnectToInternet Then Call UploadValuesReturning(vbNullString, True) : Return

        Try
            ' URL = "http://softwares.eprocompany.com/LicenseController/LicenseController.php"

            client.UploadValuesAsync(New Uri(URL), "POST", reqparm)

        Catch ex As Exception

BUSYORERROR:
            Call UploadValuesReturning(vbNullString, True)

        End Try

    End Sub


    ''' <summary>
    ''' Uses Post method to fetch URL. NOTE: You are responsible for empty result
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="CallBackFunc"></param>
    ''' <param name="reqparm"></param>
    ''' <remarks>NB: Simulating Sync for this Async Sub will not work for long reply request</remarks>
    Public Sub UploadValues2(ByVal URL As String,
                            ByVal CallBackFunc As ReplyUploadValues,
                            Optional ByVal reqparm As Specialized.NameValueCollection = Nothing)


        REM Load Wait
        If Me.MyWait Is Nothing Then
            Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

        Else

            Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

        End If



        REM IF reqparm is nothing create a dummy data
        If IsNothing(reqparm) Then
            reqparm = New Specialized.NameValueCollection
            reqparm.Add("q", "dummy")
        End If

        Me.StartTime = Now
        Me.UploadValueCallBackTarget = CallBackFunc

        If client.IsBusy Or Not Network.CanConnectToInternet Then Call UploadValuesReturning(vbNullString, True) : Return

        Try

            client.UploadValuesAsync(New Uri(URL), "POST", reqparm)

        Catch ex As Exception

BUSYORERROR:
            REM Compromise
            Dim RETURN_VALUE As String = Network.GetPageSource(
                                                                               New WebClientURLFormat(URL, reqparm).CompressedFormat
                                                                                )
            If RETURN_VALUE.Trim() <> vbNullString Then
                REM Simulate the asy return here
                Call UploadValuesReturning(RETURN_VALUE, False)
            Else
                Call UploadValuesReturning(vbNullString, True)

            End If

        End Try

    End Sub


    ''' <summary>
    ''' Upload data using a POST Method Synchronously
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="reqparm"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadValuesSync(ByVal URL As String,
                            Optional ByVal reqparm As Specialized.NameValueCollection = Nothing) As String


        REM Load Wait
        If Me.MyWait Is Nothing Then
            Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

        Else

            Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

        End If

        REM IF reqparm is nothing create a dummy data
        If IsNothing(reqparm) Then
            reqparm = New Specialized.NameValueCollection
            reqparm.Add("q", "dummy")
        End If

        Me.StartTime = Now

        Dim Result As Byte(), RETURN_VALUE As String = vbNullString

        If client.IsBusy Or Not Network.CanConnectToInternet Then GoTo CONCLUDE

        Try

            REM Given the inconsistency of the SubRoutine
            REM Retry for some time to finalize result

            Dim RetryTimes As Byte = 3

            While RetryTimes > 0
BEGINFETCH:
                Result = client.UploadValues(New Uri(URL), "POST", reqparm)

                RETURN_VALUE = (New Text.UTF8Encoding).GetString(Result)


                REM --------------------------------------------
                If RETURN_VALUE.Trim <> vbNullString Then Exit While
                RetryTimes = CByte(RetryTimes - 1)
                Threading.Thread.Sleep(1000)
                If RetryTimes = 0 Then
                    REM se Old Method incase this new method fails
                    RETURN_VALUE = Network.GetPageSource(
                                                                                New WebClientURLFormat(URL, reqparm).CompressedFormat
                                                                                 )

                End If
                'GoTo BEGINFETCH
                REM -------------------------------------------------

            End While



            GoTo CONCLUDE

        Catch ex As Exception

BUSYORERROR:
            'Return vbNullString

        End Try

CONCLUDE:

        Me.MyWait.endWaiting()

        Me.LastRequestUsedSecs = DateDiff(DateInterval.Second, StartTime, Now)

        Return RETURN_VALUE

    End Function



    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="URL">It will download the document of this url as a file</param>
    ''' <param name="FileName">The File name it will save it with locally</param>
    ''' <returns></returns>
    ''' <remarks>Note Only the Readable Source from browser</remarks>
    Public Function DownloadFileFromURLSync(ByVal URL As String,
                            ByVal FileName As String) As Boolean


        REM Load Wait
        If Me.MyWait Is Nothing Then
            Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

        Else

            Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

        End If

        Me.StartTime = Now
        Dim RETURN_VALUE As Boolean = False
        If client.IsBusy Or Not Network.CanConnectToInternet Then GoTo CONCLUDE

        Try

            client.DownloadFile(URL, FileName)
            RETURN_VALUE = True

            GoTo CONCLUDE

        Catch ex As Exception

        End Try

CONCLUDE:

        Me.MyWait.endWaiting()

        Me.LastRequestUsedSecs = DateDiff(DateInterval.Second, StartTime, Now)

        Return RETURN_VALUE

    End Function



End Class
