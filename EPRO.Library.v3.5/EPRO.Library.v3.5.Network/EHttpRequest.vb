Option Strict Off


Imports System.Net
Imports System.IO
Imports System.Text
Imports EPRO.WaitAsyncMgr

Public Class EHttpRequest

    ''Public Class RequestState
    ''    ' This class stores the State of the request. 
    ''    Private BUFFER_SIZE As Integer = 1024
    ''    Public requestData As StringBuilder
    ''    Public BufferRead() As Byte
    ''    Public request As HttpWebRequest
    ''    Public response As HttpWebResponse
    ''    Public streamResponse As Stream

    ''    Public Sub New()
    ''        BufferRead = New Byte(BUFFER_SIZE) {}
    ''        requestData = New StringBuilder("")
    ''        request = Nothing
    ''        streamResponse = Nothing
    ''    End Sub 'New 
    ''End Class 'RequestState


    ''' <summary>
    ''' Maintain the waiting Process
    ''' </summary>
    ''' <remarks></remarks>
    Dim MyWait As WaitAsync = Nothing


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


    Public Enum RequestMethods
        [GET]
        [POST]
    End Enum


    Public Property RequestMethod As RequestMethods

#Region "Constructors"

    Sub New()
        Me.RequestMethod = RequestMethods.GET
    End Sub

#End Region



#Region "Exposed"

    ''' <summary>
    ''' Upload Information Synchronously
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadValues(ByVal URL As String) As String

        Dim strReply As String = ""


        Try
            Dim objHttpRequest As HttpWebRequest = HttpWebRequest.Create(URL)

            REM Just for testing
            'objHttpRequest.Method = "POST"


            Dim objHttpResponse As HttpWebResponse = objHttpRequest.GetResponse
            Dim objStrmReader As New StreamReader(objHttpResponse.GetResponseStream)

            strReply = objStrmReader.ReadToEnd()

        Catch ex As Exception
            strReply = ""
        End Try



        Return strReply

    End Function

    ''' <summary>
    ''' Upload Information Synchronously
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadValues(ByVal URL As String, ByVal reqparm As Specialized.NameValueCollection) As String

        Dim strReply As String = ""


        Try
            Dim objHttpRequest As HttpWebRequest = HttpWebRequest.Create(
                 New EWebClient.WebClientURLFormat(URL, reqparm).CompressedFormat
                 )

            REM Just for testing
            'objHttpRequest.Method = "POST"


            Dim objHttpResponse As HttpWebResponse = objHttpRequest.GetResponse
            Dim objStrmReader As New StreamReader(objHttpResponse.GetResponseStream)

            strReply = objStrmReader.ReadToEnd()

        Catch ex As Exception
            strReply = ""
        End Try



        Return strReply

    End Function



    ''' <summary>
    ''' Upload Information Asynchronously
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="CallBackFunc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadValuesAsync(ByVal URL As String,
                                      ByVal CallBackFunc As ReplyUploadValues) As Boolean


        REM Load Wait
        If Me.MyWait Is Nothing Then
            Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

        Else

            Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

        End If

        Me.StartTime = Now
        Me.UploadValueCallBackTarget = CallBackFunc


        Try

            'Dim myRequestState As New RequestState()

            Dim objHttpRequest As HttpWebRequest = HttpWebRequest.Create(URL)

            '  myRequestState.request = objHttpRequest

            objHttpRequest.BeginGetResponse(
               New AsyncCallback(AddressOf UploadValuesAsyncReturn), objHttpRequest
               )




        Catch ex As Exception

            Call UploadValuesReturning("Error Occurred Calling Upload Function", True)

            Return False
        End Try



        Return True

    End Function



    ''' <summary>
    ''' Upload Information Asynchronously
    ''' </summary>
    ''' <param name="URL"></param>
    ''' <param name="CallBackFunc"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UploadValuesAsync(ByVal URL As String,
                                      ByVal CallBackFunc As ReplyUploadValues,
                                       ByVal reqparm As Specialized.NameValueCollection,
                                       Optional ByVal isQuiteMode As Boolean = False) As Boolean


        If Not isQuiteMode Then
            REM Load Wait
            If Me.MyWait Is Nothing Then
                Me.MyWait = New WaitAsync(, String.Format("Downloading info from: [{0}]", URL))

            Else

                Me.MyWait.startWaiting(, String.Format("Downloading info from: [{0}]", URL))

            End If
        End If


        Me.StartTime = Now
        Me.UploadValueCallBackTarget = CallBackFunc


        Try

            'Dim myRequestState As New RequestState()

            Dim objHttpRequest As HttpWebRequest = HttpWebRequest.Create(
                New EWebClient.WebClientURLFormat(URL, reqparm).CompressedFormat
                )

            '  myRequestState.request = objHttpRequest


            objHttpRequest.BeginGetResponse(
               New AsyncCallback(AddressOf UploadValuesAsyncReturn), objHttpRequest
               )




        Catch ex As Exception

            Call UploadValuesReturning("Error Occurred Calling Upload Function", True)

            Return False
        End Try



        Return True

    End Function





#End Region



    ''' <summary>
    ''' Asynchronous Operation Returning
    ''' </summary>
    ''' <param name="asynchronousResult"></param>
    ''' <remarks></remarks>
    Private Sub UploadValuesAsyncReturn(ByVal asynchronousResult As IAsyncResult)

        Try
            Dim strReply As String

            Dim myHttpWebRequest As HttpWebRequest = CType(asynchronousResult.AsyncState, HttpWebRequest)

            Dim response As HttpWebResponse = CType(
                                                    myHttpWebRequest.EndGetResponse(asynchronousResult), 
                                                    HttpWebResponse
                                                    )

            Dim objStrmReader As New StreamReader(response.GetResponseStream)

            strReply = objStrmReader.ReadToEnd()

            Me.UploadValuesReturning(strReply, False)

        Catch ex As Exception
            REM Error Occurred. The remote [epro.....] could not be resolved
            Me.UploadValuesReturning(ex.Message, True)

        End Try


    End Sub





    ''' <summary>
    ''' Returns Only through here
    ''' </summary>
    ''' <param name="ResponseText"></param>
    ''' <param name="ErrorOccurred"></param>
    ''' <remarks></remarks>
    Private Sub UploadValuesReturning(ByVal ResponseText As String, ByVal ErrorOccurred As Boolean)

        If Me.MyWait IsNot Nothing Then Me.MyWait.endWaiting()

        Me.LastRequestUsedSecs = DateDiff(DateInterval.Second, StartTime, Now)

        Call UploadValueCallBackTarget(ResponseText, ErrorOccurred)

    End Sub


End Class
