Imports System.IO
Imports System.Net


Imports EPRO.Library.v3._5


''' <summary>
''' Downloads File Using WebClient. Reusable. 
''' -   It doesn't have resume capability
''' -   It doesn't save downloads partially on disk
''' </summary>
''' <remarks></remarks>
Public Class WebClientDownloader
    Implements IDisposable



#Region "Constructors"

    ''' <summary>
    ''' Logical Ignore Download will check if the file with same name exists. It will not download it again.
    ''' </summary>
    ''' <param name="pUseLogicalIgnoreDownloadIfExist">Indicates if Download should be counted as downloaded if it already exists</param>
    ''' <param name="pDeleteDownloadedFileAfterRaisingEvent">It will try to force delete the downloaded file if indicated after raising event</param>
    ''' <remarks>Throws exception if creation of Directory Fails</remarks>
    Public Sub New(ByVal pDestinationFolder As String,
                   Optional ByVal pUseLogicalIgnoreDownloadIfExist As Boolean = True,
                   Optional ByVal pDeleteDownloadedFileAfterRaisingEvent As Boolean = False
                                                                                )
        Me._____Downloader = New ___WebClientSubClassed()
        Me._____DestinationFolder = pDestinationFolder
        Me._____UseLogicalIgnoreDownloadIfExist = pUseLogicalIgnoreDownloadIfExist
        Me._____DeleteDownloadedFileAfterRaisingEvent = pDeleteDownloadedFileAfterRaisingEvent

        ' Throws exception if creation of Directory Fails
        If Not Directory.Exists(Me._____DestinationFolder) Then Directory.CreateDirectory(Me._____DestinationFolder)


    End Sub



#End Region



#Region "Properties"

    Private _____DestinationFolder As String
    Private _____DownloadedFileName As String
    Private _____DownloadedURL As String
    Private _____UseLogicalIgnoreDownloadIfExist As Boolean
    Private _____DeleteDownloadedFileAfterRaisingEvent As Boolean



    Public ReadOnly Property DownloadedURL As String
        Get
            Return _____DownloadedURL
        End Get

    End Property

    Private ReadOnly Property DownloadedFileFullPath As String
        Get
            Return String.Format("{0}\{1}", Me._____DestinationFolder, Me._____DownloadedFileName)
        End Get
    End Property


    Private WithEvents _____Downloader As ___WebClientSubClassed

    Private _____IsBusy As Boolean
    Public ReadOnly Property IsBusy As Boolean
        Get
            Return Me._____IsBusy
        End Get
    End Property



    Private _____IsCanceled As Boolean
    Public ReadOnly Property IsCanceled As Boolean
        Get
            Return Me._____IsCanceled
        End Get
    End Property


    Private _____DownloadedFileTargetSize As Long


    ''' <summary>
    ''' It only self terminates if it discovers the file name already exists
    ''' </summary>
    ''' <remarks></remarks>
    Private _____IsSelfTerminated As Boolean


#End Region



#Region "Events"


    Public Event DownloadCompleted(pFileFullPath As String)
    Public Event DownloadFailed(pReason As String)
    Public Event DownloadProgress(pMaxValue As Long, pCurrentValue As Long)





#End Region



#Region "Private Class"
    Private Class ___WebClientSubClassed
        Inherits WebClient

        Private _____responseURI As Uri
        Public ReadOnly Property ResponseURI As Uri
            Get
                Return Me._____responseURI
            End Get
        End Property



        Protected Overrides Function GetWebResponse(request As WebRequest) As WebResponse
            Dim response = MyBase.GetWebResponse(request)

            Me._____responseURI = response.ResponseUri

            Return response

        End Function

        Protected Overrides Function GetWebResponse(request As WebRequest, result As IAsyncResult) As WebResponse
            Dim response = MyBase.GetWebResponse(request, result)
            Me._____responseURI = response.ResponseUri

            Return response
        End Function


    End Class
#End Region




#Region "Methods"


    ''' <summary>
    ''' Throws Exceptions. It extracts file name from URL
    ''' </summary>
    ''' <param name="pURL"></param>
    ''' <remarks></remarks>
    Public Sub Download(pURL As String)
        Me.Download(pURL, String.Empty)
    End Sub


    ''' <summary>
    ''' Throws Exceptions. Downloads Data Asynchronous
    ''' </summary>
    ''' <param name="pURL"></param>
    ''' <remarks></remarks>
    Public Sub Download(pURL As String, pFileName As String)
        If Me._____IsBusy OrElse Me.disposedValue Then Throw New InvalidOperationException("You can't download anything when busy or Disposed") : Return

        Me._____DownloadedFileName = pFileName
        Me._____IsSelfTerminated = False

        Try

         
            Me._____Downloader.DownloadDataAsync(New Uri(pURL))
            Me._____IsBusy = True
            Me._____DownloadedURL = pURL


            If Me._____DownloadedFileName <> String.Empty Then
                Call Me.OnRaise___FileNameSET()
            End If

            RaiseEvent DownloadProgress(0, 0)


        Catch ex As Exception
            Program.Logger.Print(ex)
            Me.OnRaise___DownloadFailed(ex.Message)

        End Try

    End Sub


    Public Sub CancelDownload()
        If Not Me.IsBusy Then Return
        Me._____Downloader.CancelAsync()
    End Sub



    Private Sub OnRaise___DownloadFailed(pReason As String)
        RaiseEvent DownloadFailed(pReason)
        Me._____IsBusy = False
        Program.Logger.Print(pReason)
    End Sub

    Private Sub OnRaise___DownloadCompleted()
        Try

            RaiseEvent DownloadCompleted(Me.DownloadedFileFullPath)

            If Me._____DeleteDownloadedFileAfterRaisingEvent Then EIO.DeleteFileIfExists(Me.DownloadedFileFullPath)
        Catch ex As Exception
            '   Ignore error here, not important. Atleast not to this class
            Program.Logger.Print(ex)
        End Try

        Me._____IsBusy = False

    End Sub

    Private Sub OnRaise___FileNameSET()
        Program.Logger.Print("File Name SET: " & Me._____DownloadedFileName)

        If Me._____UseLogicalIgnoreDownloadIfExist AndAlso File.Exists(Me.DownloadedFileFullPath) Then
            Me._____IsSelfTerminated = True
            Program.Logger.Print("Canceling Download Automatically ")

            Me.CancelDownload()

        End If


    End Sub

#End Region





#Region "WebClients Events"

    Private Sub _____Downloader_DownloadDataCompleted(sender As Object, e As DownloadDataCompletedEventArgs) Handles _____Downloader.DownloadDataCompleted
        Me._____IsCanceled = e.Cancelled


        If IsNothing(e.Error) Then

            Program.Logger.Print("No Error Under Download Completed")

            Dim _FileDownloaded As FileStream 
            Try


                _FileDownloaded = New FileStream(
                                                Me.DownloadedFileFullPath,
                                                FileMode.Create, FileAccess.Write
                                                )

                With _FileDownloaded
                    _FileDownloaded.Write(e.Result, 0, e.Result.Length)
                    .Flush()
                    .Close()
                End With

                Me.OnRaise___DownloadCompleted()

            Catch ex As Exception
                Me.OnRaise___DownloadFailed("Error saving file downloaded. .. " & ex.Message)
                Program.Logger.Print(Me.DownloadedFileFullPath, ex)
            End Try
            _FileDownloaded = Nothing

        ElseIf Me._____IsSelfTerminated Then
            Program.Logger.Print("Download Is Self Terminated")

            Me.OnRaise___DownloadCompleted()
            Me._____IsSelfTerminated = False

        ElseIf Me.IsCanceled Then
            Me.OnRaise___DownloadFailed("User canceled download. " & e.Error.Message)

        Else


            Me.OnRaise___DownloadFailed("Details: " & e.Error.Message)

        End If


    End Sub

    Private Sub _____Downloader_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs) Handles _____Downloader.DownloadProgressChanged
        RaiseEvent DownloadProgress(e.TotalBytesToReceive, e.BytesReceived)
        Me._____DownloadedFileTargetSize = e.TotalBytesToReceive

        If Me._____DownloadedFileName = String.Empty Then
            Program.Logger.Print("SETING FILE NAME Under Progress")

            Me._____DownloadedFileName = EIO.getFileName(Me._____Downloader.ResponseURI)
            Me.OnRaise___FileNameSET()
        End If


    End Sub



#End Region






#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

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
                If Me._____Downloader IsNot Nothing Then Me._____Downloader.Dispose()
                Me._____Downloader = Nothing
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
