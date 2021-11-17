Imports System.Printing


Namespace MicrosoftOS

    Public Class PrintJobDescriptor


#Region "Constructors"


        Friend Sub New(pJob As PrintSystemJobInfo
                       )
            _______PrinterName = pJob.HostingPrintQueue.Name
            Me._______JobID = pJob.JobIdentifier
            _______TotalPages = pJob.NumberOfPages

            Me._______LastUpdatedTime = Now


            Me._______IsPrintedIndicated = pJob.IsPrinted
            Me._______IsPrintingIndicated = pJob.IsPrinting
            Me._______IsSpoolingIndicated = pJob.IsSpooling

            Me._______PortName = String.Empty
            Me._______MachineName = String.Empty
            Me._______PrintProcessor = String.Empty



        End Sub




#End Region



#Region "Consts and Enums"

        Friend Const OVERDUE___INTERVAL___SECS As Int32 = 5

#End Region



#Region "Properties"




        Private _______PrinterName As String
        Public ReadOnly Property PrinterName As String
            Get
                Return Me._______PrinterName
            End Get
        End Property


        Private _______TotalPages As Int32

        Public ReadOnly Property TotalPages As Int32
            Get
                Return _______TotalPages
            End Get
        End Property



        Private _______JobID As Int32

        Public ReadOnly Property JobID As Int32
            Get
                Return _______JobID
            End Get
        End Property




        Private _______IsPrintedIndicated As Boolean
        Private _______IsPrintingIndicated As Boolean
        Private _______IsSpoolingIndicated As Boolean



        Private _______MachineName As String
        Public ReadOnly Property MachineName As String
            Get
                Return Me._______MachineName
            End Get
        End Property

        Private _______PortName As String
        Public ReadOnly Property PortName As String
            Get
                Return Me._______PortName
            End Get
        End Property

        Private _______PrintProcessor As String
        Public ReadOnly Property PrintProcessor As String
            Get
                Return Me._______PrintProcessor
            End Get
        End Property


        Private _______LastUpdatedTime As Date

        Friend ReadOnly Property LastUpdatedTime As Date
            Get
                Return _______LastUpdatedTime
            End Get
        End Property



        Public ReadOnly Property JobUniqueID As String
            Get
                Return ConstructUniqueID(Me.PrinterName, Me.JobID)
            End Get
        End Property


#End Region



#Region "Methods"

        Friend Sub Update(pJob As PrintSystemJobInfo)

            If Not Me._______IsPrintingIndicated Then Me._______IsPrintingIndicated = pJob.IsPrinting
            If Not Me._______IsPrintedIndicated Then Me._______IsPrintedIndicated = pJob.IsPrinted

            If Not Me._______IsSpoolingIndicated Then Me._______IsSpoolingIndicated = pJob.IsSpooling


            If Me._______IsPrintingIndicated OrElse Me._______IsPrintedIndicated Then Me._______TotalPages = pJob.NumberOfPages

            Me._______LastUpdatedTime = Now

            If Me._______PortName = String.Empty Then Me._______PortName = pJob.HostingPrintQueue.QueuePort.Name
            If Me._______PrintProcessor = String.Empty Then Me._______PrintProcessor = pJob.HostingPrintQueue.QueuePrintProcessor.Name
            If Me._______MachineName = String.Empty Then Me._______MachineName = pJob.HostingPrintServer.Name

            'Debug.Print(String.Empty)
            'Debug.Print("pJob.IsSpooling: " & pJob.IsSpooling)
            'Debug.Print("pJob.IsPrinting: " & pJob.IsPrinting)
            'Debug.Print("pJob.IsPrinted: " & (pJob.IsPrinted))
            'Debug.Print(String.Empty)


        End Sub

        Friend Function IsJobCompleted() As Boolean
            Return (
                    Me._______IsPrintingIndicated OrElse Me._______IsPrintedIndicated
                    ) AndAlso Me._______IsSpoolingIndicated
        End Function


        Friend Function IsOverDueForUpdate() As Boolean
            Return DateDiff(DateInterval.Second, Me.LastUpdatedTime, Now) > OVERDUE___INTERVAL___SECS
        End Function

        Friend Shared Function ConstructUniqueID(pPrinterName As String, pJobID As Int32) As String
            Return String.Format("{0}-{1}", pPrinterName, pJobID)

        End Function

#End Region







    End Class

End Namespace