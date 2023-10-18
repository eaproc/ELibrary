Imports System.IO


Namespace AppConfigurations

    Public Class CSVReader
        Implements IDisposable



#Region "Constructors"

        Public Sub New(csvFilePath As String,
                       encode As System.Text.Encoding,
                       ByVal pSplitOptions As StringSplitOptions,
                     Optional ByVal RowDelimiters As String = vbCrLf,
                     Optional ByVal CellDelimiter As String = ",")

            If Not File.Exists(csvFilePath) Then Return

            Try

                Dim f As String = File.ReadAllText(csvFilePath, encode)


                If (f.Trim() = string.Empty) Then Return

                REM Parse File
                Me.__RawRows = f.Split(New String() {RowDelimiters}, StringSplitOptions.RemoveEmptyEntries).ToList()

                If (RawRows.Count > 0) Then


                    REM Date.Now process the files using the real delimiter
                    For Each Line As String In RawRows

                        REM get keys values
                        Dim Cells As String() = Line.Split(New String() {CellDelimiter}, pSplitOptions)

                        Me.__Rows.Add(Cells)

                    Next

                End If

            Catch ex As Exception
                Debug.Print(ex.Message)

            End Try

            Me.__IsValid = True
        End Sub

        ''' <summary>
        ''' Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        ''' </summary>
        ''' <param name="csvFilePath"></param>
        ''' <param name="RowDelimiters"></param>
        ''' <param name="CellDelimiter"></param>
        ''' <remarks></remarks>
        Public Sub New(csvFilePath As String,
                        ByVal pSplitOptions As StringSplitOptions,
                       Optional ByVal RowDelimiters As String = vbCrLf,
                       Optional ByVal CellDelimiter As String = ",")

            Me.New(csvFilePath, System.Text.Encoding.Default, pSplitOptions, RowDelimiters, CellDelimiter)


        End Sub

        ''' <summary>
        ''' Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default. Removes empty cells 
        ''' </summary>
        ''' <param name="csvFilePath"></param>
        ''' <param name="RowDelimiters"></param>
        ''' <param name="CellDelimiter"></param>
        ''' <remarks></remarks>
        Public Sub New(csvFilePath As String,
                       Optional ByVal RowDelimiters As String = vbCrLf,
                       Optional ByVal CellDelimiter As String = ",")

            Me.New(csvFilePath, System.Text.Encoding.Default, StringSplitOptions.RemoveEmptyEntries, RowDelimiters, CellDelimiter)


        End Sub


#End Region



#Region "Properties"

        Private __IsValid As Boolean = False
        ''' <summary>
        ''' Checks only if file is valid. Use has rows to know if the file contains data
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsValid As Boolean
            Get
                Return Me.__IsValid
            End Get
        End Property



        Private __Rows As List(Of String()) = New List(Of String())()
        ''Public ReadOnly Property Rows(Optional IgnoreFirstLine As Boolean = False) As List(Of String())
        ''    Get
        ''        If Me.hasRows AndAlso IgnoreFirstLine Then
        ''            If Me.__Rows.Count > 1 Then
        ''                Dim cpy As String()() = Array.CreateInstance(GetType(String()), Me.__Rows.Count - 1).Cast(Of String()).ToArray()
        ''                Me.__Rows.CopyTo(cpy, 1)
        ''                Return cpy.ToList()
        ''            End If

        ''            Return New List(Of String())
        ''        End If

        ''            Return Me.__Rows
        ''    End Get
        ''End Property

        Public ReadOnly Property Rows As List(Of String())
            Get
                Return Me.__Rows
            End Get
        End Property



        Private __RawRows As List(Of String) = New List(Of String)()

        Public ReadOnly Property RawRows As List(Of String)
            Get
                Return Me.__RawRows
            End Get
        End Property


        Public ReadOnly Property Count As Int32
            Get
                Return Me.__Rows.Count
            End Get
        End Property

        Public ReadOnly Property hasRows As Boolean
            Get
                Return Me.__Rows.Count > 0
            End Get
        End Property


#End Region




#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not Me.disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                    Me.__Rows = Nothing
                    Me.__RawRows = Nothing
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
            End If
            Me.disposedValue = True
        End Sub

        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return Me.disposedValue
            End Get
        End Property

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