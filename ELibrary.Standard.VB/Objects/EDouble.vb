Namespace Objects

    ''' <summary>
    ''' Double is larger in range represented but with lower accuracy than decimal
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EDouble

        ''' <summary>
        ''' Converts An Object to Double
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function ValueOf(ByVal obj As String) As Double
            Try
                If obj Is Nothing OrElse obj.Trim() = string.Empty Then Return 0
                Return CDbl(obj)

            Catch ex As Exception
                Return 0.0

            End Try
        End Function


        ''' <summary>
        ''' Converts objects to double
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function ValueOf(ByVal obj As Object) As Double

            If obj Is Nothing Then Return 0

            If TypeOf obj Is String Then Return ValueOf(CType(obj, String))
            If TypeOf obj Is DBNull Then Return 0.0#
            If TypeOf obj Is Double Then Return CDbl(obj)
            If TypeOf obj Is Long Then Return (CType(obj, Long))
            If TypeOf obj Is Int32 Then Return CInt(obj)
            If TypeOf obj Is Int16 Then Return CInt(obj)
            If TypeOf obj Is Boolean Then Return EShort.ValueOf(CType(obj, Boolean))
            If TypeOf obj Is Decimal Then Return CDbl(obj)

            Return 0    REM Cant convert this

        End Function



    End Class

End Namespace