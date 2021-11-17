Namespace Objects



    ''' <summary>
    ''' Decimal is smaller in range but has more accuracy than double
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EDecimal

        ''' <summary>
        ''' Converts An Object to Double
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As String) As Decimal
            Try
                If obj Is Nothing OrElse obj.Trim() = String.Empty Then Return 0
                Return CDec(obj)

            Catch ex As Exception
                Return 0

            End Try
        End Function


        ''' <summary>
        ''' Converts objects to double
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Object) As Decimal

            If obj Is Nothing Then Return 0
            Dim d As Double = CDec(obj)
            If TypeOf obj Is String Then Return valueOf(CType(obj, String))
            If TypeOf obj Is DBNull Then Return 0D
            If TypeOf obj Is Double AndAlso CDbl(obj) <= Decimal.MaxValue Then Return CDec(obj)
            If TypeOf obj Is Long Then Return (CType(obj, Long))
            If TypeOf obj Is Int32 Then Return CInt(obj)
            If TypeOf obj Is Int16 Then Return CInt(obj)
            If TypeOf obj Is Boolean Then Return EShort.valueOf(CType(obj, Boolean))
            If TypeOf obj Is Decimal Then Return CDec(obj)

            Return 0    REM Cant convert this

        End Function



    End Class

End Namespace