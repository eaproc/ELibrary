Option Strict On
Option Explicit On

Namespace Objects

    Public NotInheritable Class EInt

        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As String) As Int32
            Try

                If IsNothing(obj) Then Return 0
                If obj = "" Then Return 0
                If obj.Trim = "" Then Return 0

                Return CInt(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function


        ''' <summary>
        ''' Converts An Object to int32
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As System.DBNull) As Int32

            Return 0

        

        End Function





        ''' <summary>
        ''' Converts An Object to int32
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Double) As Int32
            Try

                Return CInt(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function

        ''' <summary>
        ''' Converts An Object to int32
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Decimal) As Int32
            Try

                Return CInt(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function




        ''' <summary>
        ''' Converts objects to integer
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Object) As Int32

            If obj Is Nothing Then Return 0

            If TypeOf obj Is String Then Return valueOf(CType(obj, String))
            If TypeOf obj Is DBNull Then Return valueOf(CType(obj, DBNull))
            If TypeOf obj Is Double Then Return valueOf(CType(obj, Double))
            If TypeOf obj Is Long Then Return valueOf(CType(obj, Long))
            If TypeOf obj Is Int32 Then Return CInt(obj)
            If TypeOf obj Is Int16 Then Return CInt(obj)
            If TypeOf obj Is Decimal Then Return valueOf(CType(obj, Decimal))
            If TypeOf obj Is Boolean Then Return EShort.valueOf(CType(obj, Boolean))

            Return 0    REM Cant convert this

        End Function



    End Class

End Namespace