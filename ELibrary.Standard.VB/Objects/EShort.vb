

Imports ELibrary.Standard.VB.Modules

Namespace Objects

    Public NotInheritable Class EShort

        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As String) As Int16
            Try

                If basExtensions.IsNothing(obj) Then Return 0
                If obj = "" Then Return 0
                If obj.Trim = "" Then Return 0

                Return CShort(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function


        ''' <summary>
        ''' Converts An Object to short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As System.DBNull) As Int16
            Return 0
        End Function

        ''' <summary>
        ''' Converts An Object to Short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Double) As Int16
            Try

                Return CShort(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function


        ''' <summary>
        ''' Converts An Object to Short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Decimal) As Int16
            Try

                Return CShort(obj)

            Catch ex As Exception
                Return 0

            End Try



        End Function

        ''' <summary>
        ''' Converts An Object to Short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Int32) As Int16
            Try

                Return CShort(obj)

            Catch ex As Exception
                Return 0

            End Try



        End Function


        ''' <summary>
        ''' Converts An Object to Short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Long) As Int16
            Try

                Return CShort(obj)

            Catch ex As Exception
                Return 0

            End Try



        End Function



        ''' <summary>
        ''' Converts boolean to short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Boolean) As Int16
            Try

                Return CShort(Math.Abs(CInt(obj)))

            Catch ex As Exception
                Return 0

            End Try

        End Function

        ''' <summary>
        ''' Converts objects to short
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Object) As Short

            If obj Is Nothing Then Return 0

            If TypeOf obj Is String Then Return valueOf(CType(obj, String))
            If TypeOf obj Is DBNull Then Return valueOf(CType(obj, DBNull))
            If TypeOf obj Is Double Then Return valueOf(CType(obj, Double))
            If TypeOf obj Is Long Then Return valueOf(CType(obj, Long))
            If TypeOf obj Is Int32 Then Return valueOf(CType(obj, Int32))
            If TypeOf obj Is Int16 Then Return CType(obj, Int16)
            If TypeOf obj Is Boolean Then Return valueOf(CType(obj, Boolean))

            Return 0    REM Cant convert this

        End Function


    End Class

End Namespace