Imports System.Drawing
Imports ELibrary.Standard.VB.Modules

Namespace Objects

    Public Class ELong

        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As String) As Long
            Try

                If basExtensions.IsNothing(obj) Then Return 0
                If obj = "" Then Return 0
                If obj.Trim = "" Then Return 0

                Return CLng(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function


        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As System.DBNull) As Long
            Try

                Return 0

            Catch ex As Exception
                Return 0

            End Try

        End Function

        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Double) As Long
            Try

                Return CLng(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function


        ''' <summary>
        ''' Converts An Object to long
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function valueOf(ByVal obj As Decimal) As Long
            Try

                Return CLng(obj)

            Catch ex As Exception
                Return 0

            End Try

        End Function

        Public Overloads Shared Function valueOf(ByVal obj As Boolean) As Long
            Try

                Return Math.Abs(CLng(obj))

            Catch ex As Exception
                Return 0

            End Try

        End Function


        Public Overloads Shared Function valueOf(ByVal obj As Color) As Long
            Try

                Return obj.ToArgb

            Catch ex As Exception
                Return 0

            End Try

        End Function

        Public Overloads Shared Function valueOf(ByVal obj As Int16) As Long
            Return obj
        End Function

        Public Overloads Shared Function valueOf(ByVal obj As Int32) As Long
            Return obj
        End Function

        Public Overloads Shared Function valueOf(ByVal obj As Int64) As Long
            Return obj
        End Function

        Public Overloads Shared Function valueOf(ByVal obj As Object) As Long
            If obj Is Nothing Then Return 0

            If TypeOf obj Is String Then Return valueOf(CType(obj, String))
            If TypeOf obj Is DBNull Then Return 0
            If TypeOf obj Is Double Then Return valueOf(CType(obj, Double))
            If TypeOf obj Is Long Then Return CLng(obj)
            If TypeOf obj Is Int32 Then Return valueOf(CType(obj, Int32))
            If TypeOf obj Is Decimal Then Return valueOf(CType(obj, Decimal))
            If TypeOf obj Is Int16 Then Return valueOf(CType(obj, Int16))
            If TypeOf obj Is Boolean Then Return valueOf(CType(obj, Boolean))
            If TypeOf obj Is Color Then Return valueOf(CType(obj, Color))

            Return 0    REM Cant convert this
        End Function


    End Class

End Namespace