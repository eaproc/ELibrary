﻿
Namespace Objects



    Public Class EBoolean

        ''' <summary>
        ''' Converts An Object to Boolean
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValueOf(ByVal obj As Object) As Boolean
            Try
                If obj Is Nothing Then Return False

                If TypeOf obj Is Boolean Then
                    Return CType(obj, Boolean)

                ElseIf TypeOf obj Is DBNull Then
                    Return False

                ElseIf TypeOf obj Is String Then

                    Return ValueOf(CType(obj, String))

                ElseIf TypeOf obj Is Int16 OrElse TypeOf obj Is Int32 OrElse TypeOf obj Is Int64 OrElse TypeOf obj Is Double OrElse TypeOf obj Is Single OrElse TypeOf obj Is Decimal Then

                    Return ValueOf(ELong.ValueOf(obj))


                End If

            Catch ex As Exception

            End Try

            Return False
        End Function


        ''' <summary>
        ''' Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FALSE" 
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValueOf(ByVal obj As String) As Boolean
            Try
                If obj IsNot Nothing AndAlso obj <> String.Empty Then
                    If EStrings.equalsIgnoreCase(obj, "yes") OrElse
                        EStrings.equalsIgnoreCase(obj, "y") Then Return True
                    If EStrings.equalsIgnoreCase(obj, "no") OrElse
                        EStrings.equalsIgnoreCase(obj, "n") Then Return False
                    If EStrings.equalsIgnoreCase(obj, "true") OrElse
                        EStrings.equalsIgnoreCase(obj, "t") Then Return True
                    If EStrings.equalsIgnoreCase(obj, "false") OrElse
                        EStrings.equalsIgnoreCase(obj, "f") Then Return False
                End If
                REM Check if it is number else try returning it like True written in string
                Dim TryInt As Int16
                If Int16.TryParse(obj, TryInt) Then Return ValueOf(TryInt)

                Return CBool(obj)

            Catch ex As Exception
                Return False
            End Try

        End Function


        ''' <summary>
        ''' Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FALSE" 
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValueOf(ByVal obj As Int16) As Boolean
            Try
                REM Check if it is number else try returning it like True written in string

                Return CBool(Math.Abs(obj))

            Catch ex As Exception
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FASLSE" 
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValueOf(ByVal obj As Long) As Boolean
            Try
                REM Check if it is number else try returning it like True written in string

                Return CBool(Math.Abs(obj))

            Catch ex As Exception
                Return False
            End Try

        End Function



        ''' <summary>
        ''' Converts A  Boolean to Boolean
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ValueOf(ByVal obj As Boolean) As Boolean
            Try
                Return obj
            Catch ex As Exception
                Return False
            End Try

        End Function


    End Class

End Namespace