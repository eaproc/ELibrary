Option Explicit On
Option Strict On

Imports EPRO.Library.v3._5.Objects

REM Using Operators
REM On Objects Assignment = Automatically takes the value if same object Type[ int=int]
REM Others must be allowed using Ctype

REM Operator = is the Comparison Type

Namespace Types



    ''' <summary>
    ''' A type that can hold both date and Nothing[Null] produces Date or NULL String
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NullableDateTime

#Region "Constructors"


        Sub New(ByVal DateTimeVal As Date)
            Me.setValue(DateTimeVal)

        End Sub

        Sub New(ByVal DateTimeVal As String)

            Me.setValue(DateTimeVal)


        End Sub


        Sub New(ByVal DateTimeVal As DBNull)

            Me.__isNull = True
        End Sub



        Sub New(ByVal DateTimeVal As Object)

            Try

                If IsNothing(DateTimeVal) Then Exit Try

                If TypeOf DateTimeVal Is DBNull Then Exit Try

                If TypeOf DateTimeVal Is String Then Me.setValue(CStr(DateTimeVal)) : Exit Sub

                If TypeOf DateTimeVal Is Date Then Me.setValue(CDate(DateTimeVal)) : Exit Sub

                If TypeOf DateTimeVal Is NullableDateTime Then Me.setValue(CType(DateTimeVal, NullableDateTime)) : Exit Sub

            Catch ex As Exception

            End Try

            Me.__isNull = True
        End Sub



#End Region


#Region "Properties"


        Public Shared ReadOnly NULL_TIME As New NullableDateTime(vbNullString)

        Public Shared ReadOnly NOW_DATETIME As New NullableDateTime(Date.Now)

        ''' <summary>
        ''' the value it returns instead of nothing
        ''' </summary>
        ''' <remarks></remarks>
        Public Const NULL_RETURN As String = "NULL"


        Private ____DateTimeVal As Date

        Public ReadOnly Property DateValue As String
            Get
                If Not Me.isNull Then
                    Return EDateTime.valueOf(Me.____DateTimeVal, EDateTime.DateFormats.DateFormatsEnum.DateFormat1)
                End If
                Return NULL_RETURN
            End Get
        End Property

        Public ReadOnly Property DateTimeValue As Date
            Get
                If Not Me.isNull Then
                    Return Me.____DateTimeVal
                End If
                Return Nothing
            End Get
        End Property


        Public ReadOnly Property TimeValue As String
            Get
                If Not Me.isNull Then
                    Return EDateTime.time__valueOf(Me.____DateTimeVal)
                End If
                Return NULL_RETURN
            End Get
        End Property


        Private __isNull As Boolean = True
        Public ReadOnly Property isNull As Boolean
            Get
                Return Me.__isNull
            End Get
        End Property

        ''' <summary>
        ''' returns DateTimeValue if it is not null else returns NULL string 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DateTimeValueOrNULL As Object
            Get
                If Not Me.isNull Then Return Me.DateTimeValue
                Return NULL_RETURN
            End Get
        End Property

        ''' <summary>
        ''' returns DateTimeValue if it is not null else returns NULL object (NOTHING) 
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property DateTimeValueOrNothing As Object
            Get
                If Not Me.isNull Then Return Me.DateTimeValue
                Return Nothing
            End Get
        End Property


#End Region


#Region "Operators"

        Public Shared Operator =(ByVal ___NullableDateTime As NullableDateTime,
                                  ByVal compareValue As Object) As Boolean
            Dim cmpVal As New NullableDateTime(compareValue)
            If ___NullableDateTime.isNull And cmpVal.isNull Then Return True

            If ___NullableDateTime.isNull Or cmpVal.isNull Then Return False

            Return ___NullableDateTime.DateTimeValue.Equals(cmpVal.DateTimeValue)

        End Operator

        Public Shared Operator <>(ByVal ___NullableDateTime As NullableDateTime,
                                 ByVal compareValue As Object) As Boolean

            Return Not (___NullableDateTime = compareValue)


        End Operator




#End Region



#Region "Methods"


        Public Sub setValue(ByVal DateTimeVal As String)


            Try

                If IsNothing(DateTimeVal) Then Exit Try
                If DateTimeVal = vbNullString Then Exit Try


                Me.____DateTimeVal = CDate(DateTimeVal)
                Me.__isNull = False
                Exit Sub

            Catch ex As Exception

            End Try
            Me.__isNull = True

        End Sub

        Public Sub setValue(ByVal DateTimeVal As Date)
            Try

                If DateTimeVal.Equals(Nothing) Then Exit Try


                Me.____DateTimeVal = DateTimeVal
                Me.__isNull = False
                Exit Sub

            Catch ex As Exception

            End Try
            Me.__isNull = True

        End Sub

        Private Sub setValue(ByVal __nullableDateTime As NullableDateTime)
            If __nullableDateTime Is Nothing Then Me.__isNull = True : Exit Sub
            Me.____DateTimeVal = __nullableDateTime.DateTimeValue
            Me.__isNull = __nullableDateTime.isNull

        End Sub

        Public Overrides Function Equals(ByVal obj As Object) As Boolean
            Return (Me = obj)
        End Function

        Public Overrides Function ToString() As String
            Return Me.DateTimeValue.ToString()
        End Function

#End Region



    End Class


End Namespace