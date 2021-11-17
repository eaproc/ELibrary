Option Explicit On
Option Strict On

Imports EPRO.Library.v3._5.Objects

Namespace Modules



    Public Module basExtensions

#Region "MathsFunctions"

        ''' <summary>
        ''' Toggles Integer Values e.g betw
        ''' </summary>
        ''' <param name="current"></param>
        ''' <param name="min"></param>
        ''' <param name="max"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()> _
        Public Function ToggleValue(ByVal current As Integer, ByVal min As Integer, ByVal max As Integer) As Integer
            If min > max OrElse (current <> max And current <> min) Then Return min

            Return min + (max - current)

        End Function


#End Region


#Region "Object Conversions"
        REM Dont put object as extension

        <System.Runtime.CompilerServices.Extension()> _
        Public Function toInt32(ByVal val As Object) As Int32
            Return EInt.valueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()> _
        Public Function toLong(ByVal val As Object) As Long
            Return ELong.valueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()> _
        Public Function toDouble(ByVal val As Object) As Double
            Return EDouble.valueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()> _
        Public Function toDecimal(ByVal val As Object) As Decimal
            Return EDecimal.valueOf(val)
        End Function

        ''' <summary>
        ''' Converts an Any Object to Short Type
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()> _
        Public Function toShort(ByVal val As Object) As Short
            Return EShort.valueOf(obj:=val)
        End Function

        ''' <summary>
        ''' Converts an Any Object to Boolean
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()> _
        Public Function toBoolean(ByVal val As Object) As Boolean
            Return EBoolean.valueOf(obj:=val)
        End Function

#End Region


#Region "String Functions"

        ''' <summary>
        ''' Checks if String1 equals String2 Ignoring the Case Sensitivity
        ''' </summary>
        ''' <param name="str1"></param>
        ''' <param name="str2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()> _
        Public Function equalsIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Boolean
            Return EStrings.equalsIgnoreCase(str1, str2)
        End Function


        ''' <summary>
        ''' Converts to Proper Case
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()> _
        Public Function toProperCase(ByVal pStr As String) As String
            Return Strings.StrConv(pStr, VbStrConv.ProperCase)
        End Function


#End Region

    End Module

End Namespace