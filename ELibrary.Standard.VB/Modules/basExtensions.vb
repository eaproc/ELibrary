Option Explicit On
Option Strict On

Imports ELibrary.Standard.VB.Objects

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
        <System.Runtime.CompilerServices.Extension()>
        Public Function ToggleValue(ByVal current As Integer, ByVal min As Integer, ByVal max As Integer) As Integer
            If min > max OrElse (current <> max And current <> min) Then Return min

            Return min + (max - current)

        End Function

        ''' <summary>
        ''' Formats number into string like N2
        ''' </summary>
        ''' <param name="current"></param>
        ''' <param name="min"></param>
        ''' <param name="max"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function FormatNumber(ByVal v As Double, ByVal precision As Integer) As String
            Return v.ToString(String.Format("N{0}", precision))
        End Function


#End Region


#Region "Object Conversions"
        REM Dont put object as extension

        <System.Runtime.CompilerServices.Extension()>
        Public Function ToInt32(ByVal val As Object) As Int32
            Return EInt.ValueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()>
        Public Function ToLong(ByVal val As Object) As Long
            Return ELong.ValueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()>
        Public Function ToDouble(ByVal val As Object) As Double
            Return EDouble.ValueOf(val)
        End Function


        <System.Runtime.CompilerServices.Extension()>
        Public Function ToDecimal(ByVal val As Object) As Decimal
            Return EDecimal.ValueOf(val)
        End Function

        ''' <summary>
        ''' Converts an Any Object to Short Type
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function ToShort(ByVal val As Object) As Short
            Return EShort.ValueOf(obj:=val)
        End Function

        ''' <summary>
        ''' Converts an Any Object to Boolean
        ''' </summary>
        ''' <param name="val"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function ToBoolean(ByVal val As Object) As Boolean
            Return EBoolean.ValueOf(obj:=val)
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
            Return EStrings.EqualsIgnoreCase(str1, str2)
        End Function


        ''' <summary>
        ''' Converts to Proper Case
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function toProperCase(ByVal pStr As String) As String
            Return System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(pStr)
            'Return Strings.StrConv(pStr, VbStrConv.ProperCase)
        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function Len(ByVal pStr As String) As Integer
            Return pStr.Length
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function InStr(ByVal StartIndex As Integer, ByVal HayStack As String, ByVal Needle As String) As Integer
            Return HayStack.IndexOf(Needle, StartIndex)
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function Mid(ByVal HayStack As String, ByVal StartIndex As Integer, ByVal Length As Integer) As String
            Return HayStack.Substring(StartIndex, Length)
        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function Replace(ByVal ORIGINAL_STRING As String, ByVal Search As String, ByVal Replacement As String) As String
            ' Replace(ORIGINAL_STRING, ChrW(0), "")
            Return ORIGINAL_STRING.Replace(Search, Replacement)
        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function IsNothing(ByVal d As Object) As Boolean
            Return d Is Nothing
        End Function


        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function StrDup(ByVal d As Int32, ByVal c As Char) As String
            Return New String(Enumerable.Range(0, d).Select(Function(z) c).ToArray())
        End Function

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="pStr"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <System.Runtime.CompilerServices.Extension()>
        Public Function IsNullOrTrimmedEmpty(ByVal str As String) As Boolean
            Return str Is Nothing OrElse Not str.Trim().Any()
        End Function




#End Region

    End Module

End Namespace