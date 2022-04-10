


Namespace InputsParsing

    Public NotInheritable Class FrenchTextParsing



        ''' <summary>
        ''' . in Ascii is 46
        ''' </summary>
        ''' <remarks></remarks>
        Public Const DOT As Byte = 46

        ''' <summary>
        ''' For phone numbers
        ''' </summary>
        ''' <remarks></remarks>
        Public Const PLUS_SIGN As Byte = 43


        Public Const COMMA As Byte = 44


        ''' <summary>
        ''' Parse out double from string. reading to the last char
        ''' </summary>
        ''' <param name="pVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function parseOutDouble(pVal As String) As Double
            Return Objects.EDouble.ValueOf(parseOutDoubleAsString(pVal))
        End Function


        ''' <summary>
        ''' Parse out double from string. reading to the last char
        ''' </summary>
        ''' <param name="pVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function parseOutDoubleAsString(pVal As String) As String
            If pVal Is Nothing OrElse pVal = String.Empty Then Return pVal
            Dim pValid As IEnumerable(Of Char) = From dChar As Char In pVal.ToCharArray()
                                                 Where TextParsing.IsNumber(dChar) OrElse
                                                 AscW(dChar) = CInt(COMMA)
                                                 Select dChar

            '   Consider Negative Sign
            Dim pSymbol As String = String.Empty
            Dim pIndexOfNegative = pVal.IndexOf("-")
            If pIndexOfNegative >= 0 AndAlso pValid.Count() > 0 Then
                If pVal.IndexOf(pValid.First()) > pIndexOfNegative Then
                    pSymbol = "-"
                End If
            End If
            '   -----------------------------------------


            Dim parsedString() As String = New String(pValid.ToArray()).Split(ChrW(COMMA))
            If parsedString.Length > 1 Then Return String.Format("{3}{0}{1}{2}", parsedString(0), ChrW(COMMA), parsedString(1), pSymbol)
            If parsedString.Length = 1 Then Return String.Format("{1}{0}", parsedString(0), pSymbol)
            Return String.Empty
        End Function




    End Class

End Namespace

