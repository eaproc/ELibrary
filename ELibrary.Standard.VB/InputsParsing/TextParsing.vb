


Namespace InputsParsing

    Public NotInheritable Class TextParsing



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














        Public Shared Function IsSmallLetter(c As Char) As Boolean
            REM Char is UTF-16
            REM  ' ''97-122 small letters [a - z]
            Select Case c
                Case ChrW(97) To ChrW(122)
                    Return True
            End Select

            Return False
        End Function

        Public Shared Function IsBigLetter(c As Char) As Boolean

            REM   ' ''65-90  Big Letters [A - Z]
            Select Case c
                Case ChrW(65) To ChrW(90)
                    Return True
            End Select

            Return False
        End Function

        ''' <summary>
        ''' Currently this isn't perfect on Other Encoding Apart from ASCII because any char not ASCII letters, 
        ''' numbers, space, non readable are considered symbols. Therefore, other special chars too are symbols
        ''' </summary>
        ''' <param name="c"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsSymbol(c As Char) As Boolean

            Return Not IsBigLetter(c) AndAlso Not IsSmallLetter(c) AndAlso Not IsNonReadableCharacter(c) AndAlso
                Not IsNumber(c) AndAlso Not IsSpace(c)

        End Function


        ''' <summary>
        ''' Includes Tabs and Enter Keys. Use this if you want to allow Tabs and Enter Keys
        ''' </summary>
        ''' <param name="c"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsNonReadableCharacter(c As Char) As Boolean

            REM  ' ''0 - 31
            Select Case c
                Case ChrW(0) To ChrW(31)
                    Return True
            End Select

            Return False

        End Function

        ''' <summary>
        ''' without counting 9,10,11,13 [Tabs and Enter Keys]. You can use this if you don't want to allow those stuff
        ''' </summary>
        ''' <param name="c"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsNonReadableCharacterExceptTabsAndEnterKey(c As Char) As Boolean

            REM  ' ''0 - 31 without 9,10,11,13 
            Select Case c
                Case ChrW(0) To ChrW(8), ChrW(12), ChrW(14) To ChrW(31)
                    Return True
            End Select

            Return False

        End Function


        Public Shared Function IsNumber(c As Char) As Boolean

            ' ''48 - 57 = > Digits [0 - 9]
            Select Case c
                Case ChrW(48) To ChrW(57)
                    Return True
            End Select

            Return False
        End Function


        Public Shared Function IsValidEmail(pEmail As String) As Boolean

            Try

                Dim p = New System.Net.Mail.MailAddress(pEmail)

                If p.User = String.Empty Then Return False
                If p.User.Length < 1 OrElse p.Host.Length < 4 Then Return False

                If pEmail.IndexOf(".@") >= 0 Then Return False
                If pEmail.IndexOf("..") >= 0 Then Return False
                If pEmail.IndexOf(".@") >= 0 Then Return False
                If pEmail.IndexOf("@.") >= 0 Then Return False
                If pEmail.IndexOf(" ") >= 0 Then Return False
                If pEmail.IndexOf("*") >= 0 Then Return False
                If pEmail.IndexOf("[]") >= 0 Then Return False
                If pEmail.IndexOf("[[") >= 0 Then Return False
                If pEmail.IndexOf("]]") >= 0 Then Return False
                If pEmail.IndexOf(",") >= 0 Then Return False
                If pEmail.IndexOf(";") >= 0 Then Return False
                If pEmail.IndexOf(":") >= 0 Then Return False
                If pEmail.IndexOf("/") >= 0 Then Return False



                Return True
            Catch ex As Exception
                Return False
            End Try
        End Function


        Public Shared Function IsSpace(c As Char) As Boolean
            Return c = ChrW(32)
        End Function

        ''' <summary>
        ''' Returns if this is enter key or tab
        ''' </summary>
        ''' <param name="c"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsTabsAndEnterKeys(c As Char) As Boolean
            Select Case c
                Case ChrW(9) To ChrW(11), ChrW(13)
                    Return True

                Case Else
                    Return False
            End Select
        End Function






        ''' <summary>
        ''' Parse out all numbers from the left of a string
        ''' </summary>
        ''' <param name="pVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function parseOutIntegers(pVal As String) As String
            If pVal Is Nothing OrElse pVal = String.Empty Then Return String.Empty
            Dim pValid As IEnumerable(Of Char) = From dChar As Char In pVal.ToCharArray()
                                                 Where IsNumber(dChar)
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


            Return pSymbol & New String(pValid.ToArray())
        End Function


        ''' <summary>
        ''' Parse out double from string. reading to the last char
        ''' </summary>
        ''' <param name="pVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function parseOutDouble(pVal As String) As Double
            Return Objects.EDouble.valueOf(parseOutDoubleAsString(pVal))
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
                                                 Where IsNumber(dChar) OrElse
                                                 AscW(dChar) = CInt(DOT)
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




            Dim parsedString() As String = New String(pValid.ToArray()).Split("."c)
            If parsedString.Length > 1 Then Return String.Format("{2}{0}.{1}", parsedString(0), parsedString(1), pSymbol)
            If parsedString.Length = 1 Then Return String.Format("{1}{0}", parsedString(0), pSymbol)
            Return String.Empty
        End Function




    End Class

End Namespace

