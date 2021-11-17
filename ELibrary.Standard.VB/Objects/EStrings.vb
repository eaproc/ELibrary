Imports ELibrary.Standard.VB.Modules

Namespace Objects

    ''' <summary>
    ''' Controls the String Class
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EStrings

        ''' <summary>
        ''' For quotation mark
        ''' </summary>
        ''' <remarks></remarks>
        Public Const QUOTE As String = ChrW(34)



        ''' <summary>
        ''' Converts DBNUll to System.String.Empty()
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function valueOf(ByVal obj As DBNull) As String

            Return String.Empty

        End Function



        ''' <summary>
        ''' Converts Objects to String. Returns System.String.Empty() if it is nothing
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function valueOf(ByVal obj As Object) As String

            If obj Is Nothing Then Return String.Empty
            If TypeOf obj Is DBNull Then Return valueOf(CType(obj, DBNull))

            Return obj.ToString()

        End Function







        ''' <summary>
        ''' Checks if String1 equals String2 Ignoring the Case Sensitivity
        ''' </summary>
        ''' <param name="str1"></param>
        ''' <param name="str2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function equalsIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Boolean
            If str1 Is Nothing AndAlso str2 Is Nothing Then Return True
            If str1 Is Nothing OrElse str2 Is Nothing Then Return False


            Return str1.ToLower().Equals(str2.ToLower())
        End Function


        Public Shared Shadows Function Equals(obj1 As Object, obj2 As Object) As Boolean
            Return EStrings.valueOf(obj1) = EStrings.valueOf(obj2)
        End Function

        Public Shared Shadows Function Equals(obj1 As Object, obj2 As Object, IgnoreCase As Boolean) As Boolean
            If Not IgnoreCase Then Return Equals(obj1, obj2)
            Return equalsIgnoreCase(EStrings.valueOf(obj1), EStrings.valueOf(obj2))
        End Function






        ''' <summary>
        ''' Checks if String2 is in String1 Ignoring the Case Sensitivity
        ''' </summary>
        ''' <param name="str1"></param>
        ''' <param name="str2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function LikeIgnoreCase(ByVal str1 As String, ByVal str2 As String) As Boolean
            Return str1.ToLower().IndexOf(str2.ToLower()) >= 0
        End Function

        ''' <summary>
        ''' Checks if String2 is in String1 Case Sensitivity
        ''' </summary>
        ''' <param name="str1"></param>
        ''' <param name="str2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function [Like](ByVal str1 As String, ByVal str2 As String) As Boolean
            Return str1.IndexOf(str2) >= 0
        End Function


        ''' <summary>
        ''' Checks absolutely if a string is empty
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function isEmpty(ByVal value As String) As Boolean
            If value Is Nothing Then Return True
            If value = vbNullString Then Return True
            If String.IsNullOrEmpty(value) Then Return True
            If value.Trim().Equals(String.Empty) Then Return True
            Return False
        End Function



        '''' <summary>
        '''' Replaces the occurrence of space in the string with NULL
        '''' </summary>
        '''' <param name="strExpression"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function RemoveAllSpaces(ByVal strExpression As String) As String
        '    Return Replace(strExpression, " ", "")
        'End Function


        ''' <summary>
        ''' Quote String. Sample Man ==> "Man"
        ''' </summary>
        ''' <param name="Expression"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function QuoteString(ByVal Expression As String) As String
            Return String.Format(
                                "{0}{1}{0}", ChrW(34), Expression
                                )
        End Function


        '''
        ''' <summary>
        '''  If Original Length is Greater Than The Specified Length,
        '''     It returns the same string
        '''  Otherwise, if less than, String is buffered with specified character. Same Function as String.PadLeft or String.PadRight
        ''' 
        ''' <example>Wraps up a string to a Specified Length.  
        ''' <code>
        '''    [WrapUp("Man",4,"A")  = "ManA"]   
        '''    [WrapUp("Man",3,"A")  = "Man".]   
        '''    [WrapUp("Man",2,"A")  = "Man". ]
        ''' </code>
        ''' </example>
        '''</summary>
        Public Shared Function WrapUpAtLeast(ByVal OriginalString As String,
                                 ByVal ReturnLength As Integer,
                                 Optional ByVal BufferedCharacter As Char = "0"c,
                                 Optional ByVal BackPadding As Boolean = False) As String

            Try
                If Len(OriginalString) < ReturnLength Then
                    If BackPadding Then
                        Return OriginalString &
                            StrDup(ReturnLength - Len(OriginalString),
                                                       BufferedCharacter)
                    Else
                        Return StrDup(ReturnLength - Len(OriginalString),
                                                   BufferedCharacter) & OriginalString

                    End If
                End If

                ''If Len(OriginalString) > ReturnLength Then
                ''    Return Left(OriginalString, Len(OriginalString) - ReturnLength)
                ''End If


            Catch ex As Exception

            End Try

            Return OriginalString
        End Function

        '''
        ''' <summary>
        '''  If Original Length is Greater Than The Specified Length,
        '''      Length is cut to specified lenth
        '''  Otherwise, if less than, String is buffered with specified character. Same Function as String.PadLeft or String.PadRight
        ''' 
        ''' <example>Wraps up a string to a Specified Length.  
        ''' <code>
        '''    [WrapUp("Man",4,"A")  = "ManA"]   
        '''    [WrapUp("Man",3,"A")  = "Man".]   
        '''    [WrapUp("Man",2,"A")  = "Ma". ]
        ''' </code>
        ''' </example>
        '''</summary>
        Public Shared Function WrapUp(ByVal OriginalString As String,
                                 ByVal ReturnLength As Integer,
                                 Optional ByVal BufferedCharacter As Char = "0"c,
                                 Optional ByVal BackPadding As Boolean = False) As String

            Try
                If Len(OriginalString) < ReturnLength Then
                    If BackPadding Then
                        Return OriginalString &
                            StrDup(ReturnLength - Len(OriginalString),
                                                       BufferedCharacter)
                    Else
                        Return StrDup(ReturnLength - Len(OriginalString),
                                                   BufferedCharacter) & OriginalString

                    End If
                End If

                If Len(OriginalString) > ReturnLength Then
                    Return OriginalString.Substring(0, Len(OriginalString) - ReturnLength)
                    'Return Left(OriginalString, Len(OriginalString) - ReturnLength)
                End If


            Catch ex As Exception

            End Try

            Return OriginalString
        End Function


        ''' <summary>
        ''' Determines if Version1 precedes Version2. 
        ''' If 1 > 2 returns 1.
        ''' if 1 = 2 returns 0.
        ''' if 1 &lt; 2 returns -1.
        ''' </summary>
        ''' <param name="version1"></param>
        ''' <param name="version2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CompareVersions(ByVal version1 As String, ByVal version2 As String) As Integer

            Dim SplittedV1() As String = version1.Split("."), SplittedV2() As String = version2.Split("."), SplittedTMP() As String
            Dim isSwapped As Integer = 1


            REM Check for the longest length
            If SplittedV2.Count > SplittedV1.Count Then
                REM Swapp
                'SplittedTMP = New String() {}
                SplittedTMP = SplittedV1
                SplittedV1 = SplittedV2
                SplittedV2 = SplittedTMP
                SplittedTMP = Nothing  REM Free the space
                isSwapped = -1


            End If


            For xIndex As Integer = 0 To SplittedV1.Count - 1

                If xIndex < SplittedV2.Count Then

                    If CInt(SplittedV1(xIndex)) > CInt(SplittedV2(xIndex)) Then

                        Return 1 * isSwapped

                    ElseIf CInt(SplittedV1(xIndex)) < CInt(SplittedV2(xIndex)) Then

                        Return -1 * isSwapped

                    End If


                ElseIf CInt(SplittedV1(xIndex)) > 0 Then

                    Return 1 * isSwapped

                End If

            Next


            'CompareVersions = version1.CompareTo(version2)

            Return 0

        End Function





        ''' <summary>
        ''' Splits a string and returns result with empty string element
        ''' </summary>
        ''' <param name="expression"></param>
        ''' <param name="delimiter"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function splitWithoutNullElement(ByVal expression As String,
                                                       Optional ByVal delimiter As String = ",") As String()

            Dim s_split() As String = expression.Split(delimiter)
            If expression = vbNullString Then Return New String() {}
            If s_split Is Nothing Then Return New String() {}
            If s_split.Length = 0 Then Return New String() {}

            Dim l_split As List(Of String) = New List(Of String)(s_split.Length)

            For Each s As String In s_split

                If s.Trim = vbNullString Then Continue For
                l_split.Add(s)


            Next



            Return l_split.ToArray

        End Function


        ''' <summary>
        ''' Split strings into chunkSize
        ''' </summary>
        ''' <param name="str"></param>
        ''' <param name="chunkSize"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function SplitChunk(str As String, chunkSize As Integer) As IEnumerable(Of String)
            Return Enumerable.Range(0, CInt(str.Length / chunkSize)).[Select](
                Function(i) str.Substring(i * chunkSize, chunkSize))
        End Function









        '        ''' <summary>
        '        ''' Extract String from an expression using the first appearance of the limits. 
        '        ''' Example: StartString = lstmessagelst ... Stop String = lst/messagelst
        '        ''' </summary>
        '        ''' <param name="HTMLString"></param>
        '        ''' <param name="startString"></param>
        '        ''' <param name="stopString"></param>
        '        ''' <param name="locatePosition"></param>
        '        ''' <returns></returns>
        '        ''' <remarks></remarks>
        '        Public Shared Function ExtractStringFromHtml(ByVal HTMLString As String, ByVal startString As String,
        '                                                      ByVal stopString As String,
        '                                                      Optional ByRef locatePosition As Double = 1) As String

        '            On Error GoTo errHandler

        '            Dim fLocate As Long, rLocate As Long

        '            fLocate = InStr(CInt(locatePosition), HTMLString, startString)
        '            If fLocate <> 0 Then
        '                rLocate = InStr(CInt(fLocate), HTMLString, stopString)

        '                'Add String Position
        '                fLocate = fLocate + Len(startString)

        '                'ExtractStringFromHtml = Mid(HTMLString, fLocate, rLocate - fLocate)

        '                'Return Stop Position
        '                locatePosition = rLocate
        '                Return Mid(HTMLString, CInt(fLocate), CInt(rLocate - fLocate))

        '            End If
        '            Return vbNullString
        '            Exit Function
        'errHandler:
        '        End Function


        ''' <summary>
        ''' Fetch first letter of the string and abbreviate it
        ''' </summary>
        ''' <param name="strString"></param>
        ''' <param name="Capitalized"></param>
        ''' <param name="AbbvSymbol"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getFirstAbbreviatedLetter(ByVal strString As String,
                                                         Optional ByVal Capitalized As Boolean = True,
                                                         Optional ByVal AbbvSymbol As String = "."
                                                                                               ) As String
            If strString.Length < 1 Then Return ""
            Dim AbbvStr As String = strString.Substring(0, 1).ToLower

            If Capitalized Then AbbvStr = AbbvStr.ToUpper

            Return AbbvStr & AbbvSymbol

        End Function



        ''' <summary>
        ''' Uses this syntax *:=0 instead of the syntax of string.format {0}. [?Microsoft.VisualBasic.Strings.Format()] -Dont Confuse it
        ''' </summary>
        ''' <param name="Expression">The expression that contains</param>
        ''' <param name="args"></param>
        ''' <returns></returns>
        ''' <remarks>Becareful so you dont confuse it with [?Microsoft.VisualBasic.Strings.Format()]</remarks>
        Public Shared Function FormatString(ByVal Expression As String,
                                        ByVal ParamArray args() As String
                                        ) As String

            Dim loopInc As Byte
            For Each arg As String In args
                Expression = Replace(Expression,
                        String.Format("*:={0}", loopInc),
                        arg
                        )

                loopInc = CByte(loopInc + 1)
            Next

            Return Expression

        End Function



        ''' <summary>
        ''' Removes all non-readable characters from strings range [0 - 31]
        ''' </summary>
        ''' <param name="ORIGINAL_STRING"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function CLEAN_UP_STRING(ByVal ORIGINAL_STRING As String) As String

            Dim byteCol As Byte
            'Remove all unwanted keys 0-31
            For byteCol = 0 To 31
                ORIGINAL_STRING = Replace(ORIGINAL_STRING, ChrW(byteCol), "")

                'Prevent Freeze
                'DoEvents()
            Next

            'incase null was added, remove it again
            CLEAN_UP_STRING = Replace(ORIGINAL_STRING, ChrW(0), "")

        End Function



        ''' <summary>
        ''' Reverse the sequence of a string from back to front
        ''' </summary>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Reverse(ByVal value As String) As String
            Dim arr() As Char = value.ToCharArray()
            Array.Reverse(arr)

            Return New String(arr)
        End Function


        ''' <summary>
        ''' Dont use unless necessary. 
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns>Returns * 3 the length of string</returns>
        ''' <remarks></remarks>
        Public Shared Function EncodeToFixLengthBytes(ByVal Str As String) As String
            Dim str_bytes As String = vbNullString
            For Each c As Char In Str.ToCharArray

                str_bytes &= WrapUp(CStr(AscW(c)), 3)

            Next

            Return str_bytes

        End Function

        ''' <summary>
        ''' Reverse the process of EncodeTOFixLengthBytes
        ''' </summary>
        ''' <param name="Str"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Reverse_EncodeToFixLengthBytes(ByVal Str As String) As String

            If Len(Str) Mod 3 <> 0 Then Return vbNullString

            Dim rtr_str As String = vbNullString
            Dim i As Int16 = 0

            Try
                For i = 0 To CShort(Len(Str) - 1) Step 3

                    rtr_str &= ChrW(EInt.valueOf(Str.Substring(i, 3)))

                Next

            Catch ex As Exception

            End Try
            Return rtr_str
        End Function





        ''' <summary>
        ''' Checks if whole String is combination of escape characters like \r\n
        ''' </summary>
        ''' <param name="pVal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function isEscapeCharacters(pVal As String) As Boolean

            If pVal Is Nothing OrElse pVal = String.Empty Then Return False

            Dim IRst As IEnumerable(Of String) = From ds As String In pVal.Split("\"c)
                                                 Where ds = "n" OrElse ds = "r" OrElse ds = "t" OrElse ds = "b" OrElse ds = "f" OrElse ds = "'" OrElse ds = "\" OrElse ds = """"
                                                 Select ds

            Return IRst.Count() * 2 = pVal.Length
        End Function


        ''' <summary>
        ''' Converts Combination of escape characters to String else returns empty strings
        ''' </summary>
        ''' <param name="pEscapeChars"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function translateEscapeCharacters(pEscapeChars As String) As String
            If Not isEscapeCharacters(pEscapeChars) Then Return String.Empty
            Dim rs As String = String.Empty

            For Each ds As String In pEscapeChars.Split("\"c)
                Select Case ds
                    Case "n"
                        rs &= ChrW(10)
                    Case "r"
                        rs &= ChrW(13)
                    Case "t"
                        rs &= ChrW(9)
                    Case "b"
                        rs &= ChrW(8)
                    Case "f"
                        rs &= ChrW(12)
                    Case """"
                        rs &= ChrW(34)
                    Case "'"
                        rs &= ChrW(39)
                    Case "\"
                        rs &= ChrW(92)
                End Select
            Next
            Return rs
        End Function








        '
        '   Base64
        '
        ''' <summary>
        ''' It handles exception
        ''' </summary>
        ''' <param name="pString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToBase64(ByVal pString As String) As String
            Return ConvertToBase64(pString, System.Text.Encoding.UTF8)
        End Function
        ''' <summary>
        ''' It handles exception
        ''' </summary>
        ''' <param name="pString"></param>
        ''' <param name="pEncoding"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertToBase64(ByVal pString As String,
                                               pEncoding As System.Text.Encoding
                                               ) As String


            Return Convert.ToBase64String(
                    pEncoding.GetBytes(pString)
                    )


        End Function

        ''' <summary>
        ''' It handles exception
        ''' </summary>
        ''' <param name="pString"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertFromBase64(ByVal pString As String) As String
            Return ConvertFromBase64(pString,
                               System.Text.Encoding.UTF8
                               )
        End Function

        ''' <summary>
        ''' It handles exception
        ''' </summary>
        ''' <param name="pString"></param>
        ''' <param name="pEncoding"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertFromBase64(ByVal pString As String,
                                               pEncoding As System.Text.Encoding
                                               ) As String

            pString = pString.Trim()
                If pString = String.Empty Then Return pString

                Dim pMode = pString.Length Mod 4

                If pMode <> 0 Then pString = pString.PadRight(pString.Length + 4 - pMode, "="c)

                Return pEncoding.GetString(
                    Convert.FromBase64String(
                    pString
                   )
               )


        End Function






    End Class


End Namespace