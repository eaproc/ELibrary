
Imports System.Collections
Imports System.IO

Imports ELibrary.Standard.VB.Modules

Namespace AppConfigurations

    Public Class IniReader

#Region "Constructors"

        ''' <summary>
        ''' Reads Ini File.
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        ''' <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        ''' <param name="pisCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String,
                         encode As System.Text.Encoding,
                         ByVal KeyVDelimiter As String,
                        ByVal LineVDelimiter As String,
                       ByVal pisCaseSensitive As Boolean,
                        ByVal pIdentifySections As Boolean)
            Me.New(iniFilePath, encode,
                   KeyVDelimiter, LineVDelimiter, pisCaseSensitive, pIdentifySections, New String() {})
        End Sub

        ''' <summary>
        ''' Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        ''' <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String, ByVal KeyVDelimiter As String,
                        ByVal LineVDelimiter As String,
                        ByVal isCaseSensitive As Boolean,
                       ByVal pIdentifySections As Boolean
                        )
            Me.New(iniFilePath, System.Text.Encoding.Default,
                   KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, New String() {})
        End Sub

        ''' <summary>
        ''' Reads Ini File.
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String,
                        encode As System.Text.Encoding,
                         ByVal isCaseSensitive As Boolean,
                       ByVal pIdentifySections As Boolean
                       )
            Me.New(iniFilePath, encode,
                   DEFAULT_KeyValuePairDelimiter, DEFAULT_LineDelimiter, isCaseSensitive, pIdentifySections, New String() {})
        End Sub

        ''' <summary>
        ''' Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String,
                       ByVal isCaseSensitive As Boolean,
                       ByVal pIdentifySections As Boolean
                       )
            Me.New(iniFilePath, System.Text.Encoding.Default,
                   DEFAULT_KeyValuePairDelimiter, DEFAULT_LineDelimiter, isCaseSensitive, pIdentifySections, New String() {})
        End Sub

        ''' <summary>
        ''' Reads Ini File. Uses ANSII - WINDOWS  Encoding.  System.Text.Encoding.Default
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        ''' <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String,
                       ByVal KeyVDelimiter As String,
                       ByVal LineVDelimiter As String,
                       ByVal isCaseSensitive As Boolean,
                       ByVal pIdentifySections As Boolean,
                         ParamArray IgnoreSpecialComments As String()
                       )
            Me.New(iniFilePath, System.Text.Encoding.Default,
                    KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, IgnoreSpecialComments
                    )
        End Sub
        ''' <summary>
        ''' Reads Ini File
        ''' </summary>
        ''' <param name="iniFilePath">The Ini File to read</param>
        ''' <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        ''' <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFilePath As String,
                        encode As System.Text.Encoding,
                       ByVal KeyVDelimiter As String,
                       ByVal LineVDelimiter As String,
                       ByVal isCaseSensitive As Boolean,
                       ByVal pIdentifySections As Boolean,
                         ParamArray IgnoreSpecialComments As String()
                       )


            Dim f As String = File.ReadAllText(iniFilePath, encode)
                If (f.Trim() <> String.Empty) Then

                    Me.LoadData(f, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, New List(Of String)(IgnoreSpecialComments))

                End If

        End Sub


        ''' <summary>
        ''' Reads Ini File
        ''' </summary>
        ''' <param name="iniFileContents">Parse the contents of ini file to read</param>
        ''' <param name="KeyVDelimiter">The key value delimiter like Equals sign [ = ]</param>
        ''' <param name="LineVDelimiter">The Line delimiter or Entry delimiter like New Line [ \n ]</param>
        ''' <param name="isCaseSensitive">Indicate if fetching result will be case sensitive</param>
        ''' <param name="IgnoreSpecialComments">Entries of Special Comments Starter lines like REM. Semi-colon ; </param>
        ''' <remarks></remarks>
        Public Sub New(ByVal iniFileContents As String, ByVal KeyVDelimiter As String,
                       ByVal LineVDelimiter As String, ByVal IgnoreSpecialComments As List(Of String),
                       ByVal isCaseSensitive As Boolean,
                        ByVal pIdentifySections As Boolean
                       )


            Me.LoadData(iniFileContents, KeyVDelimiter, LineVDelimiter, isCaseSensitive, pIdentifySections, IgnoreSpecialComments)

        End Sub


#End Region



#Region "Properties"

        Private IniDetails As Dictionary(Of String, String)

        Private ___isCaseSensitive As Boolean
        Public Function isCaseSensitive() As Boolean
            Return Me.___isCaseSensitive
        End Function

        Private _readSuccessFully As Boolean = False
        Public Const DEFAULT_KeyValuePairDelimiter As String = "="
        Public Const DEFAULT_LineDelimiter As String = "\r\n"
        Private KeyValuePairDelimiter As String = DEFAULT_KeyValuePairDelimiter


        ''' <summary>
        ''' It is use to identify sections if Identify Sections was indicated
        ''' </summary>
        ''' <remarks></remarks>
        Public Const SECTION__NAME___SEPERATOR As String = "__________SEC_________"


        Private _______IsSectionIdentified As Boolean
        Public ReadOnly Property IsSectionIdentified As Boolean
            Get
                Return _______IsSectionIdentified
            End Get
        End Property



        Public Function isValid() As Boolean

            Return Me._readSuccessFully

        End Function


        Public Function KeysCounts() As Integer

            If (Not Me.isValid()) Then Return 0

            Return Me.Keys().Count()

        End Function


        Public Function Keys() As IEnumerable(Of String)

            If (Not Me.isValid()) Then Return New List(Of String)()
            Return Me.IniDetails.Keys.AsEnumerable()

        End Function

#End Region


#Region "Methods"


        Public Function getValue(ByVal Key As String) As String

            If (Not Me.___isCaseSensitive) Then Key = Key.ToLower()
            If (Me.Keys().Contains(Key)) Then Return Me.IniDetails(Key)
            Return String.Empty

        End Function



        Private Sub LoadData(ByVal iniFileContents As String, ByVal KeyVDelimiter As String,
                       ByVal LineVDelimiter As String, ByVal isCaseSensitive As Boolean, ByVal pIdentifySections As Boolean,
                          ByVal IgnoreSpecialComments As List(Of String)
                          )
            Me.KeyValuePairDelimiter = KeyVDelimiter

            REM Translate LineDelimiter
            If Objects.EStrings.IsEscapeCharacters(LineVDelimiter) Then LineVDelimiter = Objects.EStrings.TranslateEscapeCharacters(LineVDelimiter)


            If (IgnoreSpecialComments Is Nothing) Then IgnoreSpecialComments = New List(Of String)

            Dim IgnoreComments As List(Of String) = New List(Of String)(50)  REM Maximum 50 types is enough :P
            IgnoreComments.Add(";")
            If (IgnoreSpecialComments.Contains(IgnoreComments(0))) Then IgnoreComments.RemoveAt(0) REM avoid duplicates
            IgnoreComments.AddRange(IgnoreSpecialComments)


            Me.IniDetails = New Dictionary(Of String, String)(100)  REM 100 keys maximum lol :D
            Me.___isCaseSensitive = isCaseSensitive
            Me._______IsSectionIdentified = pIdentifySections



            Dim f As String = iniFileContents
                If (f.Trim() <> String.Empty) Then

                    REM Parse File
                    Dim Lines As String() = f.Split(New String() {LineVDelimiter}, StringSplitOptions.RemoveEmptyEntries)

                    If (Lines.Length > 0) Then

                        Dim queries As String = String.Empty
                        REM remove comments first
                        For Each Line As String In Lines


                            Dim IgnoreLine As Boolean = False
                            REM support comments and special comments
                            For Each comment As String In IgnoreComments

                                If (Line.StartsWith(comment, StringComparison.CurrentCultureIgnoreCase)) Then
                                    IgnoreLine = True : Exit For
                                End If

                            Next

                            If (Not IgnoreLine) Then queries &= Line & LineVDelimiter REM preserve the breaks

                        Next




                        '   If KeySections are meant to be identified
                        Dim vLastKeyIdentified As String = String.Empty


                        REM Date.Now process the files using the real delimiter
                        For Each Line As String In queries.Split(New String() {LineVDelimiter}, StringSplitOptions.RemoveEmptyEntries)

                            Dim pTestKey = Line.Trim(New Char() {CChar(" "), ChrW(9), ChrW(13)})
                            If pIdentifySections AndAlso pTestKey.StartsWith("[") AndAlso pTestKey.EndsWith("]") Then
                                ' This entry is a key
                                vLastKeyIdentified = pTestKey.Substring(1, pTestKey.Length - 2)
                                Continue For

                            End If

                            Dim vAppendSectionToKeyName As String = String.Empty

                            If vLastKeyIdentified <> String.Empty Then vAppendSectionToKeyName = vLastKeyIdentified & SECTION__NAME___SEPERATOR



                            REM get keys values
                            REM Dim keyValue As String() = Line.Split(New String() {KeyValuePairDelimiter}, StringSplitOptions.RemoveEmptyEntries)
                            Dim keyValue As String() = Line.Split(KeyValuePairDelimiter)
                            If (keyValue.Length = 2) Then

                                REM if error occured here, it means user entered same key more than once
                                If (Not Me.___isCaseSensitive) Then
                                    Me.IniDetails.Add(
                                       vAppendSectionToKeyName & keyValue(0).ToLower().Trim(New Char() {CChar(" "), ChrW(9), ChrW(13)}),
                                        keyValue(1).Trim()
                                        )
                                Else
                                    Me.IniDetails.Add(
                                         vAppendSectionToKeyName & keyValue(0).Trim(New Char() {CChar(" "), ChrW(9), ChrW(13)}),
                                         keyValue(1).Trim()
                                         )

                                End If

                            End If

                        Next




                        REM Reconfirm additions
                        If (Me.IniDetails.Count <> 0) Then Me._readSuccessFully = True

                    End If
                End If


        End Sub


#End Region



    End Class

End Namespace
