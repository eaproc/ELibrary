Imports System.Windows.Forms
Imports ELibrary.Standard.VB.InputsParsing.TextParsing


Namespace InputsParsing

    Public NotInheritable Class TextHandlers





        Public Delegate Sub dlg_addOrRemove__UseProperCase__addressOfKeyPress(ByVal s As Object,
                                                         ByVal e As System.Windows.Forms.KeyPressEventArgs,
                                                         ByRef CaretPosition As Int32)

        ''' <summary>
        ''' Meant for single line text
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <param name="CaretPosition"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__UseProperCase__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs,
                                                                 ByRef CaretPosition As Int32)


            If s Is Nothing OrElse e Is Nothing Then Return


            Dim sender As Control = CType(s, Control)
            If CaretPosition > sender.Text.Length Then Return

            REM if this is the first char capitalize it
            REM if this is not and the character @ the end is space then capitalize it
            REM Else do nothing
            Dim selStart As Int32 = CaretPosition
            sender.Text = StrConv(sender.Text, VbStrConv.ProperCase)
            CaretPosition = selStart
            REM Ignore if it is del or backspace, it will be corrected on next key stroke
            REM On next stroke if prev text is not conformed to Uppercase then go through each char one by one and correct all
            If sender.Text.Length = 0 OrElse Right(sender.Text, 1) = " " Then e.KeyChar = UCase(e.KeyChar)


        End Sub


        ''' <summary>
        ''' Disable Control V
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__DisableControlV__addressOfKeyDown(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyEventArgs)
            If s Is Nothing OrElse e Is Nothing Then Return
            If e.Control AndAlso e.KeyCode = Keys.V Then e.SuppressKeyPress = True
        End Sub

        ''' <summary>
        ''' Disable Right Click Menu
        ''' </summary>
        ''' <param name="s"></param>
        ''' <remarks></remarks>
        Public Shared Sub DisableRightClick(ByVal s As Control)

            If s Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            If sender.ContextMenu Is Nothing Then sender.ContextMenu = New ContextMenu()
            REM Throw New Exception("Just set ShortCut Enabled or set .ContextMenu=new ContextMenu() to False on the Text Box")

        End Sub



        ''' <summary>
        ''' Makes the text control to accept numbers only. Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersOnly__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''' <summary>
        '''  Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithAlphabetsAndSymbols__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If IsSpace(e.KeyChar) OrElse IsTabsAndEnterKeys(e.KeyChar) Then e.KeyChar = NULL


        End Sub



        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithSinglePlusSignInFront__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            REM Allow plus in front +905
            If AscW(e.KeyChar) = PLUS_SIGN AndAlso sender.Text = String.Empty Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub



        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithDecimalDot__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            REM Now check if it is the only one
            If sender.Text.IndexOf(".") >= 0 And AscW(e.KeyChar) = DOT Then e.KeyChar = NULL : Return
            If AscW(e.KeyChar) = DOT Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''' <summary>
        ''' Doesnt allow tabs and enterkeys. For French Double Numbers
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithCOMMA__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            REM Now check if it is the only one
            If sender.Text.IndexOf(Chr(FrenchTextParsing.COMMA)) >= 0 And AscW(e.KeyChar) = FrenchTextParsing.COMMA Then e.KeyChar = NULL : Return
            If AscW(e.KeyChar) = FrenchTextParsing.COMMA Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub

        ''' <summary>
        ''' Doesnt allow tabs and enterkeys. For French Double Numbers
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithCOMMA___ConvertDotToComma__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            REM Now check if it is the only one
            If sender.Text.IndexOf(Chr(FrenchTextParsing.COMMA)) >= 0 AndAlso
                (
                    AscW(e.KeyChar) = FrenchTextParsing.COMMA OrElse
                              AscW(e.KeyChar) = FrenchTextParsing.DOT
                    ) Then e.KeyChar = NULL : Return

            If AscW(e.KeyChar) = FrenchTextParsing.COMMA Then Return
            If AscW(e.KeyChar) = FrenchTextParsing.DOT Then e.KeyChar = ChrW(FrenchTextParsing.COMMA) : Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithSpace__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub



        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithAlphabets__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithAlphabetsAndSpace__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL



        End Sub


        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__NumbersWithSymbols__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                Not IsSymbol(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub



        Public Shared Sub addOrRemove__NumbersWithSymbolsAndSpace__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return

            REM if not number and is readable character then dont allow
            If Not IsNumber(e.KeyChar) AndAlso
                Not IsSymbol(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''' <summary>
        ''' Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__AlphabetsOnly__addressOfKeyPress(ByVal s As Object,
                                                                 ByVal e As System.Windows.Forms.KeyPressEventArgs)


            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub

        ''' <summary>
        '''  Doesnt allow tabs and enterkeys
        ''' </summary>
        ''' <param name="s"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Shared Sub addOrRemove__AlphabetsWithSpace__addressOfKeyPress(ByVal s As Object,
                                                              ByVal e As System.Windows.Forms.KeyPressEventArgs)

            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return

            REM if not number and is readable character then dont allow
            If Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        Public Shared Sub addOrRemove__AlphabetsWithSymbols__addressOfKeyPress(ByVal s As Object,
                                                              ByVal e As System.Windows.Forms.KeyPressEventArgs)



            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsSymbol(e.KeyChar) AndAlso
                Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL




        End Sub



        Public Shared Sub addOrRemove__AlphabetsWithSymbolsAndSpace__addressOfKeyPress(ByVal s As Object,
                                                              ByVal e As System.Windows.Forms.KeyPressEventArgs)



            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return

            REM if not number and is readable character then dont allow
            If Not IsSymbol(e.KeyChar) AndAlso
                Not IsSmallLetter(e.KeyChar) AndAlso
                Not IsBigLetter(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL

          
        End Sub



        Public Shared Sub addOrRemove__SymbolsOnly__addressOfKeyPress(ByVal s As Object,
                                                              ByVal e As System.Windows.Forms.KeyPressEventArgs)



            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)


            REM if not number and is readable character then dont allow
            If Not IsSymbol(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub



        Public Shared Sub addOrRemove__SymbolsWithSpace__addressOfKeyPress(ByVal s As Object,
                                                              ByVal e As System.Windows.Forms.KeyPressEventArgs)



            If s Is Nothing OrElse e Is Nothing Then Return
            Dim sender As Control = CType(s, Control)
            REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            Const NULL As Char = Chr(0)

            If IsSpace(e.KeyChar) Then Return


            REM if not number and is readable character then dont allow
            If Not IsSymbol(e.KeyChar) AndAlso
                       Not IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar) Then e.KeyChar = NULL


        End Sub


        ''Public Shared Sub addOrRemove__ValidateEmailAddress__addressOfValidating(s As Object,
        ''                                                                       e As System.ComponentModel.CancelEventArgs)

        ''    If s Is Nothing OrElse e Is Nothing Then Return
        ''    Dim sender As Control = CType(s, Control)

        ''    If Not IsValidEmail(sender.Text) Then
        ''        errorMsg("You have entered invalid email address. You can clear the control if you don't want to enter any email address.",
        ''                 MESSAGE__POPUP__TIME)
        ''        'sender.Text = String.Empty
        ''        e.Cancel = True
        ''    End If

        ''End Sub







    End Class

End Namespace