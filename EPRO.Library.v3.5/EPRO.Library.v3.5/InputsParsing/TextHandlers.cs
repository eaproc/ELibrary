using System.Windows.Forms;
using static ELibrary.Standard.InputsParsing.TextParsing;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.InputsParsing
{
    public sealed class TextHandlers
    {
        public delegate void dlg_addOrRemove__UseProperCase__addressOfKeyPress(object s, KeyPressEventArgs e, ref int CaretPosition);

        /// <summary>
        /// Meant for single line text
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <param name="CaretPosition"></param>
        /// <remarks></remarks>
        public static void addOrRemove__UseProperCase__addressOfKeyPress(object s, KeyPressEventArgs e, ref int CaretPosition)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            if (CaretPosition > sender.Text.Length)
                return;

            // REM if this is the first char capitalize it
            // REM if this is not and the character @ the end is space then capitalize it
            // REM Else do nothing
            int selStart = CaretPosition;
            sender.Text = Strings.StrConv(sender.Text, VbStrConv.ProperCase);
            CaretPosition = selStart;
            // REM Ignore if it is del or backspace, it will be corrected on next key stroke
            // REM On next stroke if prev text is not conformed to Uppercase then go through each char one by one and correct all
            if (sender.Text.Length == 0 || Strings.Right(sender.Text, 1) == " ")
                e.KeyChar = Strings.UCase(e.KeyChar);
        }


        /// <summary>
        /// Disable Control V
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__DisableControlV__addressOfKeyDown(object s, KeyEventArgs e)
        {
            if (s is null || e is null)
                return;
            if (e.Control && e.KeyCode == Keys.V)
                e.SuppressKeyPress = true;
        }

        /// <summary>
        /// Disable Right Click Menu
        /// </summary>
        /// <param name="s"></param>
        /// <remarks></remarks>
        public static void DisableRightClick(Control s)
        {
            if (s is null)
                return;
            var sender = s;
            if (sender.ContextMenu is null)
                sender.ContextMenu = new ContextMenu();
            // REM Throw New Exception("Just set ShortCut Enabled or set .ContextMenu=new ContextMenu() to False on the Text Box")

        }



        /// <summary>
        /// Makes the text control to accept numbers only. Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersOnly__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithAlphabetsAndSymbols__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (IsSpace(e.KeyChar) || IsTabsAndEnterKeys(e.KeyChar))
                e.KeyChar = NULL;
        }



        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithSinglePlusSignInFront__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';

            // REM Allow plus in front +905
            if (Strings.AscW(e.KeyChar) == PLUS_SIGN && (sender.Text ?? "") == (string.Empty ?? ""))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }



        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithDecimalDot__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';

            // REM Now check if it is the only one
            if (sender.Text.IndexOf(".") >= 0 & Strings.AscW(e.KeyChar) == DOT)
            {
                e.KeyChar = NULL;
                return;
            }

            if (Strings.AscW(e.KeyChar) == DOT)
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys. For French Double Numbers
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithCOMMA__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';

            // REM Now check if it is the only one
            if (sender.Text.IndexOf(',') >= 0 & Strings.AscW(e.KeyChar) == FrenchTextParsing.COMMA)
            {
                e.KeyChar = NULL;
                return;
            }

            if (Strings.AscW(e.KeyChar) == FrenchTextParsing.COMMA)
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        /// <summary>
        /// Doesnt allow tabs and enterkeys. For French Double Numbers
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithCOMMA___ConvertDotToComma__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';

            // REM Now check if it is the only one
            if (sender.Text.IndexOf(',') >= 0 && (Strings.AscW(e.KeyChar) == FrenchTextParsing.COMMA || Strings.AscW(e.KeyChar) == FrenchTextParsing.DOT))
            {
                e.KeyChar = NULL;
                return;
            }

            if (Strings.AscW(e.KeyChar) == FrenchTextParsing.COMMA)
                return;
            if (Strings.AscW(e.KeyChar) == FrenchTextParsing.DOT)
            {
                e.KeyChar = ',';
                return;
            }

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }



        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithAlphabets__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithAlphabetsAndSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__NumbersWithSymbols__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsSymbol(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        public static void addOrRemove__NumbersWithSymbolsAndSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsNumber(e.KeyChar) && !IsSymbol(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__AlphabetsOnly__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        /// <summary>
        /// Doesnt allow tabs and enterkeys
        /// </summary>
        /// <param name="s"></param>
        /// <param name="e"></param>
        /// <remarks></remarks>
        public static void addOrRemove__AlphabetsWithSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        public static void addOrRemove__AlphabetsWithSymbols__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsSymbol(e.KeyChar) && !IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        public static void addOrRemove__AlphabetsWithSymbolsAndSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;

            // REM if not number and is readable character then dont allow
            if (!IsSymbol(e.KeyChar) && !IsSmallLetter(e.KeyChar) && !IsBigLetter(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        public static void addOrRemove__SymbolsOnly__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';


            // REM if not number and is readable character then dont allow
            if (!IsSymbol(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }

        public static void addOrRemove__SymbolsWithSpace__addressOfKeyPress(object s, KeyPressEventArgs e)
        {
            if (s is null || e is null)
                return;
            Control sender = (Control)s;
            // REM Dim KeyUTF16 As Int32 = AscW(e.KeyChar)
            const char NULL = '\0';
            if (IsSpace(e.KeyChar))
                return;


            // REM if not number and is readable character then dont allow
            if (!IsSymbol(e.KeyChar) && !IsNonReadableCharacterExceptTabsAndEnterKey(e.KeyChar))
                e.KeyChar = NULL;
        }


        // 'Public Shared Sub addOrRemove__ValidateEmailAddress__addressOfValidating(s As Object,
        // '                                                                       e As System.ComponentModel.CancelEventArgs)

        // '    If s Is Nothing OrElse e Is Nothing Then Return
        // '    Dim sender As Control = CType(s, Control)

        // '    If Not IsValidEmail(sender.Text) Then
        // '        errorMsg("You have entered invalid email address. You can clear the control if you don't want to enter any email address.",
        // '                 MESSAGE__POPUP__TIME)
        // '        'sender.Text = String.Empty
        // '        e.Cancel = True
        // '    End If

        // 'End Sub







    }
}