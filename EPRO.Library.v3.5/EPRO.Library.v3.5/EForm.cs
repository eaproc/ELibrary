using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using CODERiT.Logger.v._3._5.Exceptions;
using ELibrary.Standard.Modules;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    public class EForm
    {


        #region New Form Display

        /// <summary>
        /// Display a Modal Child Form and return the dialog result
        /// </summary>
        /// <param name="Parent"></param>
        /// <param name="Child"></param>
        /// <param name="FadeParentOpaq"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DialogResult ShowChildFormDialog(ref Form Parent, Form Child, double FadeParentOpaq = 0.7d)
        {
            if (!isValid_Form(Parent))
                return DialogResult.No;
            double PrevOpaq = Parent.Opacity;
            try
            {
                DialogResult dlgRst;
                Parent.Opacity = FadeParentOpaq;
                Child.ShowInTaskbar = false;
                dlgRst = Child.ShowDialog(Parent);
                Parent.Opacity = PrevOpaq;
                return dlgRst;
            }
            catch (Exception ex)
            {
                if (isValid_Form(Parent))
                    Parent.Opacity = PrevOpaq;
                basMain.MyLogFile.Log(new EException(ex));
                return DialogResult.No;
            }
        }


        /// <summary>
        /// Display a Modal Child Form  on any control handle  it doesnt fade .. .and return the dialog result. User Control Doesn't support Opacity in .NET 3.5
        /// </summary>
        /// <param name="Parent"></param>
        /// <param name="Child"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DialogResult ShowChildFormDialog<T, T2>(ref T Parent, T2 Child)
            where T : UserControl
            where T2 : Form
        {
            try
            {
                DialogResult dlgRst;
                Child.ShowInTaskbar = false;
                dlgRst = Child.ShowDialog(Parent);
                return dlgRst;
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Log(new EException(ex));
                return DialogResult.No;
            }
        }

        /// <summary>
        /// Display a Modal Child Form  on any control handle  it doesnt fade .. .and return the dialog result. User Control Doesn't support Opacity in .NET 3.5
        /// </summary>
        /// <param name="Parent"></param>
        /// <param name="Child"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool ShowChildFormNonModal<T, T2>(ref T Parent, T2 Child, bool CloseParentForm = true)
            where T : UserControl
            where T2 : Form
        {
            try
            {
                Child.Show();
                if (Parent is object && CloseParentForm && !Parent.IsDisposed)
                    Parent.Dispose();
                return true;
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Log(new EException(ex));
                return false;
            }
        }

        /// <summary>
        /// Opens a non modal child form
        /// </summary>
        /// <param name="ParentForm"></param>
        /// <param name="child"></param>
        /// <param name="CloseParentForm"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool showChildFormNonModal(ref Form ParentForm, Form child, bool CloseParentForm = true)
        {
            try
            {

                // Debug.Print("CloseParentForm " & CloseParentForm)
                // Debug.Print("ParentForm IsNot Nothing" & (ParentForm IsNot Nothing))

                child.Show();
                if (CloseParentForm && isValid_Form(ParentForm))
                {
                    // Debug.Print("Closing Parent Form")
                    ParentForm.Close();
                }

                return true;
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Log(new EException(ex));
            }

            return false;
        }

        /// <summary>
        /// Opens a non modal child form and does not close the parent form
        /// </summary>
        /// <param name="child"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool showChildFormNonModal(Form child)
        {
            try
            {
                child.Show();
                return true;
            }
            catch (Exception ex)
            {
            }

            return false;
        }

        // ' ''' <summary>
        // ' ''' Opens a non modal child form
        // ' ''' </summary>
        // ' ''' <param name="ParentForm"></param>
        // ' ''' <param name="child"></param>
        // ' ''' <param name="CloseParentForm"></param>
        // ' ''' <remarks>This is running on same thread and the performance is poor</remarks>
        // 'Public Shared Sub showChildFormNonModal(ByRef ParentForm As Form, ByVal child As Form,
        // '                                                    ByVal WaitTillOpenDescription As String,
        // '                                                    Optional ByVal CloseParentForm As Boolean = True)

        // '    Dim WaitForm As New prjWaitAsyncManager.clsWaitAsyncMgr(, WaitTillOpenDescription)
        // '    Application.DoEvents()
        // '    showChildFormNonModal(ParentForm, child, CloseParentForm)
        // '    WaitForm.Dispose()
        // '    WaitForm = Nothing

        // 'End Sub


        #endregion


        #region Locating (Setting Location and Sizes) Control



        public static void centerControl(ref Control ControlToCenter, Control InReferenceTo, bool Horizontally = true)
        {
            if (Horizontally)
            {
                ControlToCenter.Left = (int)Math.Round((InReferenceTo.Width - ControlToCenter.Width) / 2d);
            }
            else
            {
                ControlToCenter.Top = (int)Math.Round((InReferenceTo.Height - ControlToCenter.Height) / 2d);
            }
        }

        public static void LockControlSize<T>(ref T pControl) where T : Control
        {
            pControl.MinimumSize = pControl.Size;
            pControl.MaximumSize = pControl.Size;
        }

        public enum FormLocationOnScreen
        {
            TOP__LEFT__CORNER,
            TOP__MIDDLE,
            TOP__RIGHT__CORNER,
            MIDDLE__LEFT__CORNER,
            CENTER,
            MIDDLE__RIGHT__CORNER,
            BOTTOM__LEFT__CORNER,
            BOTTOM__MIDDLE,
            BOTTOM__RIGHT__CORNER
        }

        public static void placeFormOnScreenAt<T>(ref T pControl, FormLocationOnScreen pPosition = FormLocationOnScreen.CENTER) where T : Form
        {
            placeFormOnScreenAt(ref pControl, new Padding(0), pPosition);
        }

        public static void placeFormOnScreenAt<T>(ref T pControl, Padding pPadding, FormLocationOnScreen pPosition = FormLocationOnScreen.CENTER) where T : Form
        {
            switch (pPosition)
            {
                case FormLocationOnScreen.TOP__LEFT__CORNER:
                    {
                        pControl.Left = pPadding.Left;
                        pControl.Top = pPadding.Top;
                        break;
                    }

                case FormLocationOnScreen.TOP__MIDDLE:
                    {
                        pControl.Left = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width) / 2d);
                        pControl.Top = pPadding.Top;
                        break;
                    }

                case FormLocationOnScreen.TOP__RIGHT__CORNER:
                    {
                        pControl.Left = My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width - pPadding.Right;
                        pControl.Top = pPadding.Top;
                        break;
                    }

                case FormLocationOnScreen.MIDDLE__LEFT__CORNER:
                    {
                        pControl.Left = pPadding.Left;
                        pControl.Top = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height) / 2d);
                        break;
                    }

                case FormLocationOnScreen.CENTER:
                    {
                        pControl.Left = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width) / 2d);
                        pControl.Top = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height) / 2d);
                        break;
                    }

                case FormLocationOnScreen.MIDDLE__RIGHT__CORNER:
                    {
                        pControl.Left = My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width - pPadding.Right;
                        pControl.Top = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height) / 2d);
                        break;
                    }

                case FormLocationOnScreen.BOTTOM__LEFT__CORNER:
                    {
                        pControl.Left = pPadding.Left;
                        pControl.Top = My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height - pPadding.Bottom;
                        break;
                    }

                case FormLocationOnScreen.BOTTOM__MIDDLE:
                    {
                        pControl.Left = (int)Math.Round((My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width) / 2d);
                        pControl.Top = My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height - pPadding.Bottom;
                        break;
                    }

                case FormLocationOnScreen.BOTTOM__RIGHT__CORNER:
                    {
                        pControl.Left = My.MyProject.Computer.Screen.WorkingArea.Width - pControl.Width - pPadding.Right;
                        pControl.Top = My.MyProject.Computer.Screen.WorkingArea.Height - pControl.Height - pPadding.Bottom;
                        break;
                    }
            }
        }




        #endregion


        /// <summary>
        /// Checks if it is not nothing and not disposed
        /// </summary>
        /// <param name="frm"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool isValid_Form(Form frm)
        {
            return frm is object && !frm.IsDisposed;
        }

        /// <summary>
        /// Fetch GUID from an obj.getType
        /// </summary>
        /// <param name="objType"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string get_guid_from_ObjectType(Type objType)
        {
            return new Guid(((GuidAttribute)objType.Assembly.GetCustomAttributes(typeof(GuidAttribute), false)[0]).Value).ToString();
        }

        /// <summary>
        /// Fetch GUID from an obj.getType
        /// </summary>
        /// <param name="objType"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string get_AppTitle_from_ObjectType(Type objType)
        {
            var arr = objType.Assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyTitleAttribute), true);
            string AppTitle = string.Empty;
            if (arr is object && arr.Length > 0)
            {
                AppTitle = ((System.Reflection.AssemblyTitleAttribute)arr[0]).Title;
            }

            return AppTitle;
        }


        /// <summary>
        /// Force close all application open forms and ignore error
        /// </summary>
        /// <remarks></remarks>
        public static void ClosedAllApplicationForms()
        {
            try
            {
                for (int intC = 0, loopTo = Application.OpenForms.Count - 1; intC <= loopTo; intC++)
                    Application.OpenForms[0].Close();
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Print(ex);
            }
        }


        #region Variables Declarations
        public const DateFormat TimeFormatUsedWithoutSeconds = DateFormat.ShortTime;
        public const string DateFormatUsed = "dd/MMM/yyyy";

        // ''97-122 small letters [a - z]
        public static byte[] SmallLetters = new byte[] { 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122 };

        // ''65-90  Big Letters [A - Z]
        public static byte[] BigLetters = new byte[] { 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90 };



        // ''48 - 57 = > Digits [0 - 9]
        public static byte[] Numbers = new byte[] { 48, 49, 50, 51, 52, 53, 54, 55, 56, 57 };


        // ''Space=>32
        public const byte SpaceBarCode = 32;

        #endregion

        #region EncodingAndDecoding
        private static byte[] UnicodeStringToBytes(string str)
        {
            return System.Text.Encoding.Unicode.GetBytes(str);
        }

        private static string UnicodeBytesToString(byte[] bytes)
        {
            return System.Text.Encoding.Unicode.GetString(bytes);
        }

        private static string UTF8BytesToString(byte[] bytes)
        {
            return System.Text.Encoding.UTF8.GetString(bytes);
        }
        #endregion

        #region Validation Rules

        public static bool ValidateShortText(ref string objInput, ref string objErrMessage, long MinimumLength = 3L, long MaximumLength = 50L)
        {
            try
            {
                if (Strings.Len(objInput) < MinimumLength)
                {
                    objErrMessage = string.Format("Input is too small (Minimum of {0} Characters)", MinimumLength);
                    return false;
                }

                if (Strings.Len(objInput) > MaximumLength)
                {
                    objErrMessage = string.Format("Input is too Large (It has been reduced to {0}. Maximum of {0} Characters)", MaximumLength);
                    objInput = Strings.Left(objInput, (int)MaximumLength);
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            return true;
        }

        public static bool ValidateLongText(ref string objInput, ref string objErrMessage, long MinimumLength = 3L, long MaximumLength = 10000L)
        {
            return ValidateShortText(ref objInput, ref objErrMessage, MinimumLength, MaximumLength);
        }

        public static bool ValidateEmail(ref string objInput, ref string objErrMessage)
        {
            // More Rules
            // If there is space within the text 
            // if there is double @
            // if there is up to 3 min letters before and after the @
            // min of 7 letters
            // Allowed Characters _ . @ [0 - 9][a - z]
            // . after @ must atleast 1 dot
            // . must not be the last thing and must not be the first thing
            // 
            if (!ValidateLongText(ref objInput, ref objErrMessage, 7L))
                return false;
            if (objInput.IndexOf(" ") > 0)
            {
                objErrMessage = "Spaces are NOT allowed within an email address.";
                return false;
            }

            int intval = objInput.IndexOf("@");
            if (intval >= 0)
            {
                if (Strings.Len(objInput) > intval + 1)
                {
                    if (objInput.IndexOf("@", intval + 1) >= 0)
                    {
                        objErrMessage = "Maximum of 1 @ Symbol is allowed within an email address.";
                        return false;
                    }
                }
            }
            else
            {
                goto InvalidEmailAddress;
            }

            if (objInput.IndexOf(".@") >= 0)
                goto InvalidEmailAddress;
            if (objInput.IndexOf("@.") >= 0)
                goto InvalidEmailAddress;
            if (objInput.IndexOf("..") >= 0)
                goto InvalidEmailAddress;
            bool localValidateShortText() { string argobjInput = objInput.Substring(0, objInput.IndexOf("@")); var ret = ValidateShortText(ref argobjInput, ref objErrMessage, 4L); return ret; }

            if (!localValidateShortText())
                goto InvalidEmailAddress;
            bool localValidateShortText1() { string argobjInput = objInput.Substring(objInput.IndexOf("@")); var ret = ValidateShortText(ref argobjInput, ref objErrMessage, 4L); return ret; }

            if (!localValidateShortText1())
                goto InvalidEmailAddress;
            if (objInput.Substring(objInput.IndexOf("@")).IndexOf(".") < 0)
                goto InvalidEmailAddress;
            for (int intC = 33; intC <= 126; intC++)
            {
                if (intC == 46)
                    intC = 47;
                if (intC == 48)
                    intC = 58;
                if (intC == 64)
                    intC = 91;
                if (intC == 95)
                    intC = 96;
                if (intC == 97)
                    intC = 123;
                if (objInput.IndexOf(Strings.Chr(intC)) >= 0)
                    goto InvalidEmailAddress;
                Application.DoEvents();
            }

            if (objInput.Substring(0, 1) == ".")
                goto InvalidEmailAddress;
            if (objInput.Substring(Strings.Len(objInput) - 1, 1) == ".")
                goto InvalidEmailAddress;
            return true;
        InvalidEmailAddress:
            ;
            objErrMessage = "Invalid Email address. Hint: Check for Invalid Symbols.";
            return false;
        }

        public static bool ValidateMobileNumber(ref string objInput, ref string objErrMessage, bool AllowPlus = true, long MinimumLength = 8L, long MaximumLength = 14L)
        {
            if (!ValidateShortText(ref objInput, ref objErrMessage, MinimumLength, MaximumLength))
                return false;
            if (AllowPlus)
            {
                if (objInput.IndexOf("+") > 0)
                {
                    objErrMessage = "(+) Sign is wrongly placed! Invalid Number";
                    return false;
                }
            }

            if (AllowPlus)
            {
                if (!Information.IsNumeric(objInput.Substring(1)))
                    goto InvalidMobileNumber;
            }
            else if (!Information.IsNumeric(objInput))
                goto InvalidMobileNumber;
            return true;
            return default;
        InvalidMobileNumber:
            ;
            objErrMessage = "Invalid Mobile Number. [Numbers Only!]";
            return false;
        }

        public static bool ValidateNumeric(ref string objInput, ref string objErrMessage, long MinimumLength = 1L, long MaximumLength = 20L)
        {
            if (!ValidateMobileNumber(ref objInput, ref objErrMessage, false, MinimumLength, MaximumLength))
            {
                objErrMessage = "Invalid Number(s). Only 0 - 9 are allowed!!!";
                return false;
            }

            return true;
        }

        public static bool ValidateCurrency(ref string objInput, ref string objErrMessage, long MinimumLength = 1L, long MaximumLength = 20L)
        {
            return ValidateNumeric(ref objInput, ref objErrMessage, MinimumLength, MaximumLength);
        }

        public static bool ValidateDate(ref string objInput, ref string objErrMessage, string DateFormatAccepted = DateFormatUsed)
        {
            try
            {
                objInput = Strings.Format(Conversions.ToDate(objInput), DateFormatAccepted);
            }
            catch (Exception ex)
            {
                objErrMessage = string.Format("Invalid Date Entry. Format [{0}]", DateFormatAccepted);
                return false;
            }

            return true;
        }

        public static bool ValidateDate(ref string objInput, ref string objErrMessage, DateFormat DateFormatAccepted)
        {
            try
            {
                objInput = Strings.FormatDateTime(Conversions.ToDate(objInput), DateFormatAccepted);
            }
            catch (Exception ex)
            {
                objErrMessage = string.Format("Invalid Date/Time Entry. Format [{0}]", DateFormatAccepted.ToString());
                return false;
            }

            return true;
        }

        public static bool ValidateTime(ref string objInput, ref string objErrMessage, DateFormat TimeFormatAccepted = TimeFormatUsedWithoutSeconds)
        {
            if (!ValidateDate(ref objInput, ref objErrMessage, TimeFormatAccepted))
            {
                objErrMessage = "Invalid Time Format";
                return false;
            }

            return true;
        }




        /// <summary>
        /// Allows Leters [a - Z] and Space Only
        /// </summary>
        /// <param name="objInput"></param>
        /// <param name="objErrMessage"></param>
        /// <param name="MinimumLength"></param>
        /// <param name="MaximumLength"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool ValidateLettersAndSpaceOnly(ref string objInput, ref string objErrMessage, long MinimumLength = 1L, long MaximumLength = 50L)
        {

            // '' '' '' ''Dim smallLetters As String = ""
            // '' '' '' ''For intC As Integer = 48 To 57
            // '' '' '' ''    smallLetters &=
            // '' '' '' ''                String.Format(
            // '' '' '' ''                        "CByte({0})", intC
            // '' '' '' ''                        )

            // '' '' '' ''    If 57 > intC Then smallLetters &= ","
            // '' '' '' ''Next

            // '' '' '' ''Debug.Print(smallLetters)

            try
            {
                // Ok, Here We correct the length 

                if (Strings.Len(objInput) < MinimumLength)
                {
                    objErrMessage = string.Format("Input is too small (Minimum of {0} Characters)", MinimumLength);
                    return false;
                }

                if (Strings.Len(objInput) > MaximumLength)
                {
                    objErrMessage = string.Format("Input is too Large (It has been reduced to {0}. Maximum of {0} Characters)", MaximumLength);
                    objInput = Strings.Left(objInput, (int)MaximumLength);
                    return false;
                }
                // ************************************************

                // 'We assume the objInput carries the readable format. Allows Big Letters, Small Letters and Space
                // 'Now we verify the contents

                // Returns Strings with each character separated by vbNull
                var Contents = UnicodeStringToBytes(objInput);
                var AllowedCharacters = new List<byte>();
                AllowedCharacters.AddRange(BigLetters);
                AllowedCharacters.AddRange(SmallLetters);
                AllowedCharacters.AddRange(new[] { SpaceBarCode, (byte)Strings.Asc(Constants.vbNullChar) });
                var interceptContents = AllowedCharacters.ToArray();
                var CharsNotAllowed = Contents.Except(interceptContents);


                // If this statement is true then the validation is perfect
                if (!(CharsNotAllowed.Count() <= 0))
                {
                    // '    'Valid
                    // '    Return True
                    // 'Else
                    objErrMessage = string.Format("Allows Leters [a - Z] and Space Only. {0}The following characters are not allowed: {0}{1}", Constants.vbCrLf, UTF8BytesToString(CharsNotAllowed.ToArray()));

                    // ''Debug.Print(
                    // ''    System.Text.Encoding.UTF8.GetString(CharsNotAllowed.ToArray)
                    // ''    )
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            // Valid
            return true;
        }


        #endregion


        #region Enforcing Form to TopMost View
        private static void ActivateFormLater(object sender, EventArgs e)
        {
            try
            {
                Timer pSender = (Timer)sender;
                pSender.Stop();
                Form pForm = (Form)pSender.Tag;
                if (pForm is object && !pForm.IsDisposed)
                {
                    pForm.Activate();
                }
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Print(ex);
            }
        }

        private static void RemoveFromTopmost(object sender, EventArgs e)
        {
            try
            {
                Timer pSender = (Timer)sender;
                pSender.Stop();
                // pSender.Enabled = False
                Form pForm = (Form)pSender.Tag;
                if (pForm is object && !pForm.IsDisposed)
                    pForm.TopMost = false;
            }
            catch (Exception ex)
            {
                basMain.MyLogFile.Print(ex);
            }
        }

        public static void RemoveFormFromTopmostIn(int pMilliseconds, Form pForm)
        {
            var p = new Timer();
            p.Interval = pMilliseconds;
            p.Tag = pForm;
            p.Tick += RemoveFromTopmost;
            p.Start();
        }

        public static void AddFormToTopmostIn(int pMilliseconds, Form pForm)
        {
            var p = new Timer();
            p.Interval = pMilliseconds;
            p.Tag = pForm;
            p.Tick += ActivateFormLater;
            pForm.TopMost = true;
            p.Start();
        }

        #endregion


    }
}