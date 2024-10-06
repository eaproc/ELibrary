using System;
using System.Linq;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.InputsParsing
{
    public sealed class TextParsing
    {



        /// <summary>
        /// . in Ascii is 46
        /// </summary>
        /// <remarks></remarks>
        public const byte DOT = 46;

        /// <summary>
        /// For phone numbers
        /// </summary>
        /// <remarks></remarks>
        public const byte PLUS_SIGN = 43;

        public static bool IsSmallLetter(char c)
        {
            // REM Char is UTF-16
            // REM  ' ''97-122 small letters [a - z]
            switch (c)
            {
                case var @case when 'a' <= @case && @case <= 'z':
                    {
                        return true;
                    }
            }

            return false;
        }

        public static bool IsBigLetter(char c)
        {

            // REM   ' ''65-90  Big Letters [A - Z]
            switch (c)
            {
                case var @case when 'A' <= @case && @case <= 'Z':
                    {
                        return true;
                    }
            }

            return false;
        }

        /// <summary>
        /// Currently this isn't perfect on Other Encoding Apart from ASCII because any char not ASCII letters, 
        /// numbers, space, non readable are considered symbols. Therefore, other special chars too are symbols
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsSymbol(char c)
        {
            return !IsBigLetter(c) && !IsSmallLetter(c) && !IsNonReadableCharacter(c) && !IsNumber(c) && !IsSpace(c);
        }


        /// <summary>
        /// Includes Tabs and Enter Keys. Use this if you want to allow Tabs and Enter Keys
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsNonReadableCharacter(char c)
        {

            // REM  ' ''0 - 31
            switch (c)
            {
                case var @case when '\0' <= @case && @case <= '\u001f':
                    {
                        return true;
                    }
            }

            return false;
        }

        /// <summary>
        /// without counting 9,10,11,13 [Tabs and Enter Keys]. You can use this if you don't want to allow those stuff
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsNonReadableCharacterExceptTabsAndEnterKey(char c)
        {

            // REM  ' ''0 - 31 without 9,10,11,13 
            switch (c)
            {
                case var @case when '\0' <= @case && @case <= '\b':
                case '\f':
                case var case1 when '\u000e' <= case1 && case1 <= '\u001f':
                    {
                        return true;
                    }
            }

            return false;
        }

        public static bool IsNumber(char c)
        {

            // ''48 - 57 = > Digits [0 - 9]
            switch (c)
            {
                case var @case when '0' <= @case && @case <= '9':
                    {
                        return true;
                    }
            }

            return false;
        }

        public static bool IsValidEmail(string pEmail)
        {
            try
            {
                var p = new System.Net.Mail.MailAddress(pEmail);
                if ((p.User ?? "") == (string.Empty ?? ""))
                    return false;
                if (p.User.Length < 1 || p.Host.Length < 4)
                    return false;
                if (pEmail.IndexOf(".@") >= 0)
                    return false;
                if (pEmail.IndexOf("..") >= 0)
                    return false;
                if (pEmail.IndexOf(".@") >= 0)
                    return false;
                if (pEmail.IndexOf("@.") >= 0)
                    return false;
                if (pEmail.IndexOf(" ") >= 0)
                    return false;
                if (pEmail.IndexOf("*") >= 0)
                    return false;
                if (pEmail.IndexOf("[]") >= 0)
                    return false;
                if (pEmail.IndexOf("[[") >= 0)
                    return false;
                if (pEmail.IndexOf("]]") >= 0)
                    return false;
                if (pEmail.IndexOf(",") >= 0)
                    return false;
                if (pEmail.IndexOf(";") >= 0)
                    return false;
                if (pEmail.IndexOf(":") >= 0)
                    return false;
                if (pEmail.IndexOf("/") >= 0)
                    return false;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsSpace(char c)
        {
            return c == ' ';
        }

        /// <summary>
        /// Returns if this is enter key or tab
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsTabsAndEnterKeys(char c)
        {
            switch (c)
            {
                case var @case when '\t' <= @case && @case <= '\v':
                case '\r':
                    {
                        return true;
                    }

                default:
                    {
                        return false;
                    }
            }
        }






        /// <summary>
        /// Parse out all numbers from the left of a string
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string parseOutIntegers(string pVal)
        {
            if (pVal is null || (pVal ?? "") == (string.Empty ?? ""))
                return string.Empty;
            var pValid = from dChar in pVal.ToCharArray()
                         where IsNumber(dChar)
                         select dChar;


            // Consider Negative Sign
            string pSymbol = string.Empty;
            int pIndexOfNegative = pVal.IndexOf("-");
            if (pIndexOfNegative >= 0 && pValid.Count() > 0)
            {
                if (pVal.IndexOf(pValid.First()) > pIndexOfNegative)
                {
                    pSymbol = "-";
                }
            }
            // -----------------------------------------


            return pSymbol + new string(pValid.ToArray());
        }


        /// <summary>
        /// Parse out double from string. reading to the last char
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static double parseOutDouble(string pVal)
        {
            return Objects.EDouble.valueOf(parseOutDoubleAsString(pVal));
        }



        /// <summary>
        /// Parse out double from string. reading to the last char
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string parseOutDoubleAsString(string pVal)
        {
            if (pVal is null || (pVal ?? "") == (string.Empty ?? ""))
                return pVal;
            var pValid = from dChar in pVal.ToCharArray()
                         where IsNumber(dChar) || Strings.AscW(dChar) == DOT
                         select dChar;



            // Consider Negative Sign
            string pSymbol = string.Empty;
            int pIndexOfNegative = pVal.IndexOf("-");
            if (pIndexOfNegative >= 0 && pValid.Count() > 0)
            {
                if (pVal.IndexOf(pValid.First()) > pIndexOfNegative)
                {
                    pSymbol = "-";
                }
            }
            // -----------------------------------------




            var parsedString = new string(pValid.ToArray()).Split('.');
            if (parsedString.Length > 1)
                return string.Format("{2}{0}.{1}", parsedString[0], parsedString[1], pSymbol);
            if (parsedString.Length == 1)
                return string.Format("{1}{0}", parsedString[0], pSymbol);
            return string.Empty;
        }
    }
}