using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{

    /// <summary>
    /// Controls the String Class
    /// </summary>
    /// <remarks></remarks>
    public class EStrings
    {

        /// <summary>
        /// For quotation mark
        /// </summary>
        /// <remarks></remarks>
        public const string QUOTE = "\"";



        /// <summary>
        /// Converts DBNUll to System.String.Empty()
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string valueOf(DBNull obj)
        {
            return string.Empty;
        }



        /// <summary>
        /// Converts Objects to String. Returns System.String.Empty() if it is nothing
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string valueOf(object obj)
        {
            if (obj is null)
                return string.Empty;
            if (obj is DBNull)
                return valueOf((DBNull)obj);
            return obj.ToString();
        }







        /// <summary>
        /// Checks if String1 equals String2 Ignoring the Case Sensitivity
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool equalsIgnoreCase(string str1, string str2)
        {
            if (str1 is null && str2 is null)
                return true;
            if (str1 is null || str2 is null)
                return false;
            return str1.ToLower().Equals(str2.ToLower());
        }

        public static new bool Equals(object obj1, object obj2)
        {
            return (valueOf(obj1) ?? "") == (valueOf(obj2) ?? "");
        }

        public static new bool Equals(object obj1, object obj2, bool IgnoreCase)
        {
            if (!IgnoreCase)
                return Equals(obj1, obj2);
            return equalsIgnoreCase(valueOf(obj1), valueOf(obj2));
        }






        /// <summary>
        /// Checks if String2 is in String1 Ignoring the Case Sensitivity
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool LikeIgnoreCase(string str1, string str2)
        {
            return str1.ToLower().IndexOf(str2.ToLower()) >= 0;
        }

        /// <summary>
        /// Checks if String2 is in String1 Case Sensitivity
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool Like(string str1, string str2)
        {
            return str1.IndexOf(str2) >= 0;
        }


        /// <summary>
        /// Checks absolutely if a string is empty
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool isEmpty(string value)
        {
            if (value is null)
                return true;
            if (string.IsNullOrEmpty(value))
                return true;
            if (string.IsNullOrEmpty(value))
                return true;
            if (value.Trim().Equals(string.Empty))
                return true;
            return false;
        }



        /// <summary>
        /// Replaces the occurrence of space in the string with NULL
        /// </summary>
        /// <param name="strExpression"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string RemoveAllSpaces(string strExpression)
        {
            return Strings.Replace(strExpression, " ", "");
        }


        /// <summary>
        /// Quote String. Sample Man ==> "Man"
        /// </summary>
        /// <param name="Expression"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string QuoteString(string Expression)
        {
            return string.Format("{0}{1}{0}", '"', Expression);
        }


        /// 
        /// <summary>
        /// If Original Length is Greater Than The Specified Length,
        ///     It returns the same string
        /// Otherwise, if less than, String is buffered with specified character. Same Function as String.PadLeft or String.PadRight
        /// 
        /// <example>Wraps up a string to a Specified Length.  
        /// <code>
        ///    [WrapUp("Man",4,"A")  = "ManA"]   
        ///    [WrapUp("Man",3,"A")  = "Man".]   
        ///    [WrapUp("Man",2,"A")  = "Man". ]
        /// </code>
        /// </example>
        /// </summary>
        public static string WrapUpAtLeast(string OriginalString, int ReturnLength, char BufferedCharacter = '0', bool BackPadding = false)
        {
            try
            {
                if (Strings.Len(OriginalString) < ReturnLength)
                {
                    if (BackPadding)
                    {
                        return OriginalString + Strings.StrDup(ReturnLength - Strings.Len(OriginalString), BufferedCharacter);
                    }
                    else
                    {
                        return Strings.StrDup(ReturnLength - Strings.Len(OriginalString), BufferedCharacter) + OriginalString;
                    }
                }
            }

            // 'If Len(OriginalString) > ReturnLength Then
            // '    Return Left(OriginalString, Len(OriginalString) - ReturnLength)
            // 'End If


            catch (Exception ex)
            {
            }

            return OriginalString;
        }

        /// 
        /// <summary>
        /// If Original Length is Greater Than The Specified Length,
        ///      Length is cut to specified lenth
        /// Otherwise, if less than, String is buffered with specified character. Same Function as String.PadLeft or String.PadRight
        /// 
        /// <example>Wraps up a string to a Specified Length.  
        /// <code>
        ///    [WrapUp("Man",4,"A")  = "ManA"]   
        ///    [WrapUp("Man",3,"A")  = "Man".]   
        ///    [WrapUp("Man",2,"A")  = "Ma". ]
        /// </code>
        /// </example>
        /// </summary>
        public static string WrapUp(string OriginalString, int ReturnLength, char BufferedCharacter = '0', bool BackPadding = false)
        {
            try
            {
                if (Strings.Len(OriginalString) < ReturnLength)
                {
                    if (BackPadding)
                    {
                        return OriginalString + Strings.StrDup(ReturnLength - Strings.Len(OriginalString), BufferedCharacter);
                    }
                    else
                    {
                        return Strings.StrDup(ReturnLength - Strings.Len(OriginalString), BufferedCharacter) + OriginalString;
                    }
                }

                if (Strings.Len(OriginalString) > ReturnLength)
                {
                    return Strings.Left(OriginalString, Strings.Len(OriginalString) - ReturnLength);
                }
            }
            catch (Exception ex)
            {
            }

            return OriginalString;
        }


        /// <summary>
        /// Determines if Version1 precedes Version2. 
        /// If 1 > 2 returns 1.
        /// if 1 = 2 returns 0.
        /// if 1 &lt; 2 returns -1.
        /// </summary>
        /// <param name="version1"></param>
        /// <param name="version2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int CompareVersions(string version1, string version2)
        {
            var SplittedV1 = Strings.Split(version1, ".");
            var SplittedV2 = Strings.Split(version2, ".");
            string[] SplittedTMP;
            int isSwapped = 1;


            // REM Check for the longest length
            if (SplittedV2.Count() > SplittedV1.Count())
            {
                // REM Swapp
                // SplittedTMP = New String() {}
                SplittedTMP = SplittedV1;
                SplittedV1 = SplittedV2;
                SplittedV2 = SplittedTMP;
                SplittedTMP = null;  // REM Free the space
                isSwapped = -1;
            }

            for (int xIndex = 0, loopTo = SplittedV1.Count() - 1; xIndex <= loopTo; xIndex++)
            {
                if (xIndex < SplittedV2.Count())
                {
                    if (Conversion.Val(SplittedV1[xIndex]) > Conversion.Val(SplittedV2[xIndex]))
                    {
                        return 1 * isSwapped;
                    }
                    else if (Conversion.Val(SplittedV1[xIndex]) < Conversion.Val(SplittedV2[xIndex]))
                    {
                        return -1 * isSwapped;
                    }
                }
                else if (Conversion.Val(SplittedV1[xIndex]) > 0d)
                {
                    return 1 * isSwapped;
                }
            }


            // CompareVersions = version1.CompareTo(version2)

            return 0;
        }





        /// <summary>
        /// Splits a string and returns result with empty string element
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="delimiter"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string[] splitWithoutNullElement(string expression, string delimiter = ",")
        {
            var s_split = Strings.Split(expression, delimiter);
            if (string.IsNullOrEmpty(expression))
                return new string[] { };
            if (s_split is null)
                return new string[] { };
            if (s_split.Length == 0)
                return new string[] { };
            var l_split = new List<string>(s_split.Length);
            foreach (string s in s_split)
            {
                if (string.IsNullOrEmpty(s.Trim()))
                    continue;
                l_split.Add(s);
            }

            return l_split.ToArray();
        }


        /// <summary>
        /// Split strings into chunkSize
        /// </summary>
        /// <param name="str"></param>
        /// <param name="chunkSize"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static IEnumerable<string> SplitChunk(string str, int chunkSize)
        {
            return Enumerable.Range(0, (int)Math.Round(str.Length / (double)chunkSize)).Select(i => str.Substring(i * chunkSize, chunkSize));
        }









//        /// <summary>
//        /// Extract String from an expression using the first appearance of the limits. 
//        /// Example: StartString = lstmessagelst ... Stop String = lst/messagelst
//        /// </summary>
//        /// <param name="HTMLString"></param>
//        /// <param name="startString"></param>
//        /// <param name="stopString"></param>
//        /// <param name="locatePosition"></param>
//        /// <returns></returns>
//        /// <remarks></remarks>
//        public static string ExtractStringFromHtml(string HTMLString, string startString, string stopString, [Optional, DefaultParameterValue(1d)] ref double locatePosition)
//        {
//            ;
//#error Cannot convert OnErrorGoToStatementSyntax - see comment for details
//            /* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToLabelStatement not implemented, please report this issue in 'On Error GoTo errHandler' at character 13739


//            Input:

//                        On Error GoTo errHandler

//             */
//            long fLocate;
//            long rLocate;
//            fLocate = Strings.InStr((int)Math.Round(locatePosition), HTMLString, startString, CompareMethod.Text);
//            if (fLocate != 0L)
//            {
//                rLocate = Strings.InStr((int)fLocate, HTMLString, stopString, CompareMethod.Text);

//                // Add String Position
//                fLocate = fLocate + Strings.Len(startString);

//                // ExtractStringFromHtml = Mid(HTMLString, fLocate, rLocate - fLocate)

//                // Return Stop Position
//                locatePosition = rLocate;
//                return Strings.Mid(HTMLString, (int)fLocate, (int)(rLocate - fLocate));
//            }

//            return Constants.vbNullString;
//            return default;
//        errHandler:
//            ;
//        }


        /// <summary>
        /// Fetch first letter of the string and abbreviate it
        /// </summary>
        /// <param name="strString"></param>
        /// <param name="Capitalized"></param>
        /// <param name="AbbvSymbol"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string getFirstAbbreviatedLetter(string strString, bool Capitalized = true, string AbbvSymbol = ".")
        {
            if (strString.Length < 1)
                return "";
            string AbbvStr = strString.Substring(0, 1).ToLower();
            if (Capitalized)
                AbbvStr = AbbvStr.ToUpper();
            return AbbvStr + AbbvSymbol;
        }



        /// <summary>
        /// Uses this syntax *:=0 instead of the syntax of string.format {0}. [?Microsoft.VisualBasic.Strings.Format()] -Dont Confuse it
        /// </summary>
        /// <param name="Expression">The expression that contains</param>
        /// <param name="args"></param>
        /// <returns></returns>
        /// <remarks>Becareful so you dont confuse it with [?Microsoft.VisualBasic.Strings.Format()]</remarks>
        public static string FormatString(string Expression, params string[] args)
        {
            var loopInc = default(byte);
            foreach (string arg in args)
            {
                Expression = Strings.Replace(Expression, string.Format("*:={0}", loopInc), arg);
                loopInc = (byte)(loopInc + 1);
            }

            return Expression;
        }



        /// <summary>
        /// Removes all non-readable characters from strings range [0 - 31]
        /// </summary>
        /// <param name="ORIGINAL_STRING"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string CLEAN_UP_STRING(string ORIGINAL_STRING)
        {
            string CLEAN_UP_STRINGRet = default;
            byte byteCol;
            // Remove all unwanted keys 0-31
            for (byteCol = 0; byteCol <= 31; byteCol++)

                // Prevent Freeze
                // DoEvents()
                ORIGINAL_STRING = Strings.Replace(ORIGINAL_STRING, Conversions.ToString(Strings.Chr(byteCol)), "");

            // incase null was added, remove it again
            CLEAN_UP_STRINGRet = Strings.Replace(ORIGINAL_STRING, Conversions.ToString('\0'), "");
            return CLEAN_UP_STRINGRet;
        }



        /// <summary>
        /// Reverse the sequence of a string from back to front
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Reverse(string value)
        {
            var arr = value.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }


        /// <summary>
        /// Dont use unless necessary. 
        /// </summary>
        /// <param name="Str"></param>
        /// <returns>Returns * 3 the length of string</returns>
        /// <remarks></remarks>
        public static string EncodeToFixLengthBytes(string Str)
        {
            string str_bytes = Constants.vbNullString;
            foreach (char c in Str.ToCharArray())
                str_bytes += WrapUp(Strings.Asc(c).ToString(), 3);
            return str_bytes;
        }

        /// <summary>
        /// Reverse the process of EncodeTOFixLengthBytes
        /// </summary>
        /// <param name="Str"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string Reverse_EncodeToFixLengthBytes(string Str)
        {
            if (Strings.Len(Str) % 3 != 0)
                return Constants.vbNullString;
            string rtr_str = Constants.vbNullString;
            short i = 0;
            try
            {
                var loopTo = (short)(Strings.Len(Str) - 1);
                for (i = 0; i <= loopTo; i += 3)
                    rtr_str += Conversions.ToString(Strings.Chr(EInt.valueOf(Str.Substring(i, 3))));
            }
            catch (Exception ex)
            {
            }

            return rtr_str;
        }





        /// <summary>
        /// Checks if whole String is combination of escape characters like \r\n
        /// </summary>
        /// <param name="pVal"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool isEscapeCharacters(string pVal)
        {
            if (pVal is null || (pVal ?? "") == (string.Empty ?? ""))
                return false;
            var IRst = from ds in pVal.Split('\\')
                       where ds == "n" || ds == "r" || ds == "t" || ds == "b" || ds == "f" || ds == "'" || ds == @"\" || ds == "\""
                       select ds;
            return IRst.Count() * 2 == pVal.Length;
        }


        /// <summary>
        /// Converts Combination of escape characters to String else returns empty strings
        /// </summary>
        /// <param name="pEscapeChars"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string translateEscapeCharacters(string pEscapeChars)
        {
            if (!isEscapeCharacters(pEscapeChars))
                return string.Empty;
            string rs = string.Empty;
            foreach (string ds in pEscapeChars.Split('\\'))
            {
                switch (ds ?? "")
                {
                    case "n":
                        {
                            rs += Conversions.ToString('\n');
                            break;
                        }

                    case "r":
                        {
                            rs += Conversions.ToString('\r');
                            break;
                        }

                    case "t":
                        {
                            rs += Conversions.ToString('\t');
                            break;
                        }

                    case "b":
                        {
                            rs += Conversions.ToString('\b');
                            break;
                        }

                    case "f":
                        {
                            rs += Conversions.ToString('\f');
                            break;
                        }

                    case "\"":
                        {
                            rs += Conversions.ToString('"');
                            break;
                        }

                    case "'":
                        {
                            rs += Conversions.ToString('\'');
                            break;
                        }

                    case @"\":
                        {
                            rs += Conversions.ToString('\\');
                            break;
                        }
                }
            }

            return rs;
        }








        //// 
        //// Base64
        //// 
        ///// <summary>
        ///// It handles exception
        ///// </summary>
        ///// <param name="pString"></param>
        ///// <returns></returns>
        ///// <remarks></remarks>
        //public static string ConvertToBase64(string pString)
        //{
        //    return ConvertToBase64(pString, System.Text.Encoding.UTF8);
        //}
        ///// <summary>
        ///// It handles exception
        ///// </summary>
        ///// <param name="pString"></param>
        ///// <param name="pEncoding"></param>
        ///// <returns></returns>
        ///// <remarks></remarks>
        //public static string ConvertToBase64(string pString, System.Text.Encoding pEncoding)
        //{
        //    try
        //    {
        //        return Convert.ToBase64String(pEncoding.GetBytes(pString));
        //    }
        //    catch (Exception ex)
        //    {
        //        Program.Logger.Print(ex);
        //        return string.Empty;
        //    }
        //}

        ///// <summary>
        ///// It handles exception
        ///// </summary>
        ///// <param name="pString"></param>
        ///// <returns></returns>
        ///// <remarks></remarks>
        //public static string ConvertFromBase64(string pString)
        //{
        //    return ConvertFromBase64(pString, System.Text.Encoding.UTF8);
        //}

        ///// <summary>
        ///// It handles exception
        ///// </summary>
        ///// <param name="pString"></param>
        ///// <param name="pEncoding"></param>
        ///// <returns></returns>
        ///// <remarks></remarks>
        //public static string ConvertFromBase64(string pString, System.Text.Encoding pEncoding)
        //{
        //    try
        //    {
        //        pString = pString.Trim();
        //        if ((pString ?? "") == (string.Empty ?? ""))
        //            return pString;
        //        int pMode = pString.Length % 4;
        //        if (pMode != 0)
        //            pString = pString.PadRight(pString.Length + 4 - pMode, '=');
        //        return pEncoding.GetString(Convert.FromBase64String(pString));
        //    }
        //    catch (Exception ex)
        //    {
        //        Program.Logger.Print(ex);
        //        return string.Empty;
        //    }
        //}
    }
}