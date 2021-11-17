using System.Linq;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.InputsParsing
{
    public sealed class FrenchTextParsing
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
        public const byte COMMA = 44;


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
                         where TextParsing.IsNumber(dChar) || Strings.AscW(dChar) == COMMA
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


            var parsedString = new string(pValid.ToArray()).Split(',');
            if (parsedString.Length > 1)
                return string.Format("{3}{0}{1}{2}", parsedString[0], ',', parsedString[1], pSymbol);
            if (parsedString.Length == 1)
                return string.Format("{1}{0}", parsedString[0], pSymbol);
            return string.Empty;
        }
    }
}