using System;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard
{
    public class EMaths
    {






        /// <summary>
    /// Pure decimal NOT string formatted. e.g. 9998,00 or 9998.00
    /// </summary>
    /// <param name="xTurkishEntry"></param>
    /// <returns>String because if the OS is Turkish and I use inbuilt format it will still be in Turkish</returns>
    /// <remarks></remarks>

        public static string Convert_Turkish_Decimal_to_Eng(string xTurkishEntry)
        {
            if (xTurkishEntry.IndexOf(",") >= 0)
            {
                var SplittedVal = Strings.Split(xTurkishEntry, ",");
                return string.Format("{0}.{1}", SplittedVal);
            }

            return xTurkishEntry;
        }

        /// <summary>
    /// Writes a Byte Value in Kb, Mb etc
    /// </summary>
    /// <param name="ByteValue"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static string getReadableByteValue(long ByteValue)
        {
            if (ByteValue > 1024 * 1024)
            {
                return Strings.FormatNumber(ByteValue / (double)(1024 * 1024), 2) + "Mb";
            }
            else
            {
                return Strings.FormatNumber(ByteValue / (double)1024, 2) + "Kb";
            }
        }


        /// <summary>
    /// Alternate Bit Value. If 0 returns 1.. If 1 returns 0
    /// </summary>
    /// <param name="BitVal"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static byte AlternateBit(byte BitVal)
        {
            // If it is greater than 0 it is true 
            // First neutralize the current sign
            return (byte)Math.Abs(Conversions.ToInteger(!Conversions.ToBoolean(Math.Abs(BitVal))));
        }

        /// <summary>
    /// Fetches a number between the specified range
    /// </summary>
    /// <param name="x"></param>
    /// <param name="y"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static int getANumberBetween(int x, int y)
        {
            VBMath.Randomize();
            {
                var withBlock = new Random();
                return withBlock.Next(60, 120);
            }
        }

        /// <summary>
    /// Fetch the Integer Portion ONLY of a Value.. NB: Sign is included
    /// </summary>
    /// <param name="dblVal"></param>
    /// <returns></returns>
    /// <remarks></remarks>
        public static int getIntegerPortionOf(string dblVal)
        {
            try
            {


                // REM Instr is based 1 function and returns 0 if not found
                if (Strings.InStr(dblVal, ".") > 0)
                {
                    return (int)Math.Round(Conversion.Val(dblVal.Substring(0, Strings.InStr(dblVal, "."))));
                }
                else
                {
                    return (int)Math.Round(Conversion.Val(dblVal));
                }
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public static double getCurrencyValue(string strValue)
        {
            if (string.IsNullOrEmpty(strValue))
                return 0d;
            if (Information.IsNumeric(Strings.Left(strValue, 1)) | Strings.Len(strValue) == 1)
                return Conversion.Val(strValue);
            return Conversion.Val(Strings.Right(strValue, Strings.Len(strValue) - 1));
        }

        public static string setCurrencyValue(double dblCurrency, string strCurrencyTpe = "$")
        {
            return strCurrencyTpe + Strings.FormatNumber(dblCurrency, 2);
        }

        public static double getPercentage(double ActualValue, double Total)
        {
            try
            {
                return Conversions.ToDouble(Strings.FormatNumber(ActualValue / Total * 100d, 2));
            }
            catch (Exception ex)
            {
                return 0d;
            }
        }

        public static double divide(double pUpperNumerator, double pDownDenominator)
        {
            if (pDownDenominator == 0d)
                return 0d;
            return pUpperNumerator / pDownDenominator;
        }

        public static decimal divide(decimal pUpperNumerator, decimal pDownDenominator)
        {
            if (pDownDenominator == 0m)
                return 0m;
            return pUpperNumerator / pDownDenominator;
        }
    }
}