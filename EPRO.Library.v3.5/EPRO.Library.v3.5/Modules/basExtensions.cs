using ELibrary.Standard.Objects;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.Modules
{
    public static class basExtensions
    {

        #region MathsFunctions

        /// <summary>
        /// Toggles Integer Values e.g betw
        /// </summary>
        /// <param name="current"></param>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int ToggleValue(this int current, int min, int max)
        {
            if (min > max || current != max & current != min)
                return min;
            return min + (max - current);
        }


        #endregion


        #region Object Conversions
        // REM Dont put object as extension

        public static int toInt32(this object val)
        {
            return EInt.valueOf(val);
        }

        public static long toLong(this object val)
        {
            return ELong.valueOf(val);
        }

        public static double toDouble(this object val)
        {
            return EDouble.valueOf(val);
        }

        public static decimal toDecimal(this object val)
        {
            return EDecimal.valueOf(val);
        }

        /// <summary>
        /// Converts an Any Object to Short Type
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short toShort(this object val)
        {
            return EShort.valueOf(obj: val);
        }

        /// <summary>
        /// Converts an Any Object to Boolean
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool toBoolean(this object val)
        {
            return EBoolean.valueOf(obj: val);
        }

        #endregion


        #region String Functions

        /// <summary>
        /// Checks if String1 equals String2 Ignoring the Case Sensitivity
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool equalsIgnoreCase(this string str1, string str2)
        {
            return EStrings.equalsIgnoreCase(str1, str2);
        }


        /// <summary>
        /// Converts to Proper Case
        /// </summary>
        /// <param name="pStr"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string toProperCase(this string pStr)
        {
            return Strings.StrConv(pStr, VbStrConv.ProperCase);
        }


        #endregion

    }
}