using System;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{



    /// <summary>
    /// Decimal is smaller in range but has more accuracy than double
    /// </summary>
    /// <remarks></remarks>
    public class EDecimal
    {

        /// <summary>
        /// Converts An Object to Double
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static decimal valueOf(string obj)
        {
            try
            {
                if (obj is null || (obj.Trim() ?? "") == (string.Empty ?? ""))
                    return 0m;
                return Conversions.ToDecimal(obj);
            }
            catch (Exception ex)
            {
                return 0m;
            }
        }


        /// <summary>
        /// Converts objects to double
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static decimal valueOf(object obj)
        {
            if (obj is null)
                return 0m;
            double d = (double)Conversions.ToDecimal(obj);
            if (obj is string)
                return valueOf(Conversions.ToString(obj));
            if (obj is DBNull)
                return 0m;
            if (obj is double && Conversions.ToDouble(obj) <= (double)decimal.MaxValue)
                return Conversions.ToDecimal(obj);
            if (obj is long)
                return Conversions.ToLong(obj);
            if (obj is int)
                return Conversions.ToInteger(obj);
            if (obj is short)
                return Conversions.ToInteger(obj);
            if (obj is bool)
                return EShort.valueOf(Conversions.ToBoolean(obj));
            if (obj is decimal)
                return Conversions.ToDecimal(obj);
            return 0m;    // REM Cant convert this
        }
    }
}