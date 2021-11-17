using System;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{

    /// <summary>
    /// Double is larger in range represented but with lower accuracy than decimal
    /// </summary>
    /// <remarks></remarks>
    public class EDouble
    {

        /// <summary>
        /// Converts An Object to Double
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static double valueOf(string obj)
        {
            try
            {
                if (obj is null || (obj.Trim() ?? "") == (string.Empty ?? ""))
                    return 0d;
                return Conversions.ToDouble(obj);
            }
            catch (Exception ex)
            {
                return 0.0d;
            }
        }


        /// <summary>
        /// Converts objects to double
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static double valueOf(object obj)
        {
            if (obj is null)
                return 0d;
            if (obj is string)
                return valueOf(Conversions.ToString(obj));
            if (obj is DBNull)
                return 0.0d;
            if (obj is double)
                return Conversions.ToDouble(obj);
            if (obj is long)
                return Conversions.ToLong(obj);
            if (obj is int)
                return Conversions.ToInteger(obj);
            if (obj is short)
                return Conversions.ToInteger(obj);
            if (obj is bool)
                return EShort.valueOf(Conversions.ToBoolean(obj));
            if (obj is decimal)
                return Conversions.ToDouble(obj);
            return 0d;    // REM Cant convert this
        }
    }
}