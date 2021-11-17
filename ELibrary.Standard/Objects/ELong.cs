using System;
using System.Drawing;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    public class ELong
    {

        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long valueOf(string obj)
        {
            try
            {
                if (string.IsNullOrEmpty(obj))
                    return 0L;
                if (string.IsNullOrEmpty(obj.Trim()))
                    return 0L;
                return Conversions.ToLong(obj);
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }


        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long valueOf(DBNull obj)
        {
            try
            {
                return 0L;
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }

        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long valueOf(double obj)
        {
            try
            {
                return (long)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }


        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long valueOf(decimal obj)
        {
            try
            {
                return (long)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }

        public static long valueOf(bool obj)
        {
            try
            {
                return Math.Abs(Conversions.ToLong(obj));
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }

        public static long valueOf(Color obj)
        {
            try
            {
                return obj.ToArgb();
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }

        public static long valueOf(short obj)
        {
            return obj;
        }

        public static long valueOf(int obj)
        {
            return obj;
        }

        public static long valueOf(long obj)
        {
            return obj;
        }

        public static long valueOf(object obj)
        {
            if (obj is null)
                return 0L;
            if (obj is string)
                return valueOf(Conversions.ToString(obj));
            if (obj is DBNull)
                return 0L;
            if (obj is double)
                return valueOf(Conversions.ToDouble(obj));
            if (obj is long)
                return Conversions.ToLong(obj);
            if (obj is int)
                return valueOf(Conversions.ToInteger(obj));
            if (obj is decimal)
                return valueOf(Conversions.ToDecimal(obj));
            if (obj is short)
                return valueOf(Conversions.ToShort(obj));
            if (obj is bool)
                return valueOf(Conversions.ToBoolean(obj));
            if (obj is Color)
                return valueOf((Color)obj);
            return 0L;    // REM Cant convert this
        }
    }
}