using System;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    public sealed class EShort
    {

        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(string obj)
        {
            try
            {
                if (Information.IsNothing(obj))
                    return 0;
                if (string.IsNullOrEmpty(obj))
                    return 0;
                if (string.IsNullOrEmpty(obj.Trim()))
                    return 0;
                return Conversions.ToShort(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        /// <summary>
        /// Converts An Object to short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(DBNull obj)
        {
            return 0;
        }

        /// <summary>
        /// Converts An Object to Short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(double obj)
        {
            try
            {
                return (short)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        /// <summary>
        /// Converts An Object to Short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(decimal obj)
        {
            try
            {
                return (short)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        /// <summary>
        /// Converts An Object to Short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(int obj)
        {
            try
            {
                return (short)obj;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        /// <summary>
        /// Converts An Object to Short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(long obj)
        {
            try
            {
                return (short)obj;
            }
            catch (Exception ex)
            {
                return 0;
            }
        }



        /// <summary>
        /// Converts boolean to short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(bool obj)
        {
            try
            {
                return (short)Math.Round(Math.Abs(Conversion.Val(obj)));
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        /// <summary>
        /// Converts objects to short
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static short valueOf(object obj)
        {
            if (obj is null)
                return 0;
            if (obj is string)
                return valueOf(Conversions.ToString(obj));
            if (obj is DBNull)
                return valueOf((DBNull)obj);
            if (obj is double)
                return valueOf(Conversions.ToDouble(obj));
            if (obj is long)
                return valueOf(Conversions.ToLong(obj));
            if (obj is int)
                return valueOf(Conversions.ToInteger(obj));
            if (obj is short)
                return Conversions.ToShort(obj);
            if (obj is bool)
                return valueOf(Conversions.ToBoolean(obj));
            return 0;    // REM Cant convert this
        }
    }
}