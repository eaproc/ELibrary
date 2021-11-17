using System;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    public sealed class EInt
    {

        /// <summary>
        /// Converts An Object to long
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int valueOf(string obj)
        {
            try
            {
                if (Information.IsNothing(obj))
                    return 0;
                if (string.IsNullOrEmpty(obj))
                    return 0;
                if (string.IsNullOrEmpty(obj.Trim()))
                    return 0;
                return Conversions.ToInteger(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }


        /// <summary>
        /// Converts An Object to int32
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int valueOf(DBNull obj)
        {
            return 0;
        }





        /// <summary>
        /// Converts An Object to int32
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int valueOf(double obj)
        {
            try
            {
                return (int)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        /// <summary>
        /// Converts An Object to int32
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int valueOf(decimal obj)
        {
            try
            {
                return (int)Math.Round(obj);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }




        /// <summary>
        /// Converts objects to integer
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int valueOf(object obj)
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
                return valueOf((decimal)Conversions.ToLong(obj));
            if (obj is int)
                return Conversions.ToInteger(obj);
            if (obj is short)
                return Conversions.ToInteger(obj);
            if (obj is decimal)
                return valueOf(Conversions.ToDecimal(obj));
            if (obj is bool)
                return EShort.valueOf(Conversions.ToBoolean(obj));
            return 0;    // REM Cant convert this
        }
    }
}