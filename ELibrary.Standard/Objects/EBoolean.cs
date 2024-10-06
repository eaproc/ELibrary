using System;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{
    public class EBoolean
    {

        /// <summary>
        /// Converts An Object to Boolean
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool valueOf(object obj)
        {
            try
            {
                if (obj is null)
                    return false;
                if (obj is bool)
                {
                    return Conversions.ToBoolean(obj);
                }
                else if (obj is DBNull)
                {
                    return false;
                }
                else if (obj is string)
                {
                    return valueOf(Conversions.ToString(obj));
                }
                else if (obj is short || obj is int || obj is long || obj is double || obj is float || obj is decimal)
                {
                    return valueOf(ELong.valueOf(obj));
                }
            }
            catch (Exception)
            {
            }

            return false;
        }


        /// <summary>
        /// Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FALSE" 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool valueOf(string obj)
        {
            try
            {
                if (obj is object && (obj ?? "") != (string.Empty ?? ""))
                {
                    if (EStrings.equalsIgnoreCase(obj, "yes") || EStrings.equalsIgnoreCase(obj, "y"))
                        return true;
                    if (EStrings.equalsIgnoreCase(obj, "no") || EStrings.equalsIgnoreCase(obj, "n"))
                        return false;
                    if (EStrings.equalsIgnoreCase(obj, "true") || EStrings.equalsIgnoreCase(obj, "t"))
                        return true;
                    if (EStrings.equalsIgnoreCase(obj, "false") || EStrings.equalsIgnoreCase(obj, "f"))
                        return false;
                }
                // REM Check if it is number else try returning it like True written in string
                short TryInt;
                if (short.TryParse(obj, out TryInt))
                    return valueOf(TryInt);
                return Conversions.ToBoolean(obj);
            }
            catch (Exception ex)
            {
                return false;
            }
        }


        /// <summary>
        /// Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FALSE" 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool valueOf(short obj)
        {
            try
            {
                // REM Check if it is number else try returning it like True written in string

                return Conversions.ToBoolean(Math.Abs(obj));
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        /// <summary>
        /// Converts Both Integer String to Boolean. Like "123" AND "TRUE" OR "FASLSE" 
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool valueOf(long obj)
        {
            try
            {
                // REM Check if it is number else try returning it like True written in string

                return Conversions.ToBoolean(Math.Abs(obj));
            }
            catch (Exception ex)
            {
                return false;
            }
        }



        /// <summary>
        /// Converts A  Boolean to Boolean
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool valueOf(bool obj)
        {
            try
            {
                return obj;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
    }
}