using System;
using System.Drawing;

namespace ELibrary.Standard.Objects
{
    public class ELong
    {
        /// <summary>
        /// Converts a string to a long.
        /// </summary>
        /// <param name="obj">The string to convert.</param>
        /// <returns>The long value of the string or 0 if conversion fails.</returns>
        public static long valueOf(string obj)
        {
            if (string.IsNullOrWhiteSpace(obj))
                return 0L;

            try
            {
                return Convert.ToInt64(obj.Trim());
            }
            catch
            {
                return 0L;
            }
        }

        /// <summary>
        /// Converts a DBNull to a long (always 0).
        /// </summary>
        public static long valueOf(DBNull obj)
        {
            return 0L;
        }

        /// <summary>
        /// Converts a double to a long by rounding.
        /// </summary>
        public static long valueOf(double obj)
        {
            return (long)Math.Round(obj);
        }

        /// <summary>
        /// Converts a decimal to a long by rounding.
        /// </summary>
        public static long valueOf(decimal obj)
        {
            return (long)Math.Round(obj);
        }

        /// <summary>
        /// Converts a boolean to a long (true = 1, false = 0).
        /// </summary>
        public static long valueOf(bool obj)
        {
            return obj ? 1L : 0L;
        }

        /// <summary>
        /// Converts a Color to a long based on its ARGB value.
        /// </summary>
        public static long valueOf(Color obj)
        {
            return obj.ToArgb();
        }

        /// <summary>
        /// Converts a short to a long.
        /// </summary>
        public static long valueOf(short obj)
        {
            return obj;
        }

        /// <summary>
        /// Converts an int to a long.
        /// </summary>
        public static long valueOf(int obj)
        {
            return obj;
        }

        /// <summary>
        /// Returns the long value directly.
        /// </summary>
        public static long valueOf(long obj)
        {
            return obj;
        }

        /// <summary>
        /// Converts an object to a long by determining its type.
        /// </summary>
        public static long valueOf(object obj)
        {
            if (obj == null || obj is DBNull)
                return 0L;

            try
            {
                switch (obj)
                {
                    case string strValue:
                        return valueOf(strValue);

                    case double doubleValue:
                        return valueOf(doubleValue);

                    case decimal decimalValue:
                        return valueOf(decimalValue);

                    case int intValue:
                        return intValue;

                    case long longValue:
                        return longValue;

                    case short shortValue:
                        return shortValue;

                    case bool boolValue:
                        return valueOf(boolValue);

                    case Color colorValue:
                        return valueOf(colorValue);

                    default:
                        return 0L; // Unconvertible type
                }
            }
            catch
            {
                return 0L;
            }
        }
    }
}
