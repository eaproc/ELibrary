using System;

namespace ELibrary.Standard.Objects
{
    public class EBoolean
    {
        /// <summary>
        /// Converts an object to a boolean.
        /// </summary>
        /// <param name="obj">The object to convert.</param>
        /// <returns>True if conversion is successful and logical; otherwise, false.</returns>
        public static bool valueOf(object obj)
        {
            if (obj == null || obj is DBNull)
                return false;

            try
            {
                switch (obj)
                {
                    case bool boolValue:
                        return boolValue;

                    case string stringValue:
                        return valueOf(stringValue);

                    case short shortValue:
                        return valueOf(shortValue);

                    case int intValue:
                        return valueOf((long)intValue);

                    case long longValue:
                        return valueOf(longValue);

                    case double doubleValue:
                    case float floatValue:
                    case decimal decimalValue:
                        // Cast the numeric types to long for a common boolean conversion
                        return valueOf(Convert.ToInt64(obj));

                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Converts a string to a boolean. Handles "true"/"false", "yes"/"no", and numeric strings.
        /// </summary>
        /// <param name="obj">The string to convert.</param>
        /// <returns>True if conversion is successful and logical; otherwise, false.</returns>
        public static bool valueOf(string obj)
        {
            if (string.IsNullOrWhiteSpace(obj))
                return false;

            try
            {
                string lowerObj = obj.Trim().ToLower();

                // Handle common boolean representations
                if (lowerObj == "yes" || lowerObj == "y" || lowerObj == "true" || lowerObj == "t")
                    return true;

                if (lowerObj == "no" || lowerObj == "n" || lowerObj == "false" || lowerObj == "f")
                    return false;

                // Attempt to parse as a numeric value
                if (short.TryParse(lowerObj, out short numericValue))
                    return valueOf(numericValue);

                // Fall back to a general boolean conversion if applicable
                return bool.TryParse(lowerObj, out bool boolValue) && boolValue;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Converts a short to a boolean. A value of 0 is false, any other value is true.
        /// </summary>
        /// <param name="obj">The short to convert.</param>
        /// <returns>True if obj is non-zero; otherwise, false.</returns>
        public static bool valueOf(short obj)
        {
            return obj != 0;
        }

        /// <summary>
        /// Converts a long to a boolean. A value of 0 is false, any other value is true.
        /// </summary>
        /// <param name="obj">The long to convert.</param>
        /// <returns>True if obj is non-zero; otherwise, false.</returns>
        public static bool valueOf(long obj)
        {
            return obj != 0;
        }

        /// <summary>
        /// Returns the boolean value directly.
        /// </summary>
        /// <param name="obj">The boolean to return.</param>
        /// <returns>The same boolean value passed in.</returns>
        public static bool valueOf(bool obj)
        {
            return obj;
        }
    }
}
