using System;
using ELibrary.Standard.Objects;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

// REM Using Operators
// REM On Objects Assignment = Automatically takes the value if same object Type[ int=int]
// REM Others must be allowed using Ctype

// REM Operator = is the Comparison Type

namespace ELibrary.Standard.Types
{



    /// <summary>
    /// A type that can hold both date and Nothing[Null] produces Date or NULL String
    /// </summary>
    /// <remarks></remarks>
    public class NullableDateTime
    {

        #region Constructors


        public NullableDateTime(DateTime DateTimeVal)
        {
            setValue(DateTimeVal);
        }

        public NullableDateTime(string DateTimeVal)
        {
            setValue(DateTimeVal);
        }

        public NullableDateTime(DBNull DateTimeVal)
        {
            __isNull = true;
        }

        public NullableDateTime(object DateTimeVal)
        {
            try
            {
                if (Information.IsNothing(DateTimeVal))
                    break;
                if (DateTimeVal is DBNull)
                    break;
                if (DateTimeVal is string)
                {
                    setValue(Conversions.ToString(DateTimeVal));
                    return;
                }

                if (DateTimeVal is DateTime)
                {
                    setValue(Conversions.ToDate(DateTimeVal));
                    return;
                }

                if (DateTimeVal is NullableDateTime)
                {
                    setValue((NullableDateTime)DateTimeVal);
                    return;
                }
            }
            catch (Exception ex)
            {
            }

            __isNull = true;
        }



        #endregion


        #region Properties


        public static readonly NullableDateTime NULL_TIME = new NullableDateTime(Constants.vbNullString);
        public static readonly NullableDateTime NOW_DATETIME = new NullableDateTime(DateTime.Now);

        /// <summary>
        /// the value it returns instead of nothing
        /// </summary>
        /// <remarks></remarks>
        public const string NULL_RETURN = "NULL";
        private DateTime ____DateTimeVal;

        public string DateValue
        {
            get
            {
                if (!isNull)
                {
                    return EDateTime.valueOf(____DateTimeVal, EDateTime.DateFormats.DateFormatsEnum.DateFormat1);
                }

                return NULL_RETURN;
            }
        }

        public DateTime DateTimeValue
        {
            get
            {
                if (!isNull)
                {
                    return ____DateTimeVal;
                }

                return default;
            }
        }

        public string TimeValue
        {
            get
            {
                if (!isNull)
                {
                    return EDateTime.time__valueOf(____DateTimeVal);
                }

                return NULL_RETURN;
            }
        }

        private bool __isNull = true;

        public bool isNull
        {
            get
            {
                return __isNull;
            }
        }

        /// <summary>
        /// returns DateTimeValue if it is not null else returns NULL string 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public object DateTimeValueOrNULL
        {
            get
            {
                if (!isNull)
                    return DateTimeValue;
                return NULL_RETURN;
            }
        }

        /// <summary>
        /// returns DateTimeValue if it is not null else returns NULL object (NOTHING) 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public object DateTimeValueOrNothing
        {
            get
            {
                if (!isNull)
                    return DateTimeValue;
                return null;
            }
        }


        #endregion


        #region Operators

        public static bool operator ==(NullableDateTime ___NullableDateTime, object compareValue)
        {
            var cmpVal = new NullableDateTime(compareValue);
            if (___NullableDateTime.isNull & cmpVal.isNull)
                return true;
            if (___NullableDateTime.isNull | cmpVal.isNull)
                return false;
            return ___NullableDateTime.DateTimeValue.Equals(cmpVal.DateTimeValue);
        }

        public static bool operator !=(NullableDateTime ___NullableDateTime, object compareValue)
        {
            return !Operators.ConditionalCompareObjectEqual(___NullableDateTime, compareValue, false);
        }




        #endregion



        #region Methods


        public void setValue(string DateTimeVal)
        {
            try
            {
                if (Information.IsNothing(DateTimeVal))
                    break;
                if (string.IsNullOrEmpty(DateTimeVal))
                    break;
                ____DateTimeVal = Conversions.ToDate(DateTimeVal);
                __isNull = false;
                return;
            }
            catch (Exception ex)
            {
            }

            __isNull = true;
        }

        public void setValue(DateTime DateTimeVal)
        {
            try
            {
                if (DateTimeVal.Equals(default))
                    break;
                ____DateTimeVal = DateTimeVal;
                __isNull = false;
                return;
            }
            catch (Exception ex)
            {
            }

            __isNull = true;
        }

        private void setValue(NullableDateTime __nullableDateTime)
        {
            if (__nullableDateTime is null)
            {
                __isNull = true;
                return;
            }

            ____DateTimeVal = __nullableDateTime.DateTimeValue;
            __isNull = __nullableDateTime.isNull;
        }

        public override bool Equals(object obj)
        {
            return Operators.ConditionalCompareObjectEqual(this, obj, false);
        }

        public override string ToString()
        {
            return DateTimeValue.ToString();
        }

        #endregion



    }
}