using System;
using static ELibrary.Standard.Objects.EStrings;
using ELibrary.Standard.Types;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.Objects
{

    /// <summary>
    /// Controls the Date Conversions [To Compare Dates, Use .Equals for better functionality]
    /// </summary>
    /// <remarks></remarks>
    public class EDateTime
    {





        #region Date Conversions between strings and date


        public enum SpecialDateTimeFormats
        {
            /// <summary>
            /// dd/MMM/yyyy
            /// </summary>
            /// <remarks></remarks>
            STYLE1,

            /// <summary>
            /// dd/MMM/yyyy hh:mm:ss tt
            /// </summary>
            /// <remarks></remarks>
            STYLE2,

            /// <summary>
            /// dd/MM/yyyy
            /// </summary>
            /// <remarks></remarks>
            STYLE3,

            /// <summary>
            /// dd/MM/yyyy hh:mm:ss tt
            /// </summary>
            /// <remarks></remarks>
            STYLE4
        }

        public static string getShortMonthNameInEnglish(DateTime pDate)
        {
            switch (pDate.Month)
            {
                case 1:
                    {
                        return "JAN";
                    }

                case 2:
                    {
                        return "FEB";
                    }

                case 3:
                    {
                        return "MAR";
                    }

                case 4:
                    {
                        return "APR";
                    }

                case 5:
                    {
                        return "MAY";
                    }

                case 6:
                    {
                        return "JUN";
                    }

                case 7:
                    {
                        return "JUL";
                    }

                case 8:
                    {
                        return "AUG";
                    }

                case 9:
                    {
                        return "SEP";
                    }

                case 10:
                    {
                        return "OCT";
                    }

                case 11:
                    {
                        return "NOV";
                    }

                default:
                    {
                        return "DEC";
                    }
            }
        }

        public struct DateFormats
        {
            /// <summary>
            /// General Formats understood by all platform [12/JUL/2001]
            /// </summary>
            /// <remarks></remarks>
            public const string DateFormat1 = "dd/MMM/yyyy";

            /// <summary>
            /// General Formats understood by all platform [12 JUL 2001]
            /// </summary>
            /// <remarks></remarks>
            public const string DateFormat2 = "dd MMM yyyy";

            /// <summary>
            /// General Formats understood by all platform [12-JUL-2001]
            /// </summary>
            /// <remarks></remarks>
            public const string DateFormat3 = "dd-MMM-yyyy";


            /// <summary>
            /// Complex Format ..Meant for display [ Monday, April 07, 2014 ]
            /// </summary>
            /// <remarks></remarks>
            public const string DateFormat4 = "dddd, MMMM dd, yyyy";

            public enum DateFormatsEnum
            {
                /// <summary>
                /// [12/JUL/2001]
                /// </summary>
                /// <remarks></remarks>
                DateFormat1,
                /// <summary>
                /// [12 JUL 2001]
                /// </summary>
                /// <remarks></remarks>
                DateFormat2,
                /// <summary>
                /// [12-JUL-2001]
                /// </summary>
                /// <remarks></remarks>
                DateFormat3,

                /// <summary>
                /// [ Monday, April 07, 2014 ]
                /// </summary>
                /// <remarks></remarks>
                DateFormat4
            }

            public static string valueOf(DateFormatsEnum _DateFormatsEnum)
            {
                switch (_DateFormatsEnum)
                {
                    case var @case when @case == DateFormatsEnum.DateFormat1:
                        {
                            return DateFormat1;
                        }

                    case var case1 when case1 == DateFormatsEnum.DateFormat2:
                        {
                            return DateFormat2;
                        }

                    case var case2 when case2 == DateFormatsEnum.DateFormat3:
                        {
                            return DateFormat3;
                        }

                    case var case3 when case3 == DateFormatsEnum.DateFormat4:
                        {
                            return DateFormat4;
                        }
                }

                return DateFormat1;
            }
        }

        /// <summary>
        /// Sample [12:00 AM]
        /// </summary>
        /// <remarks></remarks>
        public const string TimeFormatUsedWithoutSeconds = "hh:mm tt";

        /// <summary>
        /// 12:00:00 AM
        /// </summary>
        /// <remarks></remarks>
        public const string TimeFormatUsedWithSeconds = "hh:mm:ss tt";
        public const string DateTimeFormat1 = "MM/dd/yyy hh:mm:ss tt";
        public const string DateTimeFormat2 = "MMM/dd/yyy hh:mm:ss tt";





        /// <summary>
        /// Returns general date string if value is ok if it is nothing
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string valueOf(object obj)
        {
            try
            {
                if (Information.IsDBNull(obj))
                {
                    return valueOf((DBNull)obj).ToString();
                }
                else
                {
                    return Strings.Format(Conversions.ToDate(obj), DateFormats.DateFormat1);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        /// <summary>
        /// Returns general date string if value is ok if it is nothing
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string valueOf(object obj, DateFormats.DateFormatsEnum FormatNeeded)
        {
            try
            {
                if (Information.IsDBNull(obj))
                {
                    return null;
                }
                else
                {
                    return Strings.Format(Conversions.ToDate(obj), DateFormats.valueOf(FormatNeeded));
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static string valueOf(object obj, string FormatNeeded)
        {
            try
            {
                if (Information.IsDBNull(obj))
                {
                    return null;
                }
                else
                {
                    return Strings.Format(Conversions.ToDate(obj), FormatNeeded);
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// It returns NULL as string if obj is NULL
        /// </summary>
        /// <param name="obj"></param>
        /// <param name="FormatNeeded"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string valueOf(NullableDateTime obj, string FormatNeeded)
        {
            try
            {
                if (obj == null)
                    return NullableDateTime.NULL_RETURN;
                return Strings.Format(obj.DateTimeValue, FormatNeeded);
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public static string time__valueOf(object obj)
        {
            return valueOf(obj, TimeFormatUsedWithoutSeconds);
        }

        public static DateTime valueOf(DBNull obj)
        {
            return default;
        }

        public static string valueOf(NullableDateTime pDateObj, SpecialDateTimeFormats pReturnFormat, string ______ReturnIfNULL____ = "")
        {
            if (pDateObj is null || pDateObj.isNull)
                return ______ReturnIfNULL____;
            switch (pReturnFormat)
            {
                case SpecialDateTimeFormats.STYLE1:
                    {
                        return string.Format("{0}/{1}/{2}", pDateObj.DateTimeValue.Day, getShortMonthNameInEnglish(pDateObj.DateTimeValue), pDateObj.DateTimeValue.Year);
                    }

                case SpecialDateTimeFormats.STYLE2:
                    {
                        return string.Format("{0}/{1}/{2} {3}", pDateObj.DateTimeValue.Day, getShortMonthNameInEnglish(pDateObj.DateTimeValue), pDateObj.DateTimeValue.Year, pDateObj.DateTimeValue.ToString(TimeFormatUsedWithSeconds));
                    }
            }

            return ______ReturnIfNULL____;
        }

        #endregion




        #region Comparing DateTime


        /// <summary>
        /// Checks if both Time are equal without checking the Date. Just hr, min and secs
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool EqualsTimeWithoutDate(DateTime pTime1, DateTime pTime2)
        {
            if (Information.IsNothing(pTime1) || Information.IsNothing(pTime2))
                return false;
            return pTime1.Hour == pTime2.Hour && pTime1.Minute == pTime2.Minute && pTime1.Second == pTime2.Second;
        }



        /// <summary>
        /// Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr, min and secs
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsTime1IsGreaterThan(DateTime pTime1, DateTime pTime2)
        {
            if (Information.IsNothing(pTime1) || Information.IsNothing(pTime2))
                return false;
            if (pTime1.Hour > pTime2.Hour)
                return true;
            if (pTime1.Hour == pTime2.Hour)
            {
                if (pTime1.Minute > pTime2.Minute)
                    return true;
                if (pTime1.Minute == pTime2.Minute)
                {
                    return pTime1.Second > pTime2.Second;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }



        /// <summary>
        /// Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr, min and secs
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsTime1IsGreaterThanOrEquals(DateTime pTime1, DateTime pTime2)
        {
            return IsTime1IsGreaterThan(pTime1, pTime2) || EqualsTimeWithoutDate(pTime1, pTime2);
        }












        /// <summary>
        /// Checks if time1 is greater than time2 without checking DATE ONLY TIME.  Just hr and min
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsTime1HrMinsIsGreaterThan(DateTime pTime1, DateTime pTime2)
        {
            if (Information.IsNothing(pTime1) || Information.IsNothing(pTime2))
                return false;
            if (pTime1.Hour > pTime2.Hour)
                return true;
            if (pTime1.Hour == pTime2.Hour)
            {
                return pTime1.Minute > pTime2.Minute;
            }
            else
            {
                return false;
            }
        }



        /// <summary>
        /// Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr and min 
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsTime1HrMinsIsGreaterThanOrEquals(DateTime pTime1, DateTime pTime2)
        {
            return IsTime1HrMinsIsGreaterThan(pTime1, pTime2) || EqualsTimeHrMinsWithoutDate(pTime1, pTime2);
        }

        /// <summary>
        /// Checks if both Time are equal without checking the Date. Just hr and min 
        /// </summary>
        /// <param name="pTime1"></param>
        /// <param name="pTime2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool EqualsTimeHrMinsWithoutDate(DateTime pTime1, DateTime pTime2)
        {
            if (Information.IsNothing(pTime1) || Information.IsNothing(pTime2))
                return false;
            return pTime1.Hour == pTime2.Hour && pTime1.Minute == pTime2.Minute;
        }







        /// <summary>
        /// Checks if both date are equal without checking the time. Just Day, Month and Year
        /// </summary>
        /// <param name="Date1"></param>
        /// <param name="Date2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool EqualsDateWithoutTime(DateTime Date1, DateTime Date2)
        {
            if (Information.IsNothing(Date1) || Information.IsNothing(Date2))
                return false;
            return Date1.Day == Date2.Day && Date1.Month == Date2.Month && Date1.Year == Date2.Year;
        }


        /// <summary>
        /// Checks if dat1 is greater than date2 without checking time ONLY Date. Just Day, Month and Year
        /// </summary>
        /// <param name="Date1"></param>
        /// <param name="Date2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsDate1IsGreaterThan(DateTime Date1, DateTime Date2)
        {
            if (Information.IsNothing(Date1) || Information.IsNothing(Date2))
                return false;
            if (Date1.Year > Date2.Year)
                return true;
            if (Date1.Year == Date2.Year)
            {
                if (Date1.Month > Date2.Month)
                    return true;
                if (Date1.Month == Date2.Month)
                {
                    return Date1.Day > Date2.Day;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }


        /// <summary>
        /// Checks if dat1 is greater than date2 without checking time ONLY Date. Just Day, Month and Year
        /// </summary>
        /// <param name="Date1"></param>
        /// <param name="Date2"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsDate1IsGreaterThanOrEquals(DateTime Date1, DateTime Date2)
        {
            return IsDate1IsGreaterThan(Date1, Date2) || EqualsDateWithoutTime(Date1, Date2);
        }


        #endregion









        /// <summary>
        /// Get the time difference in Milliseconds [1 sec = 1000 ms]
        /// </summary>
        /// <param name="LowerTime"></param>
        /// <param name="HigherTime"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long GetTimeDifferenceInMilliseconds(DateTime LowerTime, DateTime HigherTime)
        {
            long ms;
            if (HigherTime.Millisecond < LowerTime.Millisecond)
            {
                ms = HigherTime.Millisecond + 1000 - LowerTime.Millisecond;
            }
            else
            {
                ms = HigherTime.Millisecond - LowerTime.Millisecond;
            }

            return DateAndTime.DateDiff(DateInterval.Second, LowerTime, HigherTime) * 1000L + ms;
        }


        /// <summary>
        /// Convert Secs to Hr:MM:Secs
        /// </summary>
        /// <param name="Secs"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string ConvertSecsToHrsMinsSecs(long Secs)
        {
            long SecsLeft = Secs;
            long HoursSec = getHours(SecsLeft);
            SecsLeft -= getSecs((int)HoursSec);
            long MinsSec = getMins(SecsLeft);
            SecsLeft -= getSecs(MinsSec);
            return string.Format("{0}:{1}:{2}", WrapUpAtLeast(HoursSec.ToString(), 2), WrapUp(MinsSec.ToString(), 2), WrapUp(SecsLeft.ToString(), 2));
        }

        /// <summary>
        /// Convert Secs to Hr:MM
        /// </summary>
        /// <param name="Secs"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string ConvertSecsToString(long Secs)
        {
            long SecsLeft = Secs;

            // 'Hours in the secs
            long HoursSec = getHours(SecsLeft);
            SecsLeft -= getSecs((int)HoursSec);

            // 'Mins in the Secs
            long MinsSec = getMins(SecsLeft);
            return string.Format("{0}:{1}", WrapUpAtLeast(HoursSec.ToString(), 2), WrapUp(MinsSec.ToString(), 2));
        }

        public static long getTimeFROM___TimeComponentPartOnlyAsSeconds(DateTime pLowerTime)
        {
            return pLowerTime.Hour * 60 * 60 + pLowerTime.Minute * 60 + pLowerTime.Second;
        }

        public static long getTimeDifferenceFROM___TimeComponentPartOnlyAsSeconds(DateTime pLowerTime, DateTime pHigherTime)
        {
            return getTimeFROM___TimeComponentPartOnlyAsSeconds(pHigherTime) - getTimeFROM___TimeComponentPartOnlyAsSeconds(pLowerTime);
        }

        /// <summary>
        /// Get Hours From Secs
        /// </summary>
        /// <param name="Secs"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getHours(long Secs)
        {
            return (long)Math.Round(Conversion.Int(Secs / (double)(60 * 60)));
        }

        /// <summary>
        /// Get Hours From Date [Same as Hour() function]
        /// </summary>
        /// <param name="LimitedTime">Limited to 24 hrs .. specifically 23hrs 59mins</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getHours(DateTime LimitedTime)
        {
            try
            {
                return Conversions.ToLong(Strings.Split(Strings.FormatDateTime(LimitedTime, Constants.vbShortTime), ":")[0]);
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }

        /// <summary>
        /// Get Mins From Secs
        /// </summary>
        /// <param name="Secs"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getMins(long Secs)
        {
            return (long)Math.Round(Conversion.Int((double)Secs / 60));
        }

        /// <summary>
        /// Get Mins From Date [Same as Minute() function]
        /// </summary>
        /// <param name="LimitedTime">Limited to 24 hrs .. specifically 23hrs 59mins</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getMins(DateTime LimitedTime)
        {
            try
            {
                return Conversions.ToLong(Strings.Split(Strings.FormatDateTime(LimitedTime, Constants.vbShortTime), ":")[1]);
            }
            catch (Exception ex)
            {
                return 0L;
            }
        }


        /// <summary>
        /// Get Secs From Hours
        /// </summary>
        /// <param name="Hrs"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getSecs(int Hrs)
        {
            return Conversion.Int(Hrs * 60 * 60);
        }


        /// <summary>
        /// Get Secs From Mins
        /// </summary>
        /// <param name="Mins"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static long getSecs(long Mins)
        {
            return Conversion.Int(Mins * 60);
        }



        /// <summary>
        /// Joins An instance of date with time. 
        /// </summary>
        /// <param name="sDate"></param>
        /// <param name="sTime"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime AddDayWithTime(DateTime sDate, DateTime sTime)
        {
            return sDate.Date + sTime.TimeOfDay;
        }

        /// <summary>
        /// Format day date and then join it with the formatted timeDate
        /// </summary>
        /// <param name="sDate"></param>
        /// <param name="sTime"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime AddDayDateWithTimeDate(DateTime sDate, DateTime sTime)
        {
            try
            {
                return new DateTime(Conversions.ToDate(valueOf(sDate, DateFormats.DateFormatsEnum.DateFormat1)).Ticks + Conversions.ToDate(Strings.FormatDateTime(sTime, Constants.vbLongTime)).Ticks);
            }
            catch (Exception ex)
            {
                return default;
            }
        }


        /// <summary>
        /// Format Now and Returns a 12Hr Time Format ['"hh : mm: ss tt"]
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string FormatLongTime(bool WithSpacedCharacters = true)
        {
            // REM You can add option of with spaced [Default] or without

            var dTime = DateAndTime.Now;
            if (WithSpacedCharacters)
            {
                return string.Format("{0} : {1} : {2} {3}", WrapUp((dTime.Hour % 12).ToString(), 2), WrapUp(dTime.Minute.ToString(), 2), WrapUp(dTime.Second.ToString(), 2), getAMPM(dTime));
            }
            else
            {
                return string.Format("{0}:{1}:{2} {3}", WrapUp((dTime.Hour % 12).ToString(), 2), WrapUp(dTime.Minute.ToString(), 2), WrapUp(dTime.Second.ToString(), 2), getAMPM(dTime));
            }
        }

        /// <summary>
        /// Get if it is AM or PM
        /// </summary>
        /// <param name="dTime"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string getAMPM(DateTime dTime)
        {
            if (dTime.Hour >= 12)
            {
                return "PM";
            }
            else
            {
                return "AM";
            }
        }








        /// <summary>
        /// Fecthes first day from a week. Like I want mon from the same week today falls in .. Using the Pointer day
        /// </summary>
        /// <param name="ptrDay">A day in the week</param>
        /// <param name="WkStartingDay">The first day you want from the week</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime Get_First_Day_In_A_Week_Of_a_PointerDay(DateTime ptrDay, FirstDayOfWeek WkStartingDay)
        {
            return DateAndTime.DateAdd(DateInterval.Day, 1 - DateAndTime.Weekday(ptrDay, WkStartingDay), ptrDay);
        }


        /// <summary>
        /// Fecthes a day from a week. Like I want mon from the same week today falls in .. starting from a day
        /// </summary>
        /// <param name="ptrDay">A day in the week</param>
        /// <param name="WkStartingDay">The day that connotes the starting of the week</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime Get_A_Specific_Day_In_A_Week_Using_a_PointerDay(DateTime ptrDay, FirstDayOfWeek TheDayYouWant, FirstDayOfWeek WkStartingDay)
        {
            return DateAndTime.DateAdd(DateInterval.Day, (int)TheDayYouWant - DateAndTime.Weekday(ptrDay, WkStartingDay), ptrDay);
        }




        /// <summary>
        /// Fecthes first day from a Month.  Using the Pointer day
        /// </summary>
        /// <param name="ptrDay">A day in the month </param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime Get_First_Day_In_THE_MONTH_Of_a_PointerDay(DateTime ptrDay)
        {
            return new DateTime(ptrDay.Year, ptrDay.Month, 1);
        }


        /// <summary>
        /// Fecthes Last day from a Month.  Using the Pointer day
        /// </summary>
        /// <param name="ptrDay">A day in the month </param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static DateTime Get_Last_Day_In_THE_MONTH_Of_a_PointerDay(DateTime ptrDay)
        {
            return Get_First_Day_In_THE_MONTH_Of_a_PointerDay(ptrDay).AddMonths(1).AddDays(-1);
        }










        /// <summary>
        /// Checks if day is between both endpoints or equal to either of them
        /// </summary>
        /// <param name="pDay"></param>
        /// <param name="pStartDay"></param>
        /// <param name="pEndDay"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsDayBetweenIncludingBothEndPoints(DateTime pDay, DateTime pStartDay, DateTime pEndDay)
        {
            return (IsDate1IsGreaterThan(pDay, pStartDay) || EqualsDateWithoutTime(pDay, pStartDay)) && (IsDate1IsGreaterThan(pEndDay, pDay) || EqualsDateWithoutTime(pDay, pEndDay));
        }

        /// <summary>
        /// Checks if day is between the specified start day and end day without including both end points
        /// </summary>
        /// <param name="pDay"></param>
        /// <param name="pStartDay"></param>
        /// <param name="pEndDay"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static bool IsDayBetweenWithoutBothEndPoints(DateTime pDay, DateTime pStartDay, DateTime pEndDay)
        {
            return IsDate1IsGreaterThan(pDay, pStartDay) && IsDate1IsGreaterThan(pEndDay, pDay);
        }
    }
}