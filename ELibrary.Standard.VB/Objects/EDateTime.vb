Imports ELibrary.Standard.VB.Types
Imports ELibrary.Standard.VB.Objects.EStrings
Imports ELibrary.Standard.VB.Modules

Namespace Objects

    ''' <summary>
    ''' Controls the Date Conversions [To Compare Dates, Use .Equals for better functionality]
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EDateTime





#Region "Date Conversions between strings and date"


        Public Enum SpecialDateTimeFormats
            ''' <summary>
            ''' dd/MMM/yyyy
            ''' </summary>
            ''' <remarks></remarks>
            STYLE1

            ''' <summary>
            ''' dd/MMM/yyyy hh:mm:ss tt
            ''' </summary>
            ''' <remarks></remarks>
            STYLE2

            ''' <summary>
            ''' dd/MM/yyyy
            ''' </summary>
            ''' <remarks></remarks>
            STYLE3

            ''' <summary>
            ''' dd/MM/yyyy hh:mm:ss tt
            ''' </summary>
            ''' <remarks></remarks>
            STYLE4

        End Enum




        Public Shared Function getShortMonthNameInEnglish(pDate As Date) As String
            Select Case pDate.Month
                Case 1
                    Return "JAN"

                Case 2
                    Return "FEB"

                Case 3
                    Return "MAR"

                Case 4
                    Return "APR"

                Case 5
                    Return "MAY"

                Case 6
                    Return "JUN"

                Case 7
                    Return "JUL"

                Case 8
                    Return "AUG"

                Case 9
                    Return "SEP"

                Case 10
                    Return "OCT"

                Case 11
                    Return "NOV"

                Case Else
                    Return "DEC"
            End Select
        End Function



        Public Structure DateFormats
            ''' <summary>
            ''' General Formats understood by all platform [12/JUL/2001]
            ''' </summary>
            ''' <remarks></remarks>
            Public Const DateFormat1 As String = "dd/MMM/yyyy"

            ''' <summary>
            ''' General Formats understood by all platform [12 JUL 2001]
            ''' </summary>
            ''' <remarks></remarks>
            Public Const DateFormat2 As String = "dd MMM yyyy"

            ''' <summary>
            ''' General Formats understood by all platform [12-JUL-2001]
            ''' </summary>
            ''' <remarks></remarks>
            Public Const DateFormat3 As String = "dd-MMM-yyyy"


            ''' <summary>
            ''' Complex Format ..Meant for display [ Monday, April 07, 2014 ]
            ''' </summary>
            ''' <remarks></remarks>
            Public Const DateFormat4 As String = "dddd, MMMM dd, yyyy"




            Public Enum DateFormatsEnum
                ''' <summary>
                ''' [12/JUL/2001]
                ''' </summary>
                ''' <remarks></remarks>
                DateFormat1
                ''' <summary>
                ''' [12 JUL 2001]
                ''' </summary>
                ''' <remarks></remarks>
                DateFormat2
                ''' <summary>
                ''' [12-JUL-2001]
                ''' </summary>
                ''' <remarks></remarks>
                DateFormat3

                ''' <summary>
                ''' [ Monday, April 07, 2014 ]
                ''' </summary>
                ''' <remarks></remarks>
                DateFormat4

            End Enum



            Public Shared Function valueOf(ByVal _DateFormatsEnum As DateFormatsEnum) As String

                Select Case _DateFormatsEnum
                    Case Is = DateFormatsEnum.DateFormat1
                        Return DateFormat1
                    Case Is = DateFormatsEnum.DateFormat2
                        Return DateFormat2
                    Case Is = DateFormatsEnum.DateFormat3
                        Return DateFormat3

                    Case Is = DateFormatsEnum.DateFormat4
                        Return DateFormat4

                End Select

                Return DateFormat1

            End Function


        End Structure

        ''' <summary>
        ''' Sample [12:00 AM]
        ''' </summary>
        ''' <remarks></remarks>
        Public Const TimeFormatUsedWithoutSeconds As String = "hh:mm tt"

        ''' <summary>
        ''' 12:00:00 AM
        ''' </summary>
        ''' <remarks></remarks>
        Public Const TimeFormatUsedWithSeconds As String = "hh:mm:ss tt"


        Public Const DateTimeFormat1 As String = "MM/dd/yyy hh:mm:ss tt"
        Public Const DateTimeFormat2 As String = "MMM/dd/yyy hh:mm:ss tt"





        '''' <summary>
        '''' Returns general date string if value is ok if it is nothing
        '''' </summary>
        '''' <param name="obj"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function valueOf(ByVal obj As Object) As String

        '    Try
        '        If IsDBNull(obj) Then
        '            Return valueOf(CType(obj, System.DBNull)).ToString()
        '        Else
        '            Return Microsoft.VisualBasic.Strings.Format(CDate(obj), DateFormats.DateFormat1)
        '        End If
        '    Catch ex As Exception
        '        Return Nothing
        '    End Try

        'End Function


        ''' <summary>
        ''' Returns general date string if value is ok if it is nothing
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function valueOf(ByVal obj As Object,
                                       ByVal FormatNeeded As DateFormats.DateFormatsEnum) As String

            Return valueOf(obj, DateFormats.valueOf(FormatNeeded))
        End Function

        Public Shared Function valueOf(ByVal obj As Object,
                                       ByVal FormatNeeded As String) As String
            Try
                If obj Is Nothing OrElse obj.GetType() Is GetType(DBNull) Then
                    Return Nothing
                Else
                    Return CDate(obj).ToString(FormatNeeded)
                End If
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        ''' <summary>
        ''' It returns NULL as string if obj is NULL
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <param name="FormatNeeded"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function valueOf(ByVal obj As NullableDateTime,
                                      ByVal FormatNeeded As String) As String

            Try
                If obj = Nothing Then Return NullableDateTime.NULL_RETURN


                Return obj.DateTimeValue.ToString(FormatNeeded)
            Catch ex As Exception
                Return Nothing
            End Try

        End Function

        Public Shared Function time__valueOf(ByVal obj As Object) As String
            Return valueOf(obj, TimeFormatUsedWithoutSeconds)
        End Function


        Public Shared Function valueOf(ByVal obj As System.DBNull) As Date
            Return Nothing
        End Function



        Public Shared Function valueOf(ByVal pDateObj As NullableDateTime,
                                       pReturnFormat As SpecialDateTimeFormats,
                                       Optional ______ReturnIfNULL____ As String = "") As String
            If pDateObj Is Nothing OrElse pDateObj.isNull Then Return ______ReturnIfNULL____


            Select Case pReturnFormat
                Case SpecialDateTimeFormats.STYLE1
                    Return String.Format("{0}/{1}/{2}", pDateObj.DateTimeValue.Day,
                                         getShortMonthNameInEnglish(pDateObj.DateTimeValue),
                                         pDateObj.DateTimeValue.Year)

                Case SpecialDateTimeFormats.STYLE2
                    Return String.Format("{0}/{1}/{2} {3}", pDateObj.DateTimeValue.Day,
                                        getShortMonthNameInEnglish(pDateObj.DateTimeValue),
                                        pDateObj.DateTimeValue.Year,
                                        pDateObj.DateTimeValue.ToString(TimeFormatUsedWithSeconds)
                                        )
            End Select


            Return ______ReturnIfNULL____
        End Function

#End Region




#Region "Comparing DateTime"


        ''' <summary>
        ''' Checks if both Time are equal without checking the Date. Just hr, min and secs
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EqualsTimeWithoutDate(pTime1 As Date, pTime2 As Date) As Boolean
            If basExtensions.IsNothing(pTime1) OrElse basExtensions.IsNothing(pTime2) Then Return False

            Return pTime1.Hour = pTime2.Hour AndAlso pTime1.Minute = pTime2.Minute AndAlso pTime1.Second = pTime2.Second

        End Function



        ''' <summary>
        ''' Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr, min and secs
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsTime1IsGreaterThan(pTime1 As Date, pTime2 As Date) As Boolean
            If basExtensions.IsNothing(pTime1) OrElse basExtensions.IsNothing(pTime2) Then Return False
            If pTime1.Hour > pTime2.Hour Then Return True
            If pTime1.Hour = pTime2.Hour Then
                If pTime1.Minute > pTime2.Minute Then Return True
                If pTime1.Minute = pTime2.Minute Then
                    Return pTime1.Second > pTime2.Second
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function



        ''' <summary>
        ''' Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr, min and secs
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsTime1IsGreaterThanOrEquals(pTime1 As Date, pTime2 As Date) As Boolean
            Return IsTime1IsGreaterThan(pTime1, pTime2) OrElse
                EqualsTimeWithoutDate(pTime1, pTime2)
        End Function












        ''' <summary>
        ''' Checks if time1 is greater than time2 without checking DATE ONLY TIME.  Just hr and min
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsTime1HrMinsIsGreaterThan(pTime1 As Date, pTime2 As Date) As Boolean
            If basExtensions.IsNothing(pTime1) OrElse basExtensions.IsNothing(pTime2) Then Return False
            If pTime1.Hour > pTime2.Hour Then Return True
            If pTime1.Hour = pTime2.Hour Then
                Return pTime1.Minute > pTime2.Minute
            Else
                Return False
            End If
        End Function



        ''' <summary>
        ''' Checks if time1 is greater than time2 without checking DATE ONLY TIME. Just hr and min 
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsTime1HrMinsIsGreaterThanOrEquals(pTime1 As Date, pTime2 As Date) As Boolean
            Return IsTime1HrMinsIsGreaterThan(pTime1, pTime2) OrElse
                EqualsTimeHrMinsWithoutDate(pTime1, pTime2)
        End Function

        ''' <summary>
        ''' Checks if both Time are equal without checking the Date. Just hr and min 
        ''' </summary>
        ''' <param name="pTime1"></param>
        ''' <param name="pTime2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EqualsTimeHrMinsWithoutDate(pTime1 As Date, pTime2 As Date) As Boolean
            If basExtensions.IsNothing(pTime1) OrElse basExtensions.IsNothing(pTime2) Then Return False

            Return pTime1.Hour = pTime2.Hour AndAlso pTime1.Minute = pTime2.Minute

        End Function







        ''' <summary>
        ''' Checks if both date are equal without checking the time. Just Day, Month and Year
        ''' </summary>
        ''' <param name="Date1"></param>
        ''' <param name="Date2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EqualsDateWithoutTime(Date1 As Date, Date2 As Date) As Boolean
            If basExtensions.IsNothing(Date1) OrElse basExtensions.IsNothing(Date2) Then Return False

            Return Date1.Day = Date2.Day AndAlso Date1.Month = Date2.Month AndAlso Date1.Year = Date2.Year

        End Function


        ''' <summary>
        ''' Checks if dat1 is greater than date2 without checking time ONLY Date. Just Day, Month and Year
        ''' </summary>
        ''' <param name="Date1"></param>
        ''' <param name="Date2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsDate1IsGreaterThan(Date1 As Date, Date2 As Date) As Boolean
            If basExtensions.IsNothing(Date1) OrElse basExtensions.IsNothing(Date2) Then Return False
            If Date1.Year > Date2.Year Then Return True
            If Date1.Year = Date2.Year Then
                If Date1.Month > Date2.Month Then Return True
                If Date1.Month = Date2.Month Then
                    Return Date1.Day > Date2.Day
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function


        ''' <summary>
        ''' Checks if dat1 is greater than date2 without checking time ONLY Date. Just Day, Month and Year
        ''' </summary>
        ''' <param name="Date1"></param>
        ''' <param name="Date2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsDate1IsGreaterThanOrEquals(Date1 As Date, Date2 As Date) As Boolean
            Return IsDate1IsGreaterThan(Date1, Date2) OrElse
                EqualsDateWithoutTime(Date1, Date2)
        End Function


#End Region









        ''' <summary>
        ''' Get the time difference in Milliseconds [1 sec = 1000 ms]
        ''' </summary>
        ''' <param name="LowerTime"></param>
        ''' <param name="HigherTime"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetTimeDifferenceInMilliseconds(ByVal LowerTime As Date, ByVal HigherTime As Date) As Long
            Dim ms As TimeSpan = HigherTime - LowerTime
            Return ms.TotalMilliseconds
        End Function


        ''' <summary>
        ''' Convert Secs to Hr:MM:Secs
        ''' </summary>
        ''' <param name="Secs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertSecsToHrsMinsSecs(ByVal Secs As Long) As String
            Dim SecsLeft As Long = Secs
            Dim HoursSec As Long = getHours(SecsLeft)
            SecsLeft -= getSecs(
                                CInt(HoursSec)
                                )

            Dim MinsSec As Long = getMins(SecsLeft)
            SecsLeft -= getSecs(
                                MinsSec
                                 )

            Return String.Format(
                            "{0}:{1}:{2}",
                           WrapUpAtLeast(CStr(HoursSec), 2),
                           WrapUp(CStr(MinsSec), 2),
                           WrapUp(CStr(SecsLeft), 2)
                            )

        End Function

        ''' <summary>
        ''' Convert Secs to Hr:MM
        ''' </summary>
        ''' <param name="Secs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function ConvertSecsToString(ByVal Secs As Long) As String
            Dim SecsLeft As Long = Secs

            ''Hours in the secs
            Dim HoursSec As Long = getHours(SecsLeft)
            SecsLeft -= getSecs(
                                CInt(HoursSec)
                                )

            ''Mins in the Secs
            Dim MinsSec As Long = getMins(SecsLeft)


            Return String.Format(
                            "{0}:{1}",
                            WrapUpAtLeast(CStr(HoursSec), 2),
                           WrapUp(CStr(MinsSec), 2)
                             )


        End Function














        Public Shared Function getTimeFROM___TimeComponentPartOnlyAsSeconds(pLowerTime As Date) As Long
            Return (pLowerTime.Hour * 60 * 60) + (pLowerTime.Minute * 60) + pLowerTime.Second
        End Function
        Public Shared Function getTimeDifferenceFROM___TimeComponentPartOnlyAsSeconds(pLowerTime As Date, ByVal pHigherTime As DateTime) As Long
            Return getTimeFROM___TimeComponentPartOnlyAsSeconds(pHigherTime) - getTimeFROM___TimeComponentPartOnlyAsSeconds(pLowerTime)
        End Function

        ''' <summary>
        ''' Get Hours From Secs
        ''' </summary>
        ''' <param name="Secs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getHours(ByVal Secs As Long) As Long
            Return CLng(CInt(
                        Secs / (60 * 60)
                        ))

        End Function

        '''' <summary>
        '''' Get Hours From Date [Same as Hour() function]
        '''' </summary>
        '''' <param name="LimitedTime">Limited to 24 hrs .. specifically 23hrs 59mins</param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function getHours(ByVal LimitedTime As Date) As Long

        '    Try


        '        Return CLng(Split(FormatDateTime(LimitedTime, vbShortTime), ":")(0))

        '    Catch ex As Exception

        '        Return 0

        '    End Try

        'End Function

        ''' <summary>
        ''' Get Mins From Secs
        ''' </summary>
        ''' <param name="Secs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getMins(ByVal Secs As Long) As Long
            Return CLng(CInt(
                       CDbl(Secs) / CDbl(60)
                       ))
        End Function

        '''' <summary>
        '''' Get Mins From Date [Same as Minute() function]
        '''' </summary>
        '''' <param name="LimitedTime">Limited to 24 hrs .. specifically 23hrs 59mins</param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function getMins(ByVal LimitedTime As Date) As Long
        '    Try


        '        Return CLng(Split(FormatDateTime(LimitedTime, vbShortTime), ":")(1))

        '    Catch ex As Exception

        '        Return 0

        '    End Try
        'End Function


        ''' <summary>
        ''' Get Secs From Hours
        ''' </summary>
        ''' <param name="Hrs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function getSecs(ByVal Hrs As Integer) As Long
            Return CInt(
                       Hrs * (60 * 60)
                       )
        End Function


        ''' <summary>
        ''' Get Secs From Mins
        ''' </summary>
        ''' <param name="Mins"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overloads Shared Function getSecs(ByVal Mins As Long) As Long
            Return CInt(
                       Mins * (60)
                       )

        End Function



        ''' <summary>
        ''' Joins An instance of date with time. 
        ''' </summary>
        ''' <param name="sDate"></param>
        ''' <param name="sTime"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function AddDayWithTime(ByVal sDate As Date, ByVal sTime As Date) As Date

            Return sDate.Date + sTime.TimeOfDay

        End Function

        '''' <summary>
        '''' Format day date and then join it with the formatted timeDate
        '''' </summary>
        '''' <param name="sDate"></param>
        '''' <param name="sTime"></param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function AddDayDateWithTimeDate(ByVal sDate As Date, ByVal sTime As Date) As Date
        '    Try
        '        Return New Date(
        '                            CDate(
        '                                        valueOf(sDate, DateFormats.DateFormatsEnum.DateFormat1)
        '                                        ).Ticks +
        '                                    CDate(
        '                                FormatDateTime(sTime, vbLongTime)
        '                                                            ).Ticks
        '                        )

        '    Catch ex As Exception
        '        Return Nothing
        '    End Try
        'End Function


        ''' <summary>
        ''' Format Date.Now and Returns a 12Hr Time Format ['"hh : mm: ss tt"]
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FormatLongTime(Optional ByVal WithSpacedCharacters As Boolean = True) As String
            REM You can add option of with spaced [Default] or without

            Dim dTime As Date = Date.Now

            If WithSpacedCharacters Then
                Return String.Format("{0} : {1} : {2} {3}",
                                     WrapUp(CStr(dTime.Hour Mod 12), 2),
                                     WrapUp(CStr(dTime.Minute), 2),
                                     WrapUp(CStr(dTime.Second), 2),
                                     getAMPM(dTime)
                                     )

            Else

                Return String.Format("{0}:{1}:{2} {3}",
                                     WrapUp(CStr(dTime.Hour Mod 12), 2),
                                     WrapUp(CStr(dTime.Minute), 2),
                                     WrapUp(CStr(dTime.Second), 2),
                                     getAMPM(dTime)
                                     )
            End If

        End Function

        ''' <summary>
        ''' Get if it is AM or PM
        ''' </summary>
        ''' <param name="dTime"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getAMPM(ByVal dTime As Date) As String

            If dTime.Hour >= 12 Then

                Return "PM"

            Else
                Return "AM"

            End If

        End Function








        '''' <summary>
        '''' Fecthes first day from a week. Like I want mon from the same week today falls in .. Using the Pointer day
        '''' </summary>
        '''' <param name="ptrDay">A day in the week</param>
        '''' <param name="WkStartingDay">The first day you want from the week</param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function Get_First_Day_In_A_Week_Of_a_PointerDay(ByVal ptrDay As Date,
        '                                                    ByVal WkStartingDay As FirstDayOfWeek) As Date
        '    Return DateAdd(DateInterval.Day, 1 - Weekday(ptrDay, WkStartingDay), ptrDay)


        'End Function


        '''' <summary>
        '''' Fecthes a day from a week. Like I want mon from the same week today falls in .. starting from a day
        '''' </summary>
        '''' <param name="ptrDay">A day in the week</param>
        '''' <param name="WkStartingDay">The day that connotes the starting of the week</param>
        '''' <returns></returns>
        '''' <remarks></remarks>
        'Public Shared Function Get_A_Specific_Day_In_A_Week_Using_a_PointerDay(ByVal ptrDay As Date,
        '                                                            ByVal TheDayYouWant As FirstDayOfWeek,
        '                                                    ByVal WkStartingDay As FirstDayOfWeek) As Date
        '    Return DateAdd(DateInterval.Day,
        '                   CInt(TheDayYouWant) - Weekday(ptrDay, WkStartingDay),
        '                   ptrDay)




        'End Function




        ''' <summary>
        ''' Fecthes first day from a Month.  Using the Pointer day
        ''' </summary>
        ''' <param name="ptrDay">A day in the month </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_First_Day_In_THE_MONTH_Of_a_PointerDay(ByVal ptrDay As Date) As Date
            Return New DateTime(ptrDay.Year, ptrDay.Month, 1)
        End Function


        ''' <summary>
        ''' Fecthes Last day from a Month.  Using the Pointer day
        ''' </summary>
        ''' <param name="ptrDay">A day in the month </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Get_Last_Day_In_THE_MONTH_Of_a_PointerDay(ByVal ptrDay As Date) As Date
            Return Get_First_Day_In_THE_MONTH_Of_a_PointerDay(ptrDay).AddMonths(1).AddDays(-1)
        End Function










        ''' <summary>
        ''' Checks if day is between both endpoints or equal to either of them
        ''' </summary>
        ''' <param name="pDay"></param>
        ''' <param name="pStartDay"></param>
        ''' <param name="pEndDay"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsDayBetweenIncludingBothEndPoints(ByVal pDay As Date, pStartDay As Date, ByVal pEndDay As Date) As Boolean
            Return (
                IsDate1IsGreaterThan(pDay, pStartDay) OrElse EqualsDateWithoutTime(pDay, pStartDay)
                                                ) AndAlso
                                (
                                    IsDate1IsGreaterThan(pEndDay, pDay) OrElse EqualsDateWithoutTime(pDay, pEndDay)
                                                )
        End Function

        ''' <summary>
        ''' Checks if day is between the specified start day and end day without including both end points
        ''' </summary>
        ''' <param name="pDay"></param>
        ''' <param name="pStartDay"></param>
        ''' <param name="pEndDay"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function IsDayBetweenWithoutBothEndPoints(ByVal pDay As Date, pStartDay As Date, ByVal pEndDay As Date) As Boolean
            Return (IsDate1IsGreaterThan(pDay, pStartDay)) AndAlso
                                (IsDate1IsGreaterThan(pEndDay, pDay)
                                                )
        End Function

    End Class


End Namespace