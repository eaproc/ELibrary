Imports ELibrary.Standard.VB.Modules

Public Class EMaths






    ''' <summary>
    ''' Pure decimal NOT string formatted. e.g. 9998,00 or 9998.00
    ''' </summary>
    ''' <param name="xTurkishEntry"></param>
    ''' <returns>String because if the OS is Turkish and I use inbuilt format it will still be in Turkish</returns>
    ''' <remarks></remarks>

    'Public Shared Function Convert_Turkish_Decimal_to_Eng(ByVal xTurkishEntry As String) As String

    '    If xTurkishEntry.IndexOf(",") >= 0 Then

    '        Dim SplittedCInt() As String = Split(xTurkishEntry, ",")
    '        Return String.Format("{0}.{1}", SplittedVal)

    '    End If

    '    Return xTurkishEntry

    'End Function

    ''' <summary>
    ''' Writes a Byte Value in Kb, Mb etc
    ''' </summary>
    ''' <param name="ByteValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getReadableByteValue(ByVal ByteValue As Long) As String
        If ByteValue > (1024 * 1024) Then
            Return FormatNumber(
             ByteValue / (1024 * 1024), 2
            ) & "Mb"
        Else
            Return FormatNumber(
             ByteValue / (1024), 2
            ) & "Kb"
        End If
    End Function


    ''' <summary>
    ''' Alternate Bit Value. If 0 returns 1.. If 1 returns 0
    ''' </summary>
    ''' <param name="BitVal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function AlternateBit(ByVal BitVal As Byte) As Byte
        'If it is greater than 0 it is true 
        'First neutralize the current sign
        Return CByte(Math.Abs(
                    CInt(
                    Not CBool(
                        Math.Abs(BitVal)
                            )
                        )
                    ))
    End Function


    ''' <summary>
    ''' Fetch the Integer Portion ONLY of a Value.. NB: Sign is included
    ''' </summary>
    ''' <param name="dblVal"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function getIntegerPortionOf(ByVal dblVal As String) As Integer

        Try


            REM Instr is based 1 function and returns 0 if not found
            If dblVal.IndexOf(".") > 0 Then

                Return CInt(
                            dblVal.Substring(0,
                                            dblVal.IndexOf(".")
                                            )
                                        )
            Else
                Return CInt(dblVal)

            End If

        Catch ex As Exception

            Return 0
        End Try

    End Function



    'Public Shared Function getCurrencyValue(ByVal strValue As String) As Double
    '    If strValue = vbNullString Then Return 0
    '    If IsNumeric(Left(strValue, 1)) Or Len(strValue) = 1 Then Return CInt(strValue)


    '    Return CInt(Right(strValue, Len(strValue) - 1))
    'End Function

    'Public Shared Function setCurrencyValue(ByVal dblCurrency As Double,
    '                                       Optional ByVal strCurrencyTpe As String = "$") As String
    '    Return strCurrencyTpe & FormatNumber(dblCurrency, 2)
    'End Function


    'Public Shared Function getPercentage(ByVal ActualValue As Double, ByVal Total As Double) As Double
    '    Try
    '        Return CDbl(FormatNumber((ActualValue / Total) * 100, 2))
    '    Catch ex As Exception
    '        Return 0
    '    End Try
    'End Function

    Public Shared Function divide(pUpperNumerator As Double, pDownDenominator As Double) As Double
        If pDownDenominator = 0 Then Return 0
        Return pUpperNumerator / pDownDenominator
    End Function

    Public Shared Function divide(pUpperNumerator As Decimal, pDownDenominator As Decimal) As Decimal
        If pDownDenominator = 0 Then Return 0
        Return pUpperNumerator / pDownDenominator
    End Function


End Class
