
Namespace Types



    ''' <summary>
    ''' Contains Different Delegate Definitions. Delegates is like const .. They dont need shared specifier. Once public
    ''' They are available
    ''' </summary>
    ''' <remarks></remarks>
    Public Class EDelegates

#Region "Delegates Sub"

#End Region

#Region "Delegates Functions"

#End Region

        ''' <summary>
        ''' Takes a string Parameter and returns boolean
        ''' </summary>
        ''' <param name="strParam"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Delegate Function delegateBoolFunc(ByVal strParam As String) As Boolean



        ''' <summary>
        ''' Takes a string Parameter and returns Nothing
        ''' </summary>
        ''' <param name="strParam"></param>
        ''' <remarks></remarks>
        Public Delegate Sub delegateSubString(ByVal strParam As String)


        ''' <summary>
        ''' Takes a boolean Parameter and returns Nothing
        ''' </summary>
        ''' <param name="strParam"></param>
        ''' <remarks></remarks>
        Public Delegate Sub delegateSubBool(ByVal strParam As Boolean)

        ''' <summary>
        ''' Takes a boolean Parameter,Thread Parameter and returns Nothing
        ''' </summary>
        ''' <param name="strParam"></param>
        ''' <remarks></remarks>
        Public Delegate Sub delegateSubBoolThread(ByVal strParam As Boolean, ByVal strParam As Threading.Thread)


        ''' <summary>
        ''' Takes a boolean Parameter and returns Nothing
        ''' </summary>
        ''' <param name="strParam"></param>
        ''' <remarks></remarks>
        Public Delegate Sub delegateSubThread(ByVal strParam As Threading.Thread)



        ''' <summary>
        ''' Collects No Parameters
        ''' </summary>
        ''' <remarks></remarks>
        Public Delegate Sub delegateNoParam()


    End Class


End Namespace