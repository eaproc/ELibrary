Imports EPRO.Library.v3._5.Modules

Namespace Objects


    Public Class IEqualityComparerIgnoreCase
        Implements IEqualityComparer(Of String)

#Region "IEquality Comparer - Ignore Case"
        Public Function Equals1(x As String, y As String) As Boolean Implements IEqualityComparer(Of String).Equals
            Return equalsIgnoreCase(x, y)
        End Function

        Public Function GetHashCode1(obj As String) As Integer Implements IEqualityComparer(Of String).GetHashCode
            Return obj.GetHashCode()
        End Function

#End Region

    End Class

End Namespace
