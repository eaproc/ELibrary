Namespace Objects

    Public Class LikeIgnoreCaseComparer
        Implements IEqualityComparer(Of String)


        Public Function Equals1(x As String, y As String) As Boolean Implements IEqualityComparer(Of String).Equals
            If x Is Nothing Then Return False
            Return x.ToLower().IndexOf(y.ToLower()) >= 0
        End Function

        Public Function GetHashCode1(obj As String) As Integer Implements IEqualityComparer(Of String).GetHashCode
            Return obj.GetHashCode()
        End Function




    End Class


End Namespace