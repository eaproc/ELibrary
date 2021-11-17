Imports CODERiT.Logger.v._3._5.Exceptions

Namespace Caching.Exceptions
    Public Class NotCachedException
        Inherits EException

        Sub New(ByVal FormTypeName As String)
            MyBase.New(
                String.Format(
                    "Invalid Code Usage. The requested Form [ {0} ] is NOT cached!!!",
                    FormTypeName
                    )
                )
        End Sub

    End Class

End Namespace