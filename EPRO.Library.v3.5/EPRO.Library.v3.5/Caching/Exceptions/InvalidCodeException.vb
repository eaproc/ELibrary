Imports CODERiT.Logger.v._3._5.Exceptions

Namespace Caching.Exceptions
    Public Class InvalidCodeException
        Inherits EException

        Sub New(ByVal FormTypeName As String,
                ByVal ThreadAptName As String)
            MyBase.New(
                String.Format(
                    "Invalid Code Usage. Make sure the form [ {0} ] you are trying to load conforms to your Thread Apartment State [ {1} ].",
                    FormTypeName,
                    ThreadAptName
                    )
                )
        End Sub

    End Class

End Namespace