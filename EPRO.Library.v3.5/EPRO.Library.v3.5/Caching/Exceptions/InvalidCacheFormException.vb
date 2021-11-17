Imports CODERiT.Logger.v._3._5.Exceptions

Namespace Caching.Exceptions
    Public Class InvalidCacheFormException
        Inherits EException

        Sub New()
            MyBase.New("Invalid Cache Form Type")
        End Sub

    End Class

End Namespace