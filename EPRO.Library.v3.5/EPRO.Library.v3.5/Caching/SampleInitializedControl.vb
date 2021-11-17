
Namespace Caching
    Public Class SampleInitializedControl
        Inherits System.Windows.Forms.Control

        ''' <summary>
        ''' Just a simple control that force the creation of handle under this calling thread
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            MyBase.New()
            Me.CreateHandle()

        End Sub

    End Class
End Namespace