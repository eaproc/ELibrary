Imports CODERiT.Logger.v._3._5

Friend Class Program

    Shared Sub New()

        _LogFile = New Log1(GetType(Program)).Load(Log1.Modes.FILE, True)

    End Sub

    Private Shared _LogFile As Log1


    Public Shared ReadOnly Property Logger As Log1
        Get
            Return _LogFile
        End Get
    End Property




End Class
