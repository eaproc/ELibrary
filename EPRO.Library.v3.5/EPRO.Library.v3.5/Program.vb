
Imports CODERiT.Logger.v._3._5




Friend Class Program


    Shared Sub New()
        ____Logger = New Log1(GetType(Program)).Load(Log1.Modes.FILE, True)

    End Sub




    Private Shared ____Logger As Log1
    Public Shared ReadOnly Property Logger As Log1
        Get
            Return ____Logger
        End Get
    End Property





End Class
