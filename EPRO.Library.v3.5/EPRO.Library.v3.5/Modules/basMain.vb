Imports CODERiT.Logger.v._3._5


Namespace Modules


    Friend Module basMain

        Dim _LogFile As Log1 = New Log1(GetType(basMain)).Load(Log1.Modes.FILE, True)

        Public ReadOnly Property MyLogFile() As Log1
            Get
                Return _LogFile
            End Get
        End Property

    End Module

End Namespace