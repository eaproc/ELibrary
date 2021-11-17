
Namespace SplashCoordinator

    Public Interface IProcessCoordinator
        ''' <summary>
        ''' Runs when process coordinator has completely processed all processes passed in as parameters
        ''' </summary>
        ''' <remarks></remarks>
        Sub onExitAction()

        ''' <summary>
        ''' It tries to invoke this if user terminates the class before completion
        ''' </summary>
        ''' <remarks></remarks>
        Sub onTerminateAction()


    End Interface

End Namespace
