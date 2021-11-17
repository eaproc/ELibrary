
Namespace MicrosoftOS
    Public Class PCAccountNames



        Public Enum AccountNames As Byte
            ADMINISTRATOR
            ADMINISTRATORS
            EVERYONE
            GUEST
            GUESTS
            LOCAL__SERVICE
            NETWORK
            NETWORK__SERVICE
            SERVICE
            SYSTEM
            USERS
        End Enum

        Private ___InUseCulture As OSCulture
        Public ReadOnly Property InUseCulture As OSCulture
            Get
                Return Me.___InUseCulture
            End Get
        End Property





        Public Sub New(pCulture As OSCulture)
            Me.___InUseCulture = pCulture
        End Sub



        Public Function getName(pAccount As AccountNames) As String
            Select Case Me.InUseCulture.ClassCultureType
                Case OSCulture.Cultures.ENGLISH___________USA, OSCulture.Cultures.ENGLISH___________UK
                    Select Case pAccount
                        Case AccountNames.ADMINISTRATOR
                            Return "ADMINISTRATOR"
                        Case AccountNames.ADMINISTRATORS
                            Return "ADMINISTRATOR"
                        Case AccountNames.EVERYONE
                            Return "EVERYONE"
                        Case AccountNames.GUEST
                            Return "GUEST"
                        Case AccountNames.GUESTS
                            Return "GUESTS"
                        Case AccountNames.LOCAL__SERVICE
                            Return "LOCAL SERVICE"
                        Case AccountNames.NETWORK
                            Return "NETWORK"
                        Case AccountNames.NETWORK__SERVICE
                            Return "NETWORK SERVICE"
                        Case AccountNames.SERVICE
                            Return "SERVICE"
                        Case AccountNames.SYSTEM
                            Return "SYSTEM"
                        Case AccountNames.USERS
                            Return "USERS"
                        Case Else
                            Throw New Exception("This Account Name is NOT supported. " & pAccount.ToString())
                    End Select
                Case OSCulture.Cultures.FRENCH_________FRANCE
                    Select Case pAccount
                        Case AccountNames.ADMINISTRATOR
                            Return "Administrateur"
                        Case AccountNames.ADMINISTRATORS
                            Return "Administrateurs"
                        Case AccountNames.EVERYONE
                            Return "Tout le monde"
                        Case AccountNames.GUEST
                            Return "Invité"
                        Case AccountNames.GUESTS
                            Return "Invités"
                        Case AccountNames.LOCAL__SERVICE
                            Return "SERVICE LOCAL"
                        Case AccountNames.NETWORK
                            Return "RESEAU"
                        Case AccountNames.NETWORK__SERVICE
                            Return "SERVICE RÉSEAU"
                        Case AccountNames.SERVICE
                            Return "SERVICE"
                        Case AccountNames.SYSTEM
                            Return "SYSTÈME"
                        Case AccountNames.USERS
                            Return "Utilisateurs"
                        Case Else
                            Throw New Exception("This Account Name is NOT supported. " & pAccount.ToString())
                    End Select
                Case Else
                    Throw New Exception("This culture is not Supported Yet. " &
                                        Environment.NewLine &
                                        Me.___InUseCulture.getClassCultureSummary()
                                        )
            End Select
        End Function




    End Class

End Namespace