Imports System.Text
Imports System.Globalization


Namespace MicrosoftOS
    Public Class OSCulture

#Region "Properties"


        Private _____ClassCulture As CultureInfo
        Public ReadOnly Property ClassCulture As CultureInfo
            Get
                Return Me._____ClassCulture
            End Get
        End Property


        Public ReadOnly Property ClassCultureType As OSCulture.Cultures
            Get
                Return CType(Me.ClassCulture.LCID, Cultures)
            End Get
        End Property

#End Region


#Region "Enums"

        Public Enum Cultures
            ENGLISH___________USA = 1033
            ENGLISH___________UK = 2057
            FRENCH_________FRANCE = 1036
            UNKNOWN______________ = -1
        End Enum

#End Region


#Region "Methods"

        ''' <summary>
        ''' Throws Exception Exception for not supported
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getOSCultureInstalledOnPC() As OSCulture
            Try
                Return New OSCulture(CType(CultureInfo.InstalledUICulture.LCID, OSCulture.Cultures))
            Catch ex As Exception
                Throw New Exception("Please, check to see that this OS Culture is Supported: " & Environment.NewLine() &
                                  getInstalledCultureSummary()
                                                   )
            End Try
        End Function


        ''' <summary>
        ''' Throws Exception Exception for not supported
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getOSCultureForCurrentThread() As OSCulture
            Try
                Return New OSCulture(CType(CultureInfo.CurrentUICulture.LCID, OSCulture.Cultures))
            Catch ex As Exception
                Throw New Exception("Please, check to see that this OS Culture is Supported: " & Environment.NewLine() &
                                  getCurrentThreadCultureSummary()
                                                   )
            End Try
        End Function


        ''' <summary>
        ''' The summary of the OS Current UI Culture on this calling thread. It is usually same if user didn't create specific culture using  CultureInfo.CreateSpecificCulture("fr-FR")
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function getClassCultureSummary() As String
            Return getCultureSummary(Me._____ClassCulture)
        End Function


        ''' <summary>
        ''' The summary of the OS Current UI Culture on this calling thread. It is usually same if user didn't create specific culture using  CultureInfo.CreateSpecificCulture("fr-FR")
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getCurrentThreadCultureSummary() As String
            Return getCultureSummary(CultureInfo.CurrentUICulture)
        End Function


        ''' <summary>
        ''' The summary of the OS Installed Culture
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function getInstalledCultureSummary() As String
            Return getCultureSummary(CultureInfo.InstalledUICulture)
        End Function

        Public Shared Function getCultureSummary(pCulture As CultureInfo) As String

            Dim dCulture As CultureInfo = pCulture

            Dim result As StringBuilder = New StringBuilder()
            With result
                .AppendLine("Default Language Info: ")
                .AppendLine("Name: " & vbTab & vbTab & vbTab & dCulture.Name)
                .AppendLine("Display Name: " & vbTab & vbTab & vbTab & dCulture.DisplayName)
                .AppendLine("English Name: " & vbTab & vbTab & vbTab & dCulture.EnglishName)
                .AppendLine("2-letter ISO Name: " & vbTab & vbTab & dCulture.TwoLetterISOLanguageName)
                .AppendLine("3-letter ISO Name: " & vbTab & vbTab & dCulture.ThreeLetterISOLanguageName)
                .AppendLine("3-letter Win32 API Name: " & vbTab & dCulture.ThreeLetterWindowsLanguageName)
                .AppendLine("LCID Identifier: " & vbTab & vbTab & dCulture.LCID)
                .AppendLine("NativeName: " & vbTab & vbTab & vbTab & dCulture.NativeName)
                .AppendLine("DateTimeFormat: " & vbTab & vbTab & dCulture.DateTimeFormat.ToString())

                Return .ToString()
            End With


        End Function

#End Region




#Region "Constructors"

        ''' <summary>
        ''' Use ENGLISH___________USA 
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub New()
            Me.New(Cultures.ENGLISH___________USA)
        End Sub


        ''' <summary>
        ''' If unknown is specified it switches to English
        ''' </summary>
        ''' <param name="pCulture"></param>
        ''' <remarks></remarks>
        Public Sub New(pCulture As Cultures)
            If pCulture <> Cultures.UNKNOWN______________ Then
                Me._____ClassCulture = New CultureInfo(pCulture)
            Else
                Me._____ClassCulture = New CultureInfo(Cultures.ENGLISH___________USA)
            End If
        End Sub


#End Region



    End Class
End Namespace