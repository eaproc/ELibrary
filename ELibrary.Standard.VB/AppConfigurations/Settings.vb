Imports ELibrary.Standard.VB.Modules
Imports ELibrary.Standard.VB.Objects.EStrings

Namespace AppConfigurations

    Public Class Settings

#Region "Using .ini files"
        Public Shared Function READ_APP_CONFIG(ByVal sectionName As String,
                                        ByVal KeyName As String,
                                        ByVal iniFileName As String, _
    Optional ByVal defaultValue As String = vbNullString,
    Optional ByVal valueSize As Integer = 0,
    Optional ByVal DO_CLEAN_UP As Boolean = False) As String

            If valueSize = 0 Then valueSize = 528
            READ_APP_CONFIG = IREAD_APP_CONFIG(sectionName, KeyName, iniFileName, defaultValue, valueSize)

            If DO_CLEAN_UP Then READ_APP_CONFIG = CLEAN_UP_STRING(READ_APP_CONFIG)

        End Function


        Public Shared Function WRITE_APP_CONFIG(ByVal sectionName As String,
                                         ByVal KeyName As String,
                                         ByVal keyValue As String, _
    ByVal iniFileName As String) As Boolean

            WRITE_APP_CONFIG = IWRITE_APP_CONFIG(sectionName, KeyName, keyValue, iniFileName)

        End Function

#End Region

    End Class

End Namespace
