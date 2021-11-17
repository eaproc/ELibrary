Imports Microsoft.Win32

Namespace Modules

    Friend Module basRegistry
        ''' <summary>
        ''' For 64bit Registry
        ''' </summary>
        ''' <param name="ParentKey"></param>
        ''' <param name="SubKeyName"></param>
        ''' <param name="Writeable"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function OpenSubKey(ByVal ParentKey As RegistryHive, ByVal SubKeyName As String, ByVal Writeable As Boolean) As Microsoft.Win32.RegistryKey
            REM 64bit access
            Dim rk1 As RegistryKey = RegistryKey.OpenBaseKey(ParentKey, RegistryView.Registry64)

            Try


                Return rk1.OpenSubKey(SubKeyName, Writeable)

            Catch ex As Exception

                Return Nothing

            End Try



        End Function

    End Module

End Namespace