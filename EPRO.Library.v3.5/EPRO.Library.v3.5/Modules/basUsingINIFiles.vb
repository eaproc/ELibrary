
Namespace Modules

    Module basUsingINIFiles

        'Private Declare Function GetPrivateProfileString _
        'Lib "kernel32" Alias "GetPrivateProfileStringA" _
        '(ByVal lpApplicationName As String, _
        'ByVal lpKeyName As Object, ByVal lpDefault As String, _
        'ByVal lpReturnedString As String, ByVal nSize As Long, _
        'ByVal lpFileName As String) As Long
        Private Declare Ansi Function GetPrivateProfileString _
      Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpDefault As String, _
      ByVal lpReturnedString As System.Text.StringBuilder, _
      ByVal nSize As Integer, ByVal lpFileName As String) _
      As Integer

        '<Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
        'Private Function GetPrivateProfileString(ByVal lpAppName As String, _
        '                        ByVal lpKeyName As String, _
        '                        ByVal lpDefault As String, _
        '                        ByVal lpReturnedString As String, _
        '                        ByVal nSize As Integer, _
        '                        ByVal lpFileName As String) As Integer
        'End Function

        <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
        Private Function WritePrivateProfileString _
    (ByVal lpApplicationName As String, ByVal lpKeyName _
    As String, ByVal lpString As String, ByVal lpFileName As _
    String) As Boolean

        End Function


        Public Function IREAD_APP_CONFIG(ByVal sectionName As String, ByVal KeyName As String, ByVal iniFileName As String, _
        Optional ByVal defaultValue As String = "", Optional ByVal valueSize As Integer = 256) As String

            '   Debug.Print IREAD_APP_CONFIG("MyApp", "Key1", App.Path & "\Data.ini")

            Dim BUFFER As New System.Text.StringBuilder(valueSize)
            'BUFFER = StrDup(valueSize, " ")

            Dim intCharCount As Integer =
                GetPrivateProfileString(sectionName, KeyName, defaultValue, BUFFER, valueSize, iniFileName)

            If intCharCount > 0 Then Return Left(BUFFER.ToString, intCharCount)
            'Remove vbnullcharacters
            'IREAD_APP_CONFIG = Replace(BUFFER, vbNullChar, "")

            Return ""
        End Function

        Public Function IWRITE_APP_CONFIG(ByVal sectionName As String, ByVal KeyName As String, ByVal defaultValue As String, _
        ByVal iniFileName As String) As Boolean

            '    Debug.Print WRITE_APP_CONFIG("MyApp", "Key1", _"Default value", App.Path & "\Data.ini")

            WritePrivateProfileString(sectionName, KeyName, defaultValue, iniFileName)

            IWRITE_APP_CONFIG = True

        End Function


    End Module


End Namespace