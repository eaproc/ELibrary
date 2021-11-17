using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.Modules
{
    static class basUsingINIFiles
    {

        // Private Declare Function GetPrivateProfileString _
        // Lib "kernel32" Alias "GetPrivateProfileStringA" _
        // (ByVal lpApplicationName As String, _
        // ByVal lpKeyName As Object, ByVal lpDefault As String, _
        // ByVal lpReturnedString As String, ByVal nSize As Long, _
        // ByVal lpFileName As String) As Long
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringA", CharSet = CharSet.Ansi)]
        private static extern int GetPrivateProfileString(string lpApplicationName, string lpKeyName, string lpDefault, System.Text.StringBuilder lpReturnedString, int nSize, string lpFileName);





        // <Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
        // Private Function GetPrivateProfileString(ByVal lpAppName As String, _
        // ByVal lpKeyName As String, _
        // ByVal lpDefault As String, _
        // ByVal lpReturnedString As String, _
        // ByVal nSize As Integer, _
        // ByVal lpFileName As String) As Integer
        // End Function

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool WritePrivateProfileString(string lpApplicationName, string lpKeyName, string lpString, string lpFileName);



        public static string IREAD_APP_CONFIG(string sectionName, string KeyName, string iniFileName, string defaultValue = "", int valueSize = 256)
        {

            // Debug.Print IREAD_APP_CONFIG("MyApp", "Key1", App.Path & "\Data.ini")

            var BUFFER = new System.Text.StringBuilder(valueSize);
            // BUFFER = StrDup(valueSize, " ")

            int intCharCount = basUsingINIFiles.GetPrivateProfileString(ref sectionName, ref KeyName, ref defaultValue, BUFFER, valueSize, ref iniFileName);
            if (intCharCount > 0)
                return Strings.Left(BUFFER.ToString(), intCharCount);
            // Remove vbnullcharacters
            // IREAD_APP_CONFIG = Replace(BUFFER, vbNullChar, "")

            return "";
        }

        public static bool IWRITE_APP_CONFIG(string sectionName, string KeyName, string defaultValue, string iniFileName)
        {
            bool IWRITE_APP_CONFIGRet = default;

            // Debug.Print WRITE_APP_CONFIG("MyApp", "Key1", _"Default value", App.Path & "\Data.ini")

            WritePrivateProfileString(sectionName, KeyName, defaultValue, iniFileName);
            IWRITE_APP_CONFIGRet = true;
            return IWRITE_APP_CONFIGRet;
        }
    }
}