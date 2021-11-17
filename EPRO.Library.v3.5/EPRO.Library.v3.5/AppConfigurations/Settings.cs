using ELibrary.Standard.Modules;
using static ELibrary.Standard.Objects.EStrings;
using Microsoft.VisualBasic;

namespace ELibrary.Standard.AppConfigurations
{
    public class Settings
    {

        #region Using .ini files
        public static string READ_APP_CONFIG(string sectionName, string KeyName, string iniFileName, string defaultValue = Constants.vbNullString, int valueSize = 0, bool DO_CLEAN_UP = false)
        {
            string READ_APP_CONFIGRet = default;
            if (valueSize == 0)
                valueSize = 528;
            READ_APP_CONFIGRet = basUsingINIFiles.IREAD_APP_CONFIG(sectionName, KeyName, iniFileName, defaultValue, valueSize);
            if (DO_CLEAN_UP)
                READ_APP_CONFIGRet = CLEAN_UP_STRING(READ_APP_CONFIGRet);
            return READ_APP_CONFIGRet;
        }

        public static bool WRITE_APP_CONFIG(string sectionName, string KeyName, string keyValue, string iniFileName)
        {
            bool WRITE_APP_CONFIGRet = default;
            WRITE_APP_CONFIGRet = basUsingINIFiles.IWRITE_APP_CONFIG(sectionName, KeyName, keyValue, iniFileName);
            return WRITE_APP_CONFIGRet;
        }

        #endregion

    }
}