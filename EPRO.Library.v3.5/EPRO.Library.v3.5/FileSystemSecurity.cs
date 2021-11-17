using System.IO;
using ELibrary.Standard.MicrosoftOS;

namespace ELibrary.Standard
{
    public class FileSystemSecurity
    {


        /// <summary>
    /// Using Installed OSCulture
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFolderPath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void RemoveAccountFromFolder(string pFolderPath, PCAccountNames.AccountNames pAccount)
        {
            RemoveAccountFromFolder(pFolderPath, new PCAccountNames(OSCulture.getOSCultureInstalledOnPC()), pAccount);
        }

        /// <summary>
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFolderPath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void RemoveAccountFromFolder(string pFolderPath, PCAccountNames PCAccountHandler, PCAccountNames.AccountNames pAccount)
        {
            var di = new DirectoryInfo(pFolderPath);
            var AccessRule = new System.Security.AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount), System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow);
            var dSecurity = di.GetAccessControl();
            dSecurity.RemoveAccessRule(AccessRule);
            di.SetAccessControl(dSecurity);
        }



        /// <summary>
    /// Using Installed OSCulture
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFolderPath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void AddAccountToFolder(string pFolderPath, PCAccountNames.AccountNames pAccount)
        {
            AddAccountToFolder(pFolderPath, new PCAccountNames(OSCulture.getOSCultureInstalledOnPC()), pAccount);
        }

        /// <summary>
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFolderPath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void AddAccountToFolder(string pFolderPath, PCAccountNames PCAccountHandler, PCAccountNames.AccountNames pAccount)
        {
            var di = new DirectoryInfo(pFolderPath);
            var AccessRule = new System.Security.AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount), System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow);
            var dSecurity = di.GetAccessControl();
            dSecurity.AddAccessRule(AccessRule);
            di.SetAccessControl(dSecurity);
        }



        /// <summary>
    /// Using Installed OSCulture
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFilePath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void AddAccountToFile(string pFilePath, PCAccountNames.AccountNames pAccount)
        {
            AddAccountToFile(pFilePath, new PCAccountNames(OSCulture.getOSCultureInstalledOnPC()), pAccount);
        }


        /// <summary>
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFilePath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void AddAccountToFile(string pFilePath, PCAccountNames PCAccountHandler, PCAccountNames.AccountNames pAccount)
        {
            var di = new FileInfo(pFilePath);
            var AccessRule = new System.Security.AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount), System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow);
            var fi = new FileInfo(pFilePath);
            var fSecurity = fi.GetAccessControl();
            fSecurity.AddAccessRule(AccessRule);
            fi.SetAccessControl(fSecurity);
        }


        /// <summary>
    /// Using Installed OSCulture
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFilePath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void RemoveAccountFromFile(string pFilePath, PCAccountNames.AccountNames pAccount)
        {
            RemoveAccountFromFile(pFilePath, new PCAccountNames(OSCulture.getOSCultureInstalledOnPC()), pAccount);
        }

        /// <summary>
    /// NB: This does not handle errors like File or Folder NOT Found
    /// </summary>
    /// <param name="pFilePath"></param>
    /// <param name="pAccount"></param>
    /// <remarks></remarks>
        public static void RemoveAccountFromFile(string pFilePath, PCAccountNames PCAccountHandler, PCAccountNames.AccountNames pAccount)
        {
            var di = new FileInfo(pFilePath);
            var AccessRule = new System.Security.AccessControl.FileSystemAccessRule(PCAccountHandler.getName(pAccount), System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow);
            var fi = new FileInfo(pFilePath);
            var fSecurity = fi.GetAccessControl();
            fSecurity.RemoveAccessRule(AccessRule);
            fi.SetAccessControl(fSecurity);
        }
    }
}