using System;
using System.Collections.Generic;
using System.IO;
using Shell32;

namespace ELibrary.Standard.AppConfigurations
{
    public sealed class Shortcuts
    {
        public struct ShortcutLink
        {
            public ShortcutLink(string FullName, string Arguments, string Target)
            {
                __FullName = FullName;
                __Arguments = Arguments;
                __Target = Target;
            }

            private string __FullName;
            private string __Arguments;
            private string __Target;

            public string Name
            {
                get
                {
                    return EIO.getFileName(FullName);
                }
            }

            public string FullName
            {
                get
                {
                    return __FullName;
                }
            }

            public string Target
            {
                get
                {
                    return __Target;
                }
            }

            public string Arguments
            {
                get
                {
                    return __Arguments;
                }
            }
        }

        public static IEnumerable<FileInfo> getAllShortcutsOn(string FolderPath)
        {
            var rst = new List<FileInfo>();
            try
            {
                var di = new DirectoryInfo(FolderPath);
                return di.GetFiles("*lnk");
            }
            catch (Exception ex)
            {
            }

            return rst;
        }

        public static ShortcutLink getShortcut(FileInfo linkFile)
        {
            try
            {
                Shell shell = new ShellClass();
                var folder = shell.NameSpace(linkFile.DirectoryName);
                var folderItem = folder.ParseName(linkFile.Name);
                ShellLinkObject link = (ShellLinkObject)folderItem.GetLink;
                return new ShortcutLink(linkFile.FullName, link.Arguments, link.Path);
            }
            catch (Exception ex)
            {
                return default;
            }
        }

        public static bool createShortcut(string PathLinkWithoutDot, string TargetFile, string Arguments = "", string IconPath = "", string comments = "Launch")
        {
            try
            {
                var wshShell = new IWshRuntimeLibrary.WshShellClass();
                IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)wshShell.CreateShortcut(PathLinkWithoutDot + ".lnk");
                shortcut.Arguments = Arguments;
                shortcut.TargetPath = TargetFile;
                shortcut.WorkingDirectory = Microsoft.VisualBasic.FileIO.FileSystem.GetParentPath(TargetFile);
                if (IconPath is object && !string.IsNullOrEmpty(IconPath))
                {
                    shortcut.IconLocation = IconPath;
                }

                shortcut.Description = comments;
                shortcut.Save();
                return RefreshDesktopShortcuts();
            }
            catch (Exception ex)
            {
                Modules.basMain.MyLogFile.Print("createShortcut", ex);
                return false;
            }
        }

        public static bool RefreshDesktopShortcuts()
        {
            try
            {

                // REM Clear Icon Cache for windows 7 and 8 to reflect right icon
                var cmd = new MicrosoftOS.CommandPrompt(false);
                if (!Directory.Exists(@"C:\Windows\SysWOW64"))
                {
                    // REM 32bit
                    cmd.Execute("ie4uinit.exe -ClearIconCache");
                }
                else
                {
                    // REM 64bit
                    // REM cmd.Execute("ie4uinit.exe -ClearIconCache", "%windir%\sysnative")
                    cmd.Execute(@"cd %windir%\sysnative & ie4uinit.exe -ClearIconCache");
                    // REM Interaction.Shell("cmd.exe /K ""cd %windir%\sysnative & ie4uinit.exe -ClearIconCache""", AppWinStyle.NormalFocus, True)

                }

                cmd = null;

                // OR
                // '                CD /d %userprofile%\AppData\Local
                // '                DEL IconCache.db / a
                // 'EXIT


                // '                restart Process
                // '                explorer()


                return true;
            }
            catch (Exception ex)
            {
                Modules.basMain.MyLogFile.Print("RefreshDesktopShortcuts", ex);
                return false;
            }
        }
    }
}