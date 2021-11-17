using System;
using System.Collections;
using ELibrary.Standard.Modules;
using ELibrary.Standard.Objects;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ELibrary.Standard.AppConfigurations
{


    /// <summary>
    /// Currently supports win7 and windows 8.1
    /// </summary>
    /// <remarks></remarks>
    public sealed class DesktopShortcutPinner
    {

        /// <summary>
        /// Not case sensitive
        /// </summary>
        /// <param name="pShortcutNameWithoutExtension"></param>
        /// <remarks></remarks>
        public DesktopShortcutPinner(string pShortcutNameWithoutExtension, bool PinIt = true)
        {


            // REM Worked perfectly with windows 7, windows 8.1

            // Const CSIDL_COMMON_PROGRAMS = &H17
            object ShellApp;
            object FSO;
            object Desktop;
            ShellApp = Interaction.CreateObject("Shell.Application");
            FSO = Interaction.CreateObject("Scripting.FileSystemObject");

            // Set StartMenuFolder = ShellApp.NameSpace(CSIDL_COMMON_PROGRAMS)
            Desktop = ShellApp.NameSpace(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
            object LnkFile;
            LnkFile = Operators.ConcatenateObject(Desktop.Self.Path, string.Format(@"\{0}.lnk", pShortcutNameWithoutExtension));
            if (Conversions.ToBoolean(FSO.FileExists(LnkFile)))
            {
                object desktopImtes;
                desktopImtes = Desktop.Items();
                foreach (var item in (IEnumerable)desktopImtes)
                {
                    // -- - Debug.Print("item.Name: " & item.Name)
                    if ((item.Name.ToString().ToLower() ?? "") == (pShortcutNameWithoutExtension.ToLower() ?? ""))
                    {
                        foreach (var verb in (IEnumerable)item.Verbs)
                        {


                            // Debug.Print("verb.Name: " & verb.Name)
                            if (EStrings.valueOf(verb.Name).equalsIgnoreCase("Pin to Tas&kbar") && PinIt || EStrings.valueOf(verb.Name).equalsIgnoreCase("Unpin from Tas&kbar") && !PinIt) // If (verb.Name = "锁定到任务栏(&K)")
                            {
                                verb.DoIt();
                                goto END__CREATING;
                            }

                            // 'If IsWin10 Then



                            // 'Else

                            // '    'Debug.Print("verb.Name: " & verb.Name)
                            // '    If (
                            // '        (verb.Name = "Pin to Tas&kbar") AndAlso PinIt
                            // '        ) OrElse
                            // '    (
                            // '        (verb.Name = "Unpin from Tas&kbar") AndAlso Not PinIt
                            // '        ) Then 'If (verb.Name = "锁定到任务栏(&K)")
                            // '        verb.DoIt()
                            // '        GoTo END__CREATING
                            // '    End If


                            // 'End If


                        }
                    }
                }
            }
            else
            {
                basMain.MyLogFile.Log(Operators.ConcatenateObject("Link does not exist: ", LnkFile));
            }

        END__CREATING:
            ;
            FSO = null;
            ShellApp = null;
        }
    }
}