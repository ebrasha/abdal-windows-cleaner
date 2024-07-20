using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using WindowsShortcutFactory;

namespace Abdal_Windows_Cleaner.Core
{
    internal class ShortcutMng
    {
        public static void CreateDesktopShortcut(string shortcutName, string targetPath)
        {
            string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string shortcutLocation = System.IO.Path.Combine(startupPath, shortcutName + ".lnk");

            using var shortcut = new WindowsShortcut
            {
                Path = targetPath,
                Arguments = "runas",
                WorkingDirectory = Path.GetDirectoryName(targetPath),
                Description = "Secure DataLink Manager"
            };
            shortcut.Save(shortcutLocation);
        }

        #region VBS Method

        public static void CreateStartupScript(string exePath)
        {
            string executableName = Path.GetFileName(exePath).Replace(".exe","");
            string startupFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
            string vbsFileName = Path.Combine(startupFolderPath, executableName+".vbs");

            string vbsContent = $@"

Dim fso
Set fso = CreateObject(""Scripting.FileSystemObject"")

' Define the file path
filePath = ""{exePath}""

' Check if the file exists
If fso.FileExists(filePath) Then
    If WScript.Arguments.length = 0 Then
        Set objShell = CreateObject(""Shell.Application"")
        ' Request admin privileges
        objShell.ShellExecute ""wscript.exe"", """""""" & WScript.ScriptFullName & """""" uac"", """", ""runas"", 1
    Else
        ' Your main script
        Dim objShell
        Set objShell = WScript.CreateObject(""WScript.Shell"")
        objShell.Run("""""""" & filePath & """""""")
        Set objShell = Nothing
    End If
End If

' Release resources
Set fso = Nothing


";

            File.WriteAllText(vbsFileName, vbsContent);
        }

        #endregion
    }
}