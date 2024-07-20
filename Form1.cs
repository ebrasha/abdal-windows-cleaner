using Microsoft.Win32;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Abdal_Windows_Cleaner.Core;

namespace Abdal_Windows_Cleaner
{
    public partial class Form1 : Form
    {
        #region ShortcutGen

        [ComImport]
        [Guid("00021401-0000-0000-C000-000000000046")]
        private class ShellLink
        {
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214F9-0000-0000-C000-000000000046")]
        private interface IShellLink
        {
            void GetPath([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszFile, int cchMaxPath, ref IntPtr pfd,
                int fFlags);

            void GetIDList(out IntPtr ppidl);
            void SetIDList(IntPtr pidl);
            void GetDescription([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszName, int cchMaxName);
            void SetDescription([MarshalAs(UnmanagedType.LPWStr)] string pszName);
            void GetWorkingDirectory([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszDir, int cchMaxPath);
            void SetWorkingDirectory([MarshalAs(UnmanagedType.LPWStr)] string pszDir);
            void GetArguments([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszArgs, int cchMaxPath);
            void SetArguments([MarshalAs(UnmanagedType.LPWStr)] string pszArgs);
            void GetHotkey(out short pwHotkey);
            void SetHotkey(short wHotkey);
            void GetShowCmd(out int piShowCmd);
            void SetShowCmd(int iShowCmd);
            void GetIconLocation([Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pszIconPath, int cchIconPath);
            void SetIconLocation([MarshalAs(UnmanagedType.LPWStr)] string pszIconPath);
            void SetRelativePath([MarshalAs(UnmanagedType.LPWStr)] string pszPathRel, int dwReserved);
            void Resolve(IntPtr hwnd, int fFlags);
            void SetPath([MarshalAs(UnmanagedType.LPWStr)] string pszFile);
        }

        #endregion


        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SHEmptyRecycleBin(IntPtr hwnd, string pszRootPath, uint dwFlags);

        // Constants for SHEmptyRecycleBin function
        const uint SHERB_NOCONFIRMATION = 0x00000001;
        const uint SHERB_NOPROGRESSUI = 0x00000002;
        const uint SHERB_NOSOUND = 0x00000004;

        public Form1()
        {
            InitializeComponent();
            this.Hide();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string executablePath = Process.GetCurrentProcess().MainModule.FileName;
            string executableName = Path.GetFileName(executablePath);
            string shortcutLocation = Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" +
                                      executableName.Replace(".exe", ".lnk");
            // CreateShortcut(executablePath, shortcutLocation);

            string fullPath = Process.GetCurrentProcess().MainModule.FileName;
            // ShortcutMng.CreateDesktopShortcut(executableName.Replace(".exe", ""), fullPath);
            ShortcutMng.CreateStartupScript(fullPath);


            // CreateShortcutOnRegistry(fullPath);
            // Empty the Recycle Bin without confirmation, progress UI, or sound
            int result = SHEmptyRecycleBin(IntPtr.Zero, null,
                SHERB_NOCONFIRMATION | SHERB_NOPROGRESSUI | SHERB_NOSOUND);
            string tempPath = Path.GetTempPath();
            // DEl Temp
            DeleteDirectoryContents(tempPath);
            // Remove Windows Event
            ClearEventLogs();
            // Delete Log file
            string[] directories =
            {
                Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData), //  Program Data
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Logs"), // C:\Windows\Logs
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), // AppData\Local
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), // AppData\Roaming
            };

            foreach (string dir in directories)
            {
                try
                {
                    DeleteLargeLogFiles(dir);
                }
                catch (Exception ex)
                {
                    // Console.WriteLine($"Failed to process directory {dir}: {ex.Message}");
                }
            }


            Environment.Exit(0);
        }

        #region RemoveLog

        static void DeleteLargeLogFiles(string rootDir)
        {
            try
            {
                ProcessDirectory(rootDir);
            }
            catch
            {
            }
        }

        static void ProcessDirectory(string currentDir)
        {
            try
            {
                foreach (string file in Directory.GetFiles(currentDir, "*.*"))
                {
                    try
                    {
                        string extension = Path.GetExtension(file);
                        if ((extension == ".txt" || extension == ".log") &&
                            new FileInfo(file).Length > 40 * 1024 * 1024)
                        {
                            try
                            {
                                File.Delete(file);
                                // MessageBox.Show($"Deleted file: {file}");
                            }
                            catch
                            {
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                foreach (string dir in Directory.GetDirectories(currentDir))
                {
                    try
                    {
                        ProcessDirectory(dir);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        #endregion

        static void ClearEventLogs()
        {
            foreach (var log in EventLog.GetEventLogs())
            {
                try
                {
                    // MessageBox.Show($"Clearing {log.Log}");
                    log.Clear();
                }
                catch (Exception ex)
                {
                    // Console.WriteLine($"Failed to clear {log.Log}: {ex.Message}");
                }
            }
        }


        static void DeleteDirectoryContents(string path)
        {
            try
            {
                DirectoryInfo directory = new DirectoryInfo(path);

                foreach (FileInfo file in directory.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                    }
                }

                foreach (DirectoryInfo subDirectory in directory.GetDirectories())
                {
                    try
                    {
                        subDirectory.Delete(true);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        #region Create Shortcut To startup in Reg

        public static void CreateShortcutOnRegistry(string filePath)
        {
            try
            {
                string keyName = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(keyName, writable: true))
                {
                    if (key == null)
                    {
                        // Failed to open registry key, silently handle error
                        return;
                    }

                    string valueName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    key.SetValue(valueName, filePath);
                    // Successfully added to registry, silently handle success
                }
            }
            catch (Exception)
            {
                // Silently handle any exceptions
            }
        }

        #endregion

        #region ShortcutGen2

        static void CreateShortcutOnStartupDir(string targetPath, string shortcutLocation)
        {
            IShellLink link = (IShellLink)new ShellLink();

            link.SetDescription("My Shortcut");
            link.SetPath(targetPath);
            link.SetWorkingDirectory(System.IO.Path.GetDirectoryName(targetPath));

            IPersistFile file = (IPersistFile)link;
            file.Save(shortcutLocation, false);

            Console.WriteLine("Shortcut created successfully!");
        }

        // CreateShortcut(executablePath, shortcutLocation);
        [ComImport]
        [Guid("0000010b-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPersistFile
        {
            void GetClassID(out Guid pClassID);

            [PreserveSig]
            int IsDirty();

            void Load([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, uint dwMode);
            void Save([MarshalAs(UnmanagedType.LPWStr)] string pszFileName, bool fRemember);
            void SaveCompleted([MarshalAs(UnmanagedType.LPWStr)] string pszFileName);
            void GetCurFile([MarshalAs(UnmanagedType.LPWStr)] out string ppszFileName);
        }

        #endregion
    }
}