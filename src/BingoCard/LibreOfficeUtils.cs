using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace BingoCard
{
    public static class LibreOfficeUtils
    {
        public static (bool found, string sofficePath) FindLibreOfficeBinary()
        {
            return IsLinux
                ? (false, default)
                : FindSofficeWindows();

            (bool found, string path) FindSofficeWindows()
            {
                foreach (var path in WindowsLibreOfficeBinaryCandidates())
                {
                    if (File.Exists(path))
                    {
                        return (true, path);
                    }
                }

                return (false, default);
            }

            static IEnumerable<string> WindowsLibreOfficeBinaryCandidates()
            {
                var subPath = Path.Combine("LibreOffice", "program", LibreofficeAppName);

                return new[]
                    {
                        Environment.SpecialFolder.ProgramFiles,
                        Environment.SpecialFolder.ProgramFilesX86,
                        Environment.SpecialFolder.CommonProgramFiles,
                        Environment.SpecialFolder.CommonProgramFilesX86,
                    }
                   .Select(Environment.GetFolderPath)
                   .Select(d => Path.Combine(d, subPath));
            }
        }

        public static string ExpectedProfileDir => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "libreoffice");

        public static bool HasLibreOfficeProfile()
        {
            if (!Directory.Exists(ExpectedProfileDir))
                return false;

            var versionDir = Directory.GetDirectories(ExpectedProfileDir).FirstOrDefault();

            return versionDir != null;
        }

        public static bool IsLibreOfficeOnPath()
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        WindowStyle = ProcessWindowStyle.Hidden,
                        FileName = LibreofficeAppName,
                        Arguments = "/C --headless"
                    }
                };

                process.Start();
                process.WaitForExit();

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string LibreofficeAppName => IsLinux ? "libreoffice" : "soffice.exe";
        private static bool IsLinux => System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Linux);
    }
}