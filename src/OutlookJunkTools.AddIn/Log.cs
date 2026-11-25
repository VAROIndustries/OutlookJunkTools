using System;
using System.IO;

namespace OutlookJunkTools.AddIn
{
    internal static class Log
    {
        private const long MaxBytes = 1024 * 1024;
        private static readonly object Gate = new object();

        public static string Directory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookJunkTools");

        private static string FilePath => Path.Combine(Directory, "log.txt");

        public static void Info(string message) => Write("INFO", message);

        public static void Error(string message, Exception ex) => Write("ERROR", message + " :: " + ex);

        private static void Write(string level, string message)
        {
            try
            {
                lock (Gate)
                {
                    System.IO.Directory.CreateDirectory(Directory);
                    var file = new FileInfo(FilePath);
                    if (file.Exists && file.Length > MaxBytes)
                    {
                        var old = Path.Combine(Directory, "log.old.txt");
                        if (File.Exists(old)) File.Delete(old);
                        File.Move(FilePath, old);
                    }
                    File.AppendAllText(FilePath,
                        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + level + " " + message + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never throw into Outlook.
            }
        }
    }
}
