using System;
using System.IO;
using System.Text;

namespace ProofioAddIn.Services
{
    public static class Logger
    {
        private const long MaxBytes = 5L * 1024L * 1024L;

        private static readonly object Sync = new object();

        private static string LogPath
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "Proofio");

                Directory.CreateDirectory(dir);
                return Path.Combine(dir, "proofio.log");
            }
        }

        public static void Info(string message)
        {
            Write("INFO", message, null);
        }

        public static void Error(string message, Exception ex)
        {
            Write("ERROR", message, ex);
        }

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                lock (Sync)
                {
                    RollIfNeeded();

                    var builder = new StringBuilder();
                    builder.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                    builder.Append(" [");
                    builder.Append(level);
                    builder.Append("] ");
                    builder.AppendLine(message ?? string.Empty);

                    if (ex != null)
                    {
                        builder.AppendLine(ex.ToString());
                    }

                    File.AppendAllText(LogPath, builder.ToString(), Encoding.UTF8);
                }
            }
            catch
            {
            }
        }

        private static void RollIfNeeded()
        {
            try
            {
                var path = LogPath;
                if (!File.Exists(path)) return;

                var info = new FileInfo(path);
                if (info.Length < MaxBytes) return;

                var backup = path + ".1";
                if (File.Exists(backup)) File.Delete(backup);
                File.Move(path, backup);
            }
            catch
            {
            }
        }
    }
}
