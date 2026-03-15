using System;
using System.IO;

namespace denemelikimid
{
    internal static class SimpleLogger
    {
        private static readonly object _sync = new object();
        private static readonly string _logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        private static readonly string _logFile = Path.Combine(_logDir, "app.log");

        public static void Log(string message)
        {
            try
            {
                lock (_sync)
                {
                    if (!Directory.Exists(_logDir)) Directory.CreateDirectory(_logDir);
                    File.AppendAllText(_logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // swallow logging exceptions to avoid secondary failures
            }
        }
    }
}
