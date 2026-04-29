using System;
using System.IO;

namespace PanaceaIEWrapper.Bot
{
    internal sealed class BotLogger
    {
        private readonly string _logPath;

        public BotLogger(string logPath)
        {
            _logPath = logPath;
            string dir = Path.GetDirectoryName(_logPath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }
        }

        public void Info(string message)
        {
            Write("INFO", message);
        }

        public void Error(string message)
        {
            Write("ERROR", message);
        }

        private void Write(string level, string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}";
            File.AppendAllText(_logPath, line + Environment.NewLine);
        }
    }
}
