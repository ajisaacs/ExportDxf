using System;
using System.IO;

namespace ExportDXF.Services
{
    public class LogFileService : IDisposable
    {
        private StreamWriter _exportLog;
        private static readonly string AppLogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "ExportDXF", "ExportDXF.log");

        public void StartExportLog(string logFilePath)
        {
            _exportLog?.Dispose();

            var dir = Path.GetDirectoryName(logFilePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            _exportLog = new StreamWriter(logFilePath, append: true);
        }

        public void Log(string level, string message)
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";

            _exportLog?.WriteLine(line);
            _exportLog?.Flush();

            WriteAppLog(line);
        }

        public void LogInfo(string message) => Log("INFO", message);
        public void LogWarning(string message) => Log("WARNING", message);
        public void LogError(string message) => Log("ERROR", message);

        private void WriteAppLog(string line)
        {
            try
            {
                var dir = Path.GetDirectoryName(AppLogPath);
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                File.AppendAllText(AppLogPath, line + Environment.NewLine);
            }
            catch
            {
                // Best-effort app log — don't fail exports if log write fails
            }
        }

        public void Dispose()
        {
            _exportLog?.Dispose();
            _exportLog = null;
        }
    }
}
