using System;
using System.IO;

namespace Nedev.XlsToXlsx
{
    public static class Logger
    {
        public static LogLevel LogLevel { get; set; } = LogLevel.Info;
        private static StreamWriter? _logWriter;

        static Logger()
        {
            // 默认输出到控制台
        }

        public static void Initialize(string? logFilePath = null)
        {
            if (!string.IsNullOrEmpty(logFilePath))
            {
                try
                {
                    _logWriter = new StreamWriter(logFilePath, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to initialize log file: {ex.Message}");
                }
            }
        }

        public static void Log(LogLevel level, string message, Exception? ex = null)
        {
            if (level < LogLevel) return;

            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string logMessage = $"[{timestamp}] [{level}] {message}";

            if (ex != null)
            {
                logMessage += $"\nException: {ex.Message}\n{ex.StackTrace}";
            }

            // 输出到控制台
            Console.WriteLine(logMessage);

            // 输出到文件
            if (_logWriter != null)
            {
                try
                {
                    _logWriter.WriteLine(logMessage);
                    _logWriter.Flush();
                }
                catch { }
            }
        }

        public static void Debug(string message)
        {
            Log(LogLevel.Debug, message);
        }

        public static void Info(string message)
        {
            Log(LogLevel.Info, message);
        }

        public static void Warning(string message)
        {
            Log(LogLevel.Warning, message);
        }

        public static void Warn(string message)
        {
            Warning(message);
        }

        public static void Error(string message, Exception? ex = null)
        {
            Log(LogLevel.Error, message, ex);
        }

        public static void Fatal(string message, Exception? ex = null)
        {
            Log(LogLevel.Fatal, message, ex);
        }

        public static void Close()
        {
            _logWriter?.Dispose();
        }
    }

    public enum LogLevel
    {
        Debug = 0,
        Info = 1,
        Warning = 2,
        Error = 3,
        Fatal = 4
    }
}
