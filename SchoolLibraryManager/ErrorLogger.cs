using System.IO;
using System.Text;

namespace SchoolLibraryManager;

public static class ErrorLogger
{
    private static readonly object Sync = new();
    private static readonly string LogPath = Path.Combine(AppContext.BaseDirectory, "ErrorLog.txt");

    public static void Log(Exception exception, string? context = null)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR");

        if (!string.IsNullOrWhiteSpace(context))
        {
            builder.AppendLine($"Context: {context}");
        }

        builder.AppendLine(exception.ToString());
        builder.AppendLine(new string('-', 80));

        Write(builder.ToString());
    }

    public static void Log(string message, string? context = null)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO");

        if (!string.IsNullOrWhiteSpace(context))
        {
            builder.AppendLine($"Context: {context}");
        }

        builder.AppendLine(message);
        builder.AppendLine(new string('-', 80));

        Write(builder.ToString());
    }

    private const long MaxLogBytes = 2 * 1024 * 1024;

    private static string FallbackLogPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SchoolLibraryManager", "ErrorLog.txt");

    private static void Write(string content)
    {
        lock (Sync)
        {
            if (TryWrite(LogPath, content)) return;
            TryWrite(FallbackLogPath, content);
        }
    }

    private static bool TryWrite(string path, string content)
    {
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

            var info = new FileInfo(path);
            if (info.Exists && info.Length > MaxLogBytes)
            {
                var archive = Path.ChangeExtension(path, ".old.txt");
                File.Move(path, archive, true);
            }

            File.AppendAllText(path, content, Encoding.UTF8);
            return true;
        }
        catch
        {
            return false;
        }
    }
}