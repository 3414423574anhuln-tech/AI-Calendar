namespace CalendarReminder.Core.Services;

public static class AppPaths
{
    public static string DataDirectory => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CalendarReminder");
    public static string DatabasePath => Path.Combine(DataDirectory, "calendar.db");
    public static string LogDirectory => Path.Combine(DataDirectory, "logs");
    public static string AiSettingsPath => Path.Combine(DataDirectory, "api-settings.json");
    public static string AiImportDirectory => Path.Combine(DataDirectory, "ai-imports");
    public static string ReminderCheckpointPath => Path.Combine(DataDirectory, "last-reminder-check.txt");
    public static void Ensure() { Directory.CreateDirectory(DataDirectory); Directory.CreateDirectory(LogDirectory); Directory.CreateDirectory(AiImportDirectory); }
}

public sealed class ReminderCheckpointStore
{
    private readonly string _path;
    public ReminderCheckpointStore(string? path = null) => _path = path ?? AppPaths.ReminderCheckpointPath;
    public DateTime? Read()
    {
        try
        {
            if (!File.Exists(_path)) return null;
            return DateTime.TryParse(File.ReadAllText(_path), System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind, out var value) ? value : null;
        }
        catch (Exception ex) { AppLog.Error("读取提醒检查点", ex); return null; }
    }
    public void Write(DateTime value)
    {
        try { Directory.CreateDirectory(Path.GetDirectoryName(_path)!); File.WriteAllText(_path, value.ToString("O", System.Globalization.CultureInfo.InvariantCulture)); }
        catch (Exception ex) { AppLog.Error("保存提醒检查点", ex); }
    }
}

public static class ReminderSummaryFormatter
{
    public static string Build(IEnumerable<CalendarReminder.Core.Models.ScheduleItem> items)
        => string.Join("\n\n", items.OrderBy(x => x.ReminderAt).ThenBy(x => x.ScheduledAt).Select((x, i) =>
            $"{i + 1}. 日程时间：{CalendarReminder.Core.DateTimeFormat.Format(x.ScheduledAt)}\n提醒时间：{CalendarReminder.Core.DateTimeFormat.Format(x.ReminderAt)}\n安排：{x.Content}"));
}

public static class AppLog
{
    private static readonly object Gate = new();
    public static void Info(string message)
    {
        try { AppPaths.Ensure(); lock (Gate) File.AppendAllText(Path.Combine(AppPaths.LogDirectory, $"{DateTime.Now:yyyy-MM-dd}.log"), $"[{DateTime.Now:O}] INFO {message}\r\n"); }
        catch { }
    }
    public static void Error(string context, Exception ex)
    {
        try
        {
            AppPaths.Ensure();
            lock (Gate)
                File.AppendAllText(Path.Combine(AppPaths.LogDirectory, $"{DateTime.Now:yyyy-MM-dd}.log"), $"[{DateTime.Now:O}] {context}\r\n{ex}\r\n\r\n");
        }
        catch { }
    }
}
