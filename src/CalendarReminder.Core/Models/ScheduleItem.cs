namespace CalendarReminder.Core.Models;

public sealed class ScheduleItem
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public DateTime ScheduledAt { get; set; }
    public DateTime ReminderAt { get; set; }
    public string Content { get; set; } = "";
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
    public bool ReminderTriggered { get; set; }
    public string DisplayTime => DateTimeFormat.Format(ScheduledAt);
    public string DisplayReminderTime => DateTimeFormat.Format(ReminderAt == default ? ScheduledAt : ReminderAt);
    public string Summary => Content.Length <= 100 ? Content : Content[..100] + "…";
}

public enum ImportStatus { 可导入, 需要确认, 日期无效, 内容为空, 疑似重复, 高度重合 }
public enum DuplicateAction { 跳过重复项, 仍然导入, 替换原日程 }

public sealed class ImportCandidate
{
    public bool IsSelected { get; set; } = true;
    public DateTime? ScheduledAt { get; set; }
    public DateTime? ReminderAt { get; set; }
    public string Content { get; set; } = "";
    public ImportStatus Status { get; set; }
    public string ErrorReason { get; set; } = "";
    public bool IsDuplicate { get; set; }
    public DateTime? Date { get => ScheduledAt?.Date; set { if (value.HasValue) { var oldDate=ScheduledAt?.Date; ScheduledAt=value.Value.Date.Add(ScheduledAt?.TimeOfDay??TimeSpan.Zero); if(oldDate.HasValue&&ReminderAt.HasValue) ReminderAt=ReminderAt.Value.Add(value.Value.Date-oldDate.Value); } else ScheduledAt=null; } }
    public int Hour { get => ScheduledAt?.Hour ?? 0; set { if (ScheduledAt.HasValue && value is >= 0 and <= 23) ScheduledAt = ScheduledAt.Value.Date.AddHours(value).AddMinutes(ScheduledAt.Value.Minute); } }
    public int Minute { get => ScheduledAt?.Minute ?? 0; set { if (ScheduledAt.HasValue && value is >= 0 and <= 59) ScheduledAt = ScheduledAt.Value.Date.AddHours(ScheduledAt.Value.Hour).AddMinutes(value); } }
    public string DisplayTime => ScheduledAt.HasValue ? DateTimeFormat.Format(ScheduledAt.Value) : "";
    public string DisplayReminderTime => ReminderAt.HasValue ? DateTimeFormat.Format(ReminderAt.Value) : "";
}

public sealed record ExportResult(IReadOnlyList<string> Files, int ItemCount);
public sealed record ImportResult(int Imported, int Skipped, int Replaced);
