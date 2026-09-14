using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using System.Media;
using System.Windows;

namespace CalendarReminder.App;

public partial class MissedRemindersWindow : Window
{
    public MissedRemindersWindow(IReadOnlyList<ScheduleItem> items, DateTime from, DateTime to)
    {
        InitializeComponent();
        SummaryText.Text = $"从 {DateTimeFormat.Format(from)} 至 {DateTimeFormat.Format(to)}，共有 {items.Count} 条日程在程序未运行期间到达提醒时间。";
        ContentText.Text = ReminderSummaryFormatter.Build(items);
        Loaded += (_, _) => { AppLog.Info($"关机期间汇总提醒已显示，共 {items.Count} 条"); SystemSounds.Exclamation.Play(); Activate(); };
    }
    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
