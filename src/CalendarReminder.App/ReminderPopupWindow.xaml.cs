using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using System.Media;
using System.Windows;

namespace CalendarReminder.App;

public partial class ReminderPopupWindow:Window
{
    public ReminderPopupWindow(ScheduleItem item){InitializeComponent();ReminderTimeText.Text=$"提醒时间：{DateTimeFormat.Format(item.ReminderAt)}";ScheduleTimeText.Text=$"日程时间：{DateTimeFormat.Format(item.ScheduledAt)}";ContentText.Text=item.Content;Loaded+=(_,_)=>{AppLog.Info($"提醒弹窗已显示 {item.Id}");SystemSounds.Exclamation.Play();Activate();};}
    private void Close_Click(object sender,RoutedEventArgs e)=>Close();
}
