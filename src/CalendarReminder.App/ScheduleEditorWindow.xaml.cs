using CalendarReminder.Core.Models;
using System.Windows;

namespace CalendarReminder.App;

public partial class ScheduleEditorWindow:Window
{
    private readonly ScheduleItem? _original; public ScheduleItem? Result{get;private set;} public bool WasDeleted{get;private set;}
    public ScheduleEditorWindow(ScheduleItem? item){InitializeComponent();_original=item;var hours=Enumerable.Range(0,24).Select(x=>x.ToString("00")).ToArray();var minutes=Enumerable.Range(0,60).Select(x=>x.ToString("00")).ToArray();HourInput.ItemsSource=hours;MinuteInput.ItemsSource=minutes;ReminderHourInput.ItemsSource=hours;ReminderMinuteInput.ItemsSource=minutes;var at=item?.ScheduledAt??DateTime.Now;var reminder=item?.ReminderAt is { } saved&&saved!=default?saved:at;DateInput.SelectedDate=at.Date;HourInput.SelectedIndex=at.Hour;MinuteInput.SelectedIndex=at.Minute;ReminderDateInput.SelectedDate=reminder.Date;ReminderHourInput.SelectedIndex=reminder.Hour;ReminderMinuteInput.SelectedIndex=reminder.Minute;if(item!=null){Heading.Text="查看和修改日程";DeleteButton.Visibility=Visibility.Visible;ContentInput.Text=item.Content;}ContentInput.Focus();}
    private bool Validate(out DateTime at,out DateTime reminder){at=reminder=default;if(DateInput.SelectedDate is null){MessageBox.Show("请选择日程日期。","输入有误");return false;}if(HourInput.SelectedIndex<0||MinuteInput.SelectedIndex<0){MessageBox.Show("请选择有效的日程时间。","输入有误");return false;}if(ReminderDateInput.SelectedDate is null){MessageBox.Show("请选择提醒日期。","输入有误");return false;}if(ReminderHourInput.SelectedIndex<0||ReminderMinuteInput.SelectedIndex<0){MessageBox.Show("请选择有效的提醒时间。","输入有误");return false;}if(string.IsNullOrWhiteSpace(ContentInput.Text)){MessageBox.Show("安排内容不能为空。","输入有误");return false;}at=DateInput.SelectedDate.Value.Date.AddHours(HourInput.SelectedIndex).AddMinutes(MinuteInput.SelectedIndex);reminder=ReminderDateInput.SelectedDate.Value.Date.AddHours(ReminderHourInput.SelectedIndex).AddMinutes(ReminderMinuteInput.SelectedIndex);return true;}
    private void Confirm_Click(object sender,RoutedEventArgs e){if(!Validate(out var at,out var reminder))return;Result=new ScheduleItem{Id=_original?.Id??Guid.NewGuid(),ScheduledAt=at,ReminderAt=reminder,Content=ContentInput.Text.Trim(),CreatedAt=_original?.CreatedAt??DateTime.Now,UpdatedAt=DateTime.Now,ReminderTriggered=_original?.ReminderTriggered??false};DialogResult=true;}
    private void Cancel_Click(object sender,RoutedEventArgs e)=>DialogResult=false;
    private void Delete_Click(object sender,RoutedEventArgs e){if(MessageBox.Show("确定删除这条日程吗？删除后无法恢复。","确认删除",MessageBoxButton.OKCancel,MessageBoxImage.Warning)==MessageBoxResult.OK){WasDeleted=true;Result=_original;DialogResult=true;}}
}
