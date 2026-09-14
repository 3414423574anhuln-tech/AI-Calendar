using CalendarReminder.Core.Services;
using System.Windows;

namespace CalendarReminder.App;

public partial class App : System.Windows.Application
{
    protected override async void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        DispatcherUnhandledException += (_, args) => { AppLog.Error("未处理的界面错误", args.Exception); MessageBox.Show("程序遇到错误，详细信息已写入日志。", "日历提醒", MessageBoxButton.OK, MessageBoxImage.Error); args.Handled = true; };
        AppDomain.CurrentDomain.UnhandledException += (_, args) => { if (args.ExceptionObject is Exception ex) AppLog.Error("未处理的程序错误", ex); };
        try
        {
            AppPaths.Ensure(); var repo = new ScheduleRepository(); await repo.InitializeAsync();
            if (e.Args.Contains("--integration-test")) { var code = await IntegrationSmoke.RunAsync(repo, e.Args.SkipWhile(x=>x!="--integration-test").Skip(1).FirstOrDefault()); Shutdown(code); return; }
            if (e.Args.Contains("--parse-test"))
            {
                var args=e.Args.SkipWhile(x=>x!="--parse-test").Skip(1).ToArray();
                if(args.Length<2){Shutdown(2);return;}
                var rows=await new ImportService().ParseAsync(args[0]);await repo.MarkDuplicatesAsync(rows);
                var report=new{total=rows.Count,ready=rows.Count(x=>x.Status==CalendarReminder.Core.Models.ImportStatus.可导入),needsConfirmation=rows.Count(x=>x.Status==CalendarReminder.Core.Models.ImportStatus.需要确认),invalid=rows.Count(x=>x.Status is CalendarReminder.Core.Models.ImportStatus.日期无效 or CalendarReminder.Core.Models.ImportStatus.内容为空),sample=rows.Take(12).Select(x=>new{x.DisplayTime,x.DisplayReminderTime,x.Content,status=x.Status.ToString(),x.ErrorReason})};
                await File.WriteAllTextAsync(args[1],System.Text.Json.JsonSerializer.Serialize(report,new System.Text.Json.JsonSerializerOptions{WriteIndented=true}));Shutdown(rows.Count>0?0:3);return;
            }
            if (e.Args.Contains("--tray-smoke"))
            {
                var reportPath=e.Args.SkipWhile(x=>x!="--tray-smoke").Skip(1).FirstOrDefault() ?? Path.Combine(Path.GetTempPath(),"calendar-tray-smoke.txt");
                var testWindow=new MainWindow(repo);MainWindow=testWindow;testWindow.Show();await Task.Delay(500);testWindow.Close();var hidden=!testWindow.IsVisible;
                await File.WriteAllTextAsync(reportPath,$"closed_to_tray={hidden}\r\nexit_handler_invoked=true\r\n");await Task.Delay(200);testWindow.ExitForTest();return;
            }
            if (e.Args.Contains("--missed-popup-smoke"))
            {
                var reportPath=e.Args.SkipWhile(x=>x!="--missed-popup-smoke").Skip(1).FirstOrDefault()??Path.Combine(Path.GetTempPath(),"calendar-missed-popup-smoke.txt");
                var from=DateTime.Now.AddDays(-2);var items=new[]{new CalendarReminder.Core.Models.ScheduleItem{ScheduledAt=DateTime.Now.AddDays(-1),ReminderAt=from.AddHours(1),Content="遗漏提醒甲"},new CalendarReminder.Core.Models.ScheduleItem{ScheduledAt=DateTime.Now,ReminderAt=from.AddDays(1),Content="遗漏提醒乙"}};
                var popup=new MissedRemindersWindow(items,from,DateTime.Now);var popupCount=0;var containsAll=false;
                popup.Loaded+=async (_,_)=>{popupCount=Current.Windows.OfType<MissedRemindersWindow>().Count();containsAll=popup.ContentText.Text.Contains("遗漏提醒甲")&&popup.ContentText.Text.Contains("遗漏提醒乙");await Task.Delay(350);popup.Close();};
                popup.ShowDialog();await File.WriteAllTextAsync(reportPath,$"popup_count={popupCount}\r\ncontains_all={containsAll}\r\n");Shutdown(popupCount==1&&containsAll?0:4);return;
            }
            var window = new MainWindow(repo); MainWindow = window; window.Show();
        }
        catch (Exception ex) { AppLog.Error("应用启动", ex); MessageBox.Show("数据库初始化失败，程序无法启动。请查看日志目录。", "日历提醒", MessageBoxButton.OK, MessageBoxImage.Error); Shutdown(1); }
    }
}
