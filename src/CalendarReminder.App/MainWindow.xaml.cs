using CalendarReminder.App.ViewModels;
using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace CalendarReminder.App;

public partial class MainWindow : Window
{
    private readonly ScheduleRepository _repository;
    private readonly MainViewModel _viewModel;
    private readonly Forms.NotifyIcon _tray;
    private readonly System.Drawing.Icon? _trayIcon;
    private readonly DispatcherTimer _reminderTimer = new() { Interval=TimeSpan.FromSeconds(25) };
    private readonly AiApiSettingsService _apiSettingsService = new();
    private readonly WindowsCredentialStore _credentialStore = new();
    private readonly ReminderCheckpointStore _reminderCheckpoint = new();
    private readonly DateTime? _previousReminderCheck;
    private AiApiSettings _apiSettings = new();
    private bool _realExit, _settingStartup, _settingApi, _checkingReminders;

    public MainWindow(ScheduleRepository repository)
    {
        InitializeComponent(); _repository=repository; _previousReminderCheck=_reminderCheckpoint.Read(); _viewModel=new(repository); DataContext=_viewModel; ScheduleList.ItemsSource=_viewModel.Items;
        DataPathText.Text=$"数据库：{AppPaths.DatabasePath}\n日志：{AppPaths.LogDirectory}";
        _trayIcon=LoadApplicationIcon();
        _tray=new Forms.NotifyIcon { Icon=_trayIcon??System.Drawing.SystemIcons.Application, Text="日历提醒", Visible=true };
        var menu=new Forms.ContextMenuStrip(); menu.Items.Add("打开日历提醒",null,(_,_)=>Dispatcher.Invoke(ShowFromTray)); menu.Items.Add("退出",null,(_,_)=>Dispatcher.Invoke(ExitApplication)); _tray.ContextMenuStrip=menu; _tray.DoubleClick+=(_,_)=>Dispatcher.Invoke(ShowFromTray);
        _reminderTimer.Tick+=async (_,_)=>await RunReminderCycleAsync();
        Loaded+=async (_,_)=>
        {
            await RefreshAsync(); ReadStartup(); LoadApiSettings();
            await RunStartupReminderCycleAsync(); _reminderTimer.Start();
        };
    }

    private async Task RefreshAsync(string? feedback=null) { await _viewModel.LoadAsync(); CountText.Text=$"共 {_viewModel.Items.Count} 条日程"; EmptyText.Visibility=_viewModel.Items.Count==0?Visibility.Visible:Visibility.Collapsed; ScheduleList.Visibility=_viewModel.Items.Count==0?Visibility.Collapsed:Visibility.Visible; UpdateSelectionUi(); if(feedback is not null) FeedbackText.Text=feedback; }
    private async void NewSchedule_Click(object sender,RoutedEventArgs e) { var dialog=new ScheduleEditorWindow(null){Owner=this}; if(dialog.ShowDialog()==true){ await _repository.AddAsync(dialog.Result!); await RefreshAsync("日程已保存。"); } }
    private async void ScheduleList_MouseDoubleClick(object sender,MouseButtonEventArgs e) { if(ScheduleList.SelectedItem is not ScheduleItem item)return; var dialog=new ScheduleEditorWindow(item){Owner=this}; if(dialog.ShowDialog()==true){ if(dialog.WasDeleted){await _repository.DeleteAsync(item.Id);await RefreshAsync("日程已删除。");}else{await _repository.UpdateAsync(dialog.Result!);await RefreshAsync("日程已更新。");} } }
    private async void ScheduleDelete_Click(object sender,RoutedEventArgs e){e.Handled=true;if(sender is not System.Windows.Controls.Button{Tag:ScheduleItem item})return;if(MessageBox.Show("确定删除这条日程吗？删除后无法恢复。","确认删除",MessageBoxButton.OKCancel,MessageBoxImage.Warning)!=MessageBoxResult.OK)return;await _repository.DeleteAsync(item.Id);await RefreshAsync("日程已删除。");}
    private void ScheduleList_SelectionChanged(object sender,System.Windows.Controls.SelectionChangedEventArgs e)=>UpdateSelectionUi();
    private void UpdateSelectionUi()
    {
        if(SelectedCountText is null||BatchDeleteButton is null||SelectAllButton is null||_viewModel is null)return;var selected=ScheduleList.SelectedItems.Count;SelectedCountText.Text=$" · 已选 {selected} 条";
        BatchDeleteButton.IsEnabled=selected>0;SelectAllButton.IsEnabled=_viewModel.Items.Count>0;SelectAllButton.Content=_viewModel.Items.Count>0&&selected==_viewModel.Items.Count?"取消全选":"全选";
    }
    private void SelectAll_Click(object sender,RoutedEventArgs e){if(_viewModel.Items.Count==0)return;if(ScheduleList.SelectedItems.Count==_viewModel.Items.Count)ScheduleList.UnselectAll();else ScheduleList.SelectAll();UpdateSelectionUi();}
    private async void BatchDelete_Click(object sender,RoutedEventArgs e)
    {
        var selected=ScheduleList.SelectedItems.Cast<ScheduleItem>().ToList();if(selected.Count==0){MessageBox.Show("请先勾选需要删除的日程。","批量删除",MessageBoxButton.OK,MessageBoxImage.Information);return;}
        if(MessageBox.Show($"确定删除选中的 {selected.Count} 条日程吗？删除后无法恢复。","确认批量删除",MessageBoxButton.OKCancel,MessageBoxImage.Warning)!=MessageBoxResult.OK)return;
        try{BatchDeleteButton.IsEnabled=false;SelectAllButton.IsEnabled=false;var deleted=await _repository.DeleteManyAsync(selected.Select(x=>x.Id));await RefreshAsync($"已批量删除 {deleted} 条日程。");}
        catch(Exception ex){AppLog.Error("批量删除日程",ex);MessageBox.Show("批量删除失败，数据库未发生部分删除。详细信息已写入日志。","删除失败",MessageBoxButton.OK,MessageBoxImage.Error);UpdateSelectionUi();}
    }
    private void HomeNav_Click(object sender,RoutedEventArgs e)=>ShowPage(HomePage,HomeNav);
    private void ProfileNav_Click(object sender,RoutedEventArgs e)=>ShowPage(ProfilePage,ProfileNav);
    private void ApiNav_Click(object sender,RoutedEventArgs e)=>ShowPage(ApiPage,ApiNav);
    private void ShowPage(UIElement page,System.Windows.Controls.Button nav)
    {
        HomePage.Visibility=ProfilePage.Visibility=ApiPage.Visibility=Visibility.Collapsed;page.Visibility=Visibility.Visible;
        HomeNav.Background=ProfileNav.Background=ApiNav.Background=System.Windows.Media.Brushes.Transparent;nav.Background=(System.Windows.Media.Brush)FindResource("Primary");
    }
    private async void Export_Click(object sender,RoutedEventArgs e){var d=new ExportWindow(_repository){Owner=this};d.ShowDialog();await RefreshAsync(d.SuccessMessage);}
    private async void Import_Click(object sender,RoutedEventArgs e)
    {
        var picker=new OpenFileDialog{Title="选择要导入的日历文件",Filter="支持的文件|*.xls;*.xlsx;*.docx;*.pdf;*.pptx;*.png;*.jpg;*.jpeg"}; if(picker.ShowDialog()!=true)return;
        try{FeedbackText.Text="正在解析文件，请稍候…";var progress=new Progress<int>(x=>FeedbackText.Text=$"正在解析文件… {x}%");var rows=await new ImportService().ParseAsync(picker.FileName,progress);await _repository.MarkDuplicatesAsync(rows);if(rows.Count==0){MessageBox.Show("未识别到可预览的日程。","日历提醒",MessageBoxButton.OK,MessageBoxImage.Information);return;}var preview=new ImportPreviewWindow(_repository,rows){Owner=this};if(preview.ShowDialog()==true)await RefreshAsync($"导入完成：新增 {preview.Result?.Imported??0} 条，跳过 {preview.Result?.Skipped??0} 条。");else FeedbackText.Text="已取消导入，数据库未发生变化。";}
        catch(NotSupportedException ex){MessageBox.Show(ex.Message,"导入不可用",MessageBoxButton.OK,MessageBoxImage.Warning);}catch(Exception ex){AppLog.Error("导入",ex);MessageBox.Show(ex.Message,"导入失败",MessageBoxButton.OK,MessageBoxImage.Error);}
    }

    private async Task RunStartupReminderCycleAsync()
    {
        if(_checkingReminders)return;_checkingReminders=true;var checkedThrough=DateTime.Now;
        try
        {
            if(StartupCheck.IsChecked==true&&_previousReminderCheck.HasValue&&_previousReminderCheck.Value<checkedThrough)
            {
                var missed=await _repository.GetDueRangeAsync(_previousReminderCheck.Value,checkedThrough);
                if(missed.Count>0)
                {
                    var popup=new MissedRemindersWindow(missed,_previousReminderCheck.Value,checkedThrough){Owner=this};popup.ShowDialog();
                    await _repository.MarkTriggeredAsync(missed.Select(x=>x.Id));
                }
            }
            await CheckRemindersAsync(checkedThrough);_reminderCheckpoint.Write(checkedThrough);
        }
        catch(Exception ex){AppLog.Error("启动提醒检查",ex);}
        finally{_checkingReminders=false;}
    }

    private async Task RunReminderCycleAsync()
    {
        if(_checkingReminders)return;_checkingReminders=true;var checkedThrough=DateTime.Now;
        try{await CheckRemindersAsync(checkedThrough);_reminderCheckpoint.Write(checkedThrough);}
        catch(Exception ex){AppLog.Error("检查提醒",ex);}
        finally{_checkingReminders=false;}
    }

    private async Task CheckRemindersAsync(DateTime checkedThrough)
    {
        var due=await _repository.GetDueAsync(checkedThrough);
        foreach(var item in due)
        {
            var content=item.Content.Length>180?item.Content[..180]+"…":item.Content;AppLog.Info($"触发提醒 {item.Id}，提醒时间 {DateTimeFormat.Format(item.ReminderAt)}");
            _tray.ShowBalloonTip(10000,"日程提醒",$"提醒时间：{DateTimeFormat.Format(item.ReminderAt)}\n日程时间：{DateTimeFormat.Format(item.ScheduledAt)}\n{content}",Forms.ToolTipIcon.Info);
            var popup=new ReminderPopupWindow(item);popup.ShowDialog();await _repository.MarkTriggeredAsync(item.Id);await Task.Delay(300);
        }
    }

    protected override void OnClosing(CancelEventArgs e){if(!_realExit){e.Cancel=true;Hide();_tray.ShowBalloonTip(2500,"日历提醒","程序已最小化到托盘，将继续检查日程。",Forms.ToolTipIcon.Info);}base.OnClosing(e);}
    private void ShowFromTray(){Show();WindowState=WindowState.Normal;Activate();}
    private void ExitApplication(){_realExit=true;_reminderTimer.Stop();_tray.Visible=false;_tray.Dispose();_trayIcon?.Dispose();Close();Application.Current.Shutdown();}
    internal void ExitForTest()=>ExitApplication();

    private static System.Drawing.Icon? LoadApplicationIcon()
    {
        try
        {
            var path=Environment.ProcessPath;
            return string.IsNullOrWhiteSpace(path)?null:System.Drawing.Icon.ExtractAssociatedIcon(path);
        }
        catch(Exception ex)
        {
            AppLog.Error("读取应用图标",ex);
            return null;
        }
    }

    private void ReadStartup(){_settingStartup=true;try{using var key=Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run");StartupCheck.IsChecked=key?.GetValue("CalendarReminder") is string;}finally{_settingStartup=false;}}
    private void StartupCheck_Changed(object sender,RoutedEventArgs e)
    {
        if(_settingStartup)return;try{using var key=Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run",true);if(StartupCheck.IsChecked==true)key.SetValue("CalendarReminder",$"\"{Environment.ProcessPath}\"");else key.DeleteValue("CalendarReminder",false);}
        catch(Exception ex){AppLog.Error("设置开机启动",ex);MessageBox.Show("无法修改开机启动设置，请检查当前用户注册表权限。","日历提醒",MessageBoxButton.OK,MessageBoxImage.Error);ReadStartup();}
    }

    private void LoadApiSettings()
    {
        _settingApi=true;
        try
        {
            _apiSettings=_apiSettingsService.Load();ProviderInput.ItemsSource=AiProviderCatalog.All;
            ProviderInput.SelectedItem=AiProviderCatalog.Get(_apiSettings.Provider);BindModels(_apiSettings.Model);UpdateApiKeyStatus();
        }
        catch(Exception ex){AppLog.Error("载入 API 设置",ex);ApiKeyStatusText.Text="无法读取 API 设置。";}
        finally{_settingApi=false;}
    }

    private void BindModels(string? preferred=null)
    {
        if(ProviderInput.SelectedItem is not AiProviderDefinition provider)return;
        ModelInput.ItemsSource=provider.Models;ModelInput.SelectedItem=provider.Models.FirstOrDefault(x=>x.Id==preferred)??provider.Models[0];
    }

    private void ProviderInput_SelectionChanged(object sender,System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if(_settingApi)return;BindModels();UpdateApiKeyStatus();
    }

    private void UpdateApiKeyStatus()
    {
        try
        {
            if(ProviderInput.SelectedItem is not AiProviderDefinition provider)return;
            ApiKeyStatusText.Text=_credentialStore.Read(provider.Provider) is null?"尚未保存此供应商的 API Key。":"API Key 已由 Windows 凭据管理器加密保存。";
        }
        catch(Exception ex){AppLog.Error("检查 API Key",ex);ApiKeyStatusText.Text="无法读取 Windows 凭据。";}
    }

    private (AiProviderDefinition Provider,AiModelOption Model) SelectedApi()
    {
        if(ProviderInput.SelectedItem is not AiProviderDefinition provider||ModelInput.SelectedItem is not AiModelOption model)throw new InvalidOperationException("请选择供应商和模型。");
        return(provider,model);
    }

    private void SaveApi_Click(object sender,RoutedEventArgs e)
    {
        try
        {
            var selected=SelectedApi();_apiSettings=new(){Provider=selected.Provider.Provider,Model=selected.Model.Id};_apiSettingsService.Save(_apiSettings);
            if(!string.IsNullOrWhiteSpace(ApiKeyInput.Password)){_credentialStore.Save(selected.Provider.Provider,ApiKeyInput.Password);ApiKeyInput.Clear();}
            UpdateApiKeyStatus();MessageBox.Show("API 配置已保存。密钥存放在当前用户的 Windows 凭据管理器中。","日历提醒",MessageBoxButton.OK,MessageBoxImage.Information);
        }
        catch(Exception ex){AppLog.Error("保存 API 配置",ex);MessageBox.Show(ex.Message,"保存失败",MessageBoxButton.OK,MessageBoxImage.Error);}
    }

    private void DeleteApiKey_Click(object sender,RoutedEventArgs e)
    {
        try
        {
            var selected=SelectedApi();if(MessageBox.Show($"确定删除 {selected.Provider.DisplayName} 的已保存 API Key 吗？","删除密钥",MessageBoxButton.OKCancel,MessageBoxImage.Warning)!=MessageBoxResult.OK)return;
            _credentialStore.Delete(selected.Provider.Provider);ApiKeyInput.Clear();UpdateApiKeyStatus();
        }
        catch(Exception ex){AppLog.Error("删除 API Key",ex);MessageBox.Show(ex.Message,"删除失败",MessageBoxButton.OK,MessageBoxImage.Error);}
    }

    private async void TestApi_Click(object sender,RoutedEventArgs e)
    {
        try
        {
            var selected=SelectedApi();var key=string.IsNullOrWhiteSpace(ApiKeyInput.Password)?_credentialStore.Read(selected.Provider.Provider):ApiKeyInput.Password;
            if(string.IsNullOrWhiteSpace(key))throw new InvalidOperationException("请先输入或保存 API Key。");
            AiProgressText.Text="正在测试连接…";await new AiImportService().TestConnectionAsync(selected.Provider.Provider,selected.Model.Id,key);AiProgressText.Text="连接测试成功。";
        }
        catch(Exception ex){AppLog.Error("测试 API 连接",ex);AiProgressText.Text="连接测试失败。";MessageBox.Show(ex.Message,"连接失败",MessageBoxButton.OK,MessageBoxImage.Error);}
    }

    private async void AiImport_Click(object sender,RoutedEventArgs e)
    {
        var picker=new OpenFileDialog{Title="选择要用 AI 识别的日历文件",Filter="支持的文件|*.xls;*.xlsx;*.docx;*.pdf;*.pptx;*.png;*.jpg;*.jpeg"};if(picker.ShowDialog()!=true)return;
        try
        {
            var selected=SelectedApi();var key=string.IsNullOrWhiteSpace(ApiKeyInput.Password)?_credentialStore.Read(selected.Provider.Provider):ApiKeyInput.Password;
            if(string.IsNullOrWhiteSpace(key))throw new InvalidOperationException("请先输入或保存 API Key。");
            var existing=await _repository.GetAllAsync();
            var consent=$"将向 {selected.Provider.DisplayName}（{selected.Model.DisplayName}）发送：\n\n• 所选文件中抽取的文字\n• 全部 {existing.Count} 条已有日程的时间、提醒时间和安排描述\n\n供应商可能记录数据并收取 API 费用。是否继续？";
            if(MessageBox.Show(consent,"发送数据前确认",MessageBoxButton.OKCancel,MessageBoxImage.Warning)!=MessageBoxResult.OK){AiProgressText.Text="已取消，未向供应商发送数据。";return;}
            _apiSettings=new(){Provider=selected.Provider.Provider,Model=selected.Model.Id};_apiSettingsService.Save(_apiSettings);AiImportButton.IsEnabled=false;
            var progress=new Progress<string>(message=>AiProgressText.Text=message);
            var result=await new AiImportService().AnalyzeAsync(picker.FileName,selected.Provider.Provider,selected.Model.Id,key,existing,progress);
            await _repository.MarkDuplicatesAsync(result.Candidates);
            var preview=new ImportPreviewWindow(_repository,result.Candidates){Owner=this};
            AiProgressText.Text=$"已生成规范 Excel：{result.ExcelPath}\n已检查 {result.ExistingScheduleCount} 条已有日程，拦截 {result.SemanticOverlapCount} 条高度重合项。";
            if(preview.ShowDialog()==true)await RefreshAsync($"AI 导入完成：新增 {preview.Result?.Imported??0} 条，跳过 {preview.Result?.Skipped??0} 条。规范文件：{result.ExcelPath}");
        }
        catch(NotSupportedException ex){MessageBox.Show(ex.Message,"AI 导入不可用",MessageBoxButton.OK,MessageBoxImage.Warning);}
        catch(Exception ex){AppLog.Error("AI 导入",ex);MessageBox.Show(ex.Message,"AI 导入失败",MessageBoxButton.OK,MessageBoxImage.Error);AiProgressText.Text="AI 导入失败，数据库未发生变化。";}
        finally{AiImportButton.IsEnabled=true;}
    }

    private void OpenAiFolder_Click(object sender,RoutedEventArgs e)
    {
        try{AppPaths.Ensure();Process.Start(new ProcessStartInfo("explorer.exe",AppPaths.AiImportDirectory){UseShellExecute=true});}
        catch(Exception ex){AppLog.Error("打开 AI 文件目录",ex);MessageBox.Show("无法打开目录。","日历提醒",MessageBoxButton.OK,MessageBoxImage.Error);}
    }
}
