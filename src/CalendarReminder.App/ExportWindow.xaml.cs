using CalendarReminder.Core.Services;
using Microsoft.Win32;
using System.Diagnostics;
using System.Windows;

namespace CalendarReminder.App;

public partial class ExportWindow:Window
{
    private readonly ScheduleRepository _repository; private string? _folder; public string? SuccessMessage{get;private set;}
    private static readonly (string name,ExportFormat format,string ext,string filter)[] Formats=[("Excel 工作簿 (.xlsx)",ExportFormat.Excel,".xlsx","Excel 工作簿|*.xlsx"),("Word 文档 (.docx)",ExportFormat.Word,".docx","Word 文档|*.docx"),("PDF 文档 (.pdf)",ExportFormat.PDF,".pdf","PDF 文档|*.pdf"),("PowerPoint 演示文稿 (.pptx)",ExportFormat.PowerPoint,".pptx","PowerPoint 演示文稿|*.pptx"),("PNG 图片 (.png)",ExportFormat.PNG,".png","PNG 图片|*.png"),("JPG 图片 (.jpg)",ExportFormat.JPG,".jpg","JPG 图片|*.jpg")];
    public ExportWindow(ScheduleRepository repository){InitializeComponent();_repository=repository;var nums=Enumerable.Range(0,60).Select(x=>x.ToString("00")).ToArray();StartHour.ItemsSource=nums.Take(24);EndHour.ItemsSource=nums.Take(24);StartMinute.ItemsSource=nums;EndMinute.ItemsSource=nums;StartDate.SelectedDate=DateTime.Today;EndDate.SelectedDate=DateTime.Today.AddMonths(1);StartHour.SelectedIndex=0;StartMinute.SelectedIndex=0;EndHour.SelectedIndex=23;EndMinute.SelectedIndex=59;FormatInput.ItemsSource=Formats.Select(x=>x.name);FormatInput.SelectedIndex=0;}
    private bool TryRange(out DateTime start,out DateTime end){start=end=default;if(StartDate.SelectedDate is null||EndDate.SelectedDate is null||StartHour.SelectedIndex<0||StartMinute.SelectedIndex<0||EndHour.SelectedIndex<0||EndMinute.SelectedIndex<0){MessageBox.Show("开始和结束日期时间不能为空。","日历提醒");return false;}start=StartDate.SelectedDate.Value.Date.AddHours(StartHour.SelectedIndex).AddMinutes(StartMinute.SelectedIndex);end=EndDate.SelectedDate.Value.Date.AddHours(EndHour.SelectedIndex).AddMinutes(EndMinute.SelectedIndex);if(end<start){MessageBox.Show("结束时间不得早于开始时间。","日历提醒");return false;}return true;}
    private async void Export_Click(object sender,RoutedEventArgs e){if(!TryRange(out var start,out var end))return;var f=Formats[FormatInput.SelectedIndex];var picker=new SaveFileDialog{Title="保存导出文件",FileName="日历安排"+f.ext,DefaultExt=f.ext,Filter=f.filter,AddExtension=true};if(picker.ShowDialog()!=true)return;try{ExportButton.IsEnabled=false;ExportButton.Content="正在导出…";var all=await _repository.GetRangeAsync(start,end);var result=await new ExportService().ExportAsync(all,start,end,f.format,picker.FileName);_folder=Path.GetDirectoryName(result.Files[0]);SuccessMessage=$"导出成功：{result.ItemCount} 条日程，生成 {result.Files.Count} 个文件。";ResultText.Text=SuccessMessage+"\n"+string.Join("\n",result.Files);ResultPanel.Visibility=Visibility.Visible;}catch(Exception ex){MessageBox.Show(ex.Message,"导出失败",MessageBoxButton.OK,MessageBoxImage.Error);}finally{ExportButton.IsEnabled=true;ExportButton.Content="选择位置并导出";}}
    private void OpenFolder_Click(object sender,RoutedEventArgs e){if(_folder is not null)Process.Start(new ProcessStartInfo("explorer.exe",$"\"{_folder}\""){UseShellExecute=true});}
    private void Cancel_Click(object sender,RoutedEventArgs e)=>Close();
}
