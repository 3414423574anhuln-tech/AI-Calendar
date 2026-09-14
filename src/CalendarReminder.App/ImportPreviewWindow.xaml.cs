using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using System.Collections.ObjectModel;
using System.Windows;

namespace CalendarReminder.App;

public partial class ImportPreviewWindow:Window
{
    private readonly ScheduleRepository _repository;private readonly ObservableCollection<ImportCandidate> _rows;public ImportResult? Result{get;private set;}
    public ImportPreviewWindow(ScheduleRepository repository,IEnumerable<ImportCandidate> rows){InitializeComponent();_repository=repository;_rows=new(rows);Grid.ItemsSource=_rows;DuplicateActionInput.ItemsSource=Enum.GetValues<DuplicateAction>();DuplicateActionInput.SelectedItem=DuplicateAction.跳过重复项;}
    private void DeleteRow_Click(object sender,RoutedEventArgs e){if(Grid.SelectedItem is ImportCandidate row)_rows.Remove(row);}
    private async void Import_Click(object sender,RoutedEventArgs e){Grid.CommitEdit();Grid.CommitEdit();foreach(var row in _rows){row.Content=row.Content?.Trim()??"";if(row.ScheduledAt is null){row.Status=ImportStatus.日期无效;row.ErrorReason="日期无效。";}else if(string.IsNullOrWhiteSpace(row.Content)){row.Status=ImportStatus.内容为空;row.ErrorReason="安排内容为空。";}else if(row.Status is ImportStatus.日期无效 or ImportStatus.内容为空 or ImportStatus.需要确认){row.Status=row.IsDuplicate?ImportStatus.疑似重复:ImportStatus.可导入;row.ErrorReason="";}}
        if(_rows.Any(x=>x.IsSelected&&(x.ScheduledAt is null||string.IsNullOrWhiteSpace(x.Content)))){Grid.Items.Refresh();MessageBox.Show("仍有无效日期或空内容，请修正或取消勾选后再导入。","无法导入",MessageBoxButton.OK,MessageBoxImage.Warning);return;}try{ImportButton.IsEnabled=false;Result=await _repository.BatchImportAsync(_rows,(DuplicateAction)DuplicateActionInput.SelectedItem);DialogResult=true;}catch(Exception ex){MessageBox.Show("批量导入失败，所有更改均已回滚。\n"+ex.Message,"导入失败",MessageBoxButton.OK,MessageBoxImage.Error);}finally{ImportButton.IsEnabled=true;}}
    private void Cancel_Click(object sender,RoutedEventArgs e)=>DialogResult=false;
}
