using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.ObjectModel;

namespace CalendarReminder.App.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ScheduleRepository _repository;
    public ObservableCollection<ScheduleItem> Items { get; } = [];
    [ObservableProperty] private bool isBusy;
    public MainViewModel(ScheduleRepository repository) => _repository = repository;
    public async Task LoadAsync() { IsBusy=true; try { var items=await _repository.GetAllAsync(); Items.Clear(); foreach(var item in items) Items.Add(item); } finally { IsBusy=false; } }
}
