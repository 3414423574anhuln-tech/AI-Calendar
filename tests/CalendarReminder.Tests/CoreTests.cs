using CalendarReminder.Core;
using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;
using Xunit;

namespace CalendarReminder.Tests;

public sealed class CoreTests : IAsyncLifetime
{
    private readonly string _dir=Path.Combine(Path.GetTempPath(),"CalendarReminderTests",Guid.NewGuid().ToString("N"));
    private ScheduleRepository _repo=null!;
    public async Task InitializeAsync(){Directory.CreateDirectory(_dir);_repo=new ScheduleRepository(Path.Combine(_dir,"test.db"));await _repo.InitializeAsync();}
    public Task DisposeAsync(){try{Directory.Delete(_dir,true);}catch{}return Task.CompletedTask;}

    [Fact] public void FormatsRequiredDateTime(){Assert.Equal("2026/08/02 21:07",DateTimeFormat.Format(new DateTime(2026,8,2,21,7,44)));}
    [Fact] public void Parses24HourTime(){Assert.True(DateTimeFormat.TryParse("2026-8-2 21:07",out var value));Assert.Equal(21,value.Hour);Assert.False(DateTimeFormat.TryParse("2026-8-2 25:07",out _));}
    [Fact] public void ParsesCommonChineseDate(){Assert.True(DateTimeFormat.TryParse("2026年08月02日 21点07分",out var value));Assert.Equal(new DateTime(2026,8,2,21,7,0),value);}
    [Theory][InlineData("2026/02/30 10:00")][InlineData("2026/08/02")][InlineData("8月2日 10:00")][InlineData("")]
    public void RejectsInvalidOrIncompleteDate(string text)=>Assert.False(DateTimeFormat.TryParse(text,out _));

    [Fact] public async Task RangeIncludesBothBoundaries(){var start=new DateTime(2026,1,1,9,0,0);await _repo.AddAsync(new(){ScheduledAt=start,Content="A"});await _repo.AddAsync(new(){ScheduledAt=start.AddHours(1),Content="B"});var rows=await _repo.GetRangeAsync(start,start.AddHours(1));Assert.Equal(2,rows.Count);}
    [Fact] public async Task SchedulesSortAscending(){await _repo.AddAsync(new(){ScheduledAt=new(2026,1,2),Content="later"});await _repo.AddAsync(new(){ScheduledAt=new(2026,1,1),Content="early"});Assert.Equal(["early","later"],(await _repo.GetAllAsync()).Select(x=>x.Content));}
    [Fact] public async Task DetectsTrimmedExactDuplicate(){var at=new DateTime(2026,2,3,4,5,0);await _repo.AddAsync(new(){ScheduledAt=at,Content=" 同一事项 "});var rows=new[]{new ImportCandidate{ScheduledAt=at,Content="同一事项"}};await _repo.MarkDuplicatesAsync(rows);Assert.True(rows[0].IsDuplicate);}
    [Fact] public async Task DatabaseAddWorks(){var item=new ScheduleItem{ScheduledAt=DateTime.Now,Content="新增"};await _repo.AddAsync(item);Assert.Contains(await _repo.GetAllAsync(),x=>x.Id==item.Id);}
    [Fact] public async Task DatabaseUpdateWorks(){var item=new ScheduleItem{ScheduledAt=DateTime.Now,Content="原始"};await _repo.AddAsync(item);item.Content="修改";await _repo.UpdateAsync(item);Assert.Equal("修改",(await _repo.GetAllAsync()).Single().Content);}
    [Fact] public async Task DatabaseDeleteWorks(){var item=new ScheduleItem{ScheduledAt=DateTime.Now,Content="删除"};await _repo.AddAsync(item);await _repo.DeleteAsync(item.Id);Assert.Empty(await _repo.GetAllAsync());}
    [Fact] public async Task BatchDeleteRemovesOnlySelectedSchedules(){var a=new ScheduleItem{ScheduledAt=DateTime.Now,Content="批量甲"};var b=new ScheduleItem{ScheduledAt=DateTime.Now.AddMinutes(1),Content="保留"};var c=new ScheduleItem{ScheduledAt=DateTime.Now.AddMinutes(2),Content="批量乙"};await _repo.AddAsync(a);await _repo.AddAsync(b);await _repo.AddAsync(c);var deleted=await _repo.DeleteManyAsync(new[]{a.Id,c.Id,c.Id});Assert.Equal(2,deleted);var remaining=Assert.Single(await _repo.GetAllAsync());Assert.Equal(b.Id,remaining.Id);}
    [Fact] public async Task ChangingDateResetsReminder(){var item=new ScheduleItem{ScheduledAt=DateTime.Now,Content="提醒",ReminderTriggered=true};await _repo.AddAsync(item);item.ScheduledAt=item.ScheduledAt.AddMinutes(1);await _repo.UpdateAsync(item);Assert.False((await _repo.GetAllAsync()).Single().ReminderTriggered);}
    [Fact] public async Task ReminderTimeControlsDueQuery(){var now=DateTime.Now;var item=new ScheduleItem{ScheduledAt=now.AddDays(2),ReminderAt=now.AddMinutes(-1),Content="提前提醒"};await _repo.AddAsync(item);Assert.Contains(await _repo.GetDueAsync(now),x=>x.Id==item.Id);}
    [Fact] public async Task ChangingReminderTimeResetsTriggeredState(){var item=new ScheduleItem{ScheduledAt=DateTime.Now.AddDays(1),ReminderAt=DateTime.Now.AddHours(1),Content="独立提醒",ReminderTriggered=true};await _repo.AddAsync(item);item.ReminderAt=item.ReminderAt.AddMinutes(10);await _repo.UpdateAsync(item);Assert.False((await _repo.GetAllAsync()).Single().ReminderTriggered);}
    [Fact] public async Task LegacyDatabaseMigratesReminderTime()
    {
        var path=Path.Combine(_dir,"legacy.db");var at=new DateTime(2026,8,2,21,7,0);await using(var c=new SqliteConnection($"Data Source={path}")){await c.OpenAsync();var cmd=c.CreateCommand();cmd.CommandText="CREATE TABLE Schedules(Id TEXT PRIMARY KEY,ScheduledAt TEXT NOT NULL,Content TEXT NOT NULL,CreatedAt TEXT NOT NULL,UpdatedAt TEXT NOT NULL,ReminderTriggered INTEGER NOT NULL DEFAULT 0);INSERT INTO Schedules VALUES($id,$at,'旧数据',$at,$at,0)";cmd.Parameters.AddWithValue("$id",Guid.NewGuid().ToString());cmd.Parameters.AddWithValue("$at",at.ToString("O"));await cmd.ExecuteNonQueryAsync();}
        var repo=new ScheduleRepository(path);await repo.InitializeAsync();var loaded=Assert.Single(await repo.GetAllAsync());Assert.Equal(at,loaded.ReminderAt);
    }

    [Fact] public async Task ExcelImportFindsNamedColumns(){var path=Path.Combine(_dir,"input.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("日历");ws.Cell(1,1).Value="日期时间";ws.Cell(1,2).Value="安排";ws.Cell(2,1).Value="2026/08/02 21:07";ws.Cell(2,2).Value="测试安排";wb.SaveAs(path);}var rows=await new ImportService(new FakeOcr()).ParseAsync(path);Assert.Single(rows);Assert.Equal("测试安排",rows[0].Content);Assert.Equal(new DateTime(2026,8,2,21,7,0),rows[0].ScheduledAt);}
    [Fact] public async Task ExcelDateWithoutTimeDefaultsToSevenThirtyAndPreviousDayReminder(){var path=Path.Combine(_dir,"date-only.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("日历");ws.Cell(1,1).Value="日期";ws.Cell(1,2).Value="安排";ws.Cell(2,1).Value=new DateTime(2026,8,2);ws.Cell(2,2).Value="拜访客户";wb.SaveAs(path);}var row=Assert.Single(await new ImportService(new FakeOcr()).ParseAsync(path));Assert.Equal(new DateTime(2026,8,2,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,8,1,7,30,0),row.ReminderAt);Assert.Equal(ImportStatus.可导入,row.Status);}
    [Fact] public void TextDateWithoutTimeUsesImportDefaults(){var row=Assert.Single(ImportService.ParseText("2026年8月2日 拜访客户"));Assert.Equal(new DateTime(2026,8,2,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,8,1,7,30,0),row.ReminderAt);Assert.Equal("拜访客户",row.Content);}
    [Fact] public async Task PlanningGridProducesPreviewCandidates()
    {
        var path=Path.Combine(_dir,"plan.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("26年新品开发计划");ws.Cell(1,1).Value="26年新品开发计划";ws.Cell(3,1).Value="序号";ws.Cell(3,2).Value="客户";ws.Cell(3,3).Value="内容";ws.Cell(3,4).Value="拜访日程（周）";ws.Cell(4,4).Value="7月";ws.Cell(5,4).Value=27;ws.Cell(6,1).Value=1;ws.Cell(6,2).Value="渝江";ws.Cell(6,3).Value="邀请崔总";ws.Cell(6,4).Value="7月22日发邀请";ws.Cell(7,1).Value=2;ws.Cell(7,2).Value="雄邦";ws.Cell(7,3).Value="谈合同";ws.Cell(7,4).Value="8月底拜访";wb.SaveAs(path);}
        var rows=await new ImportService(new FakeOcr()).ParseAsync(path);Assert.Equal(2,rows.Count);var exact=Assert.Single(rows,x=>x.ScheduledAt==new DateTime(2026,7,22,7,30,0));Assert.Equal(new DateTime(2026,7,21,7,30,0),exact.ReminderAt);Assert.Equal(ImportStatus.可导入,exact.Status);var monthEnd=Assert.Single(rows,x=>x.ScheduledAt==new DateTime(2026,8,25,7,30,0));Assert.Equal(new DateTime(2026,8,24,7,30,0),monthEnd.ReminderAt);Assert.Equal(ImportStatus.可导入,monthEnd.Status);
    }
    [Fact] public async Task PlanningDateRangeUsesStartOnlyAndRequiresConfirmation()
    {
        var path=Path.Combine(_dir,"range.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="26年工作计划（7月）";ws.Cell(2,1).Value="序号";ws.Cell(2,2).Value="项目";ws.Cell(2,3).Value="内容";ws.Cell(2,4).Value="说明";ws.Cell(3,1).Value=1;ws.Cell(3,2).Value="中汽研";ws.Cell(3,3).Value="展会";ws.Cell(3,4).Value="9月16~17日参展";wb.SaveAs(path);}
        var rows=await new ImportService(new FakeOcr()).ParseAsync(path);Assert.Equal(2,rows.Count);Assert.Equal([new DateTime(2026,9,16,7,30,0),new DateTime(2026,9,17,7,30,0)],rows.Select(x=>x.ScheduledAt!.Value));Assert.All(rows,x=>{Assert.Equal(new DateTime(2026,9,15,7,30,0),x.ReminderAt);Assert.Equal(ImportStatus.可导入,x.Status);Assert.Contains("日期范围：2026/09/16—2026/09/17",x.Content);});Assert.Single(rows.Select(x=>x.Content).Distinct());
    }
    [Fact] public async Task PlanningRowsWithoutDatesAreNotSilentlyDropped()
    {
        var path=Path.Combine(_dir,"missing-date.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="26年工作计划（7月）";ws.Cell(2,1).Value="序号";ws.Cell(2,2).Value="项目";ws.Cell(2,3).Value="内容";ws.Cell(3,1).Value=1;ws.Cell(3,2).Value="宝涞";ws.Cell(3,3).Value="催款";wb.SaveAs(path);}
        var row=Assert.Single(await new ImportService(new FakeOcr()).ParseAsync(path));Assert.Null(row.ScheduledAt);Assert.Equal(ImportStatus.需要确认,row.Status);
    }
    [Fact] public async Task BatchImportRollsBackOnInvalidRow(){var rows=new[]{new ImportCandidate{ScheduledAt=new DateTime(2026,1,1),Content="有效"},new ImportCandidate{ScheduledAt=null,Content="无日期"}};await Assert.ThrowsAsync<InvalidDataException>(()=>_repo.BatchImportAsync(rows,DuplicateAction.仍然导入));Assert.Empty(await _repo.GetAllAsync());}
    [Fact] public async Task DuplicateCanBeSkipped(){var at=new DateTime(2026,1,1);await _repo.AddAsync(new(){ScheduledAt=at,Content="A"});var result=await _repo.BatchImportAsync([new(){ScheduledAt=at,Content=" A "}],DuplicateAction.跳过重复项);Assert.Equal(1,result.Skipped);Assert.Single(await _repo.GetAllAsync());}
    [Fact] public async Task BatchImportPersistsCandidateReminderTime(){var scheduled=new DateTime(2026,8,2,7,30,0);var reminder=new DateTime(2026,8,1,7,30,0);await _repo.BatchImportAsync([new(){ScheduledAt=scheduled,ReminderAt=reminder,Content="默认提醒"}],DuplicateAction.仍然导入);var item=Assert.Single(await _repo.GetAllAsync());Assert.Equal(scheduled,item.ScheduledAt);Assert.Equal(reminder,item.ReminderAt);}
    [Fact] public void YearOnlyDefaultsToFirstDayAndPreviousYearLastDay(){var row=Assert.Single(ImportService.ParseText("2026年 年度产品规划"));Assert.Equal(new DateTime(2026,1,1,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2025,12,31,7,30,0),row.ReminderAt);Assert.Equal(ImportStatus.可导入,row.Status);}
    [Theory]
    [InlineData("2026年8月 客户走访",8,1,7,31)]
    [InlineData("2026年8月上旬 客户走访",8,1,7,31)]
    [InlineData("2026年八月初 客户走访",8,1,7,31)]
    [InlineData("2026年8月中旬 客户走访",8,13,8,12)]
    [InlineData("2026年8月下旬 客户走访",8,19,8,18)]
    [InlineData("2026年8月底 客户走访",8,25,8,24)]
    [InlineData("2026年8月末 客户走访",8,25,8,24)]
    public void PartialMonthUsesRequestedDefaultDay(string text,int month,int day,int reminderMonth,int reminderDay){var row=Assert.Single(ImportService.ParseText(text));Assert.Equal(new DateTime(2026,month,day,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,reminderMonth,reminderDay,7,30,0),row.ReminderAt);Assert.Equal(ImportStatus.可导入,row.Status);}
    [Fact] public void JanuaryEarlyReminderCrossesYear(){var row=Assert.Single(ImportService.ParseText("2026年1月初 启动会"));Assert.Equal(new DateTime(2026,1,1,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2025,12,31,7,30,0),row.ReminderAt);}
    [Fact] public void MonthWithoutAnyYearStillNeedsConfirmation(){var row=Assert.Single(ImportService.ParseText("8月中旬 客户走访"));Assert.Null(row.ScheduledAt);Assert.False(row.IsSelected);Assert.Equal(ImportStatus.需要确认,row.Status);Assert.Contains("没有可确定的年份",row.ErrorReason);}
    [Fact] public async Task ExcelMonthUsesWorkbookYearContext(){var path=Path.Combine(_dir,"month-context.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="2026年市场计划";ws.Cell(2,1).Value="日期";ws.Cell(2,2).Value="安排";ws.Cell(3,1).Value="9月中旬";ws.Cell(3,2).Value="拜访客户";wb.SaveAs(path);}var row=Assert.Single(await new ImportService(new FakeOcr()).ParseAsync(path));Assert.Equal(new DateTime(2026,9,13,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,9,12,7,30,0),row.ReminderAt);}
    [Fact] public async Task FutureYearColumnDoesNotEraseTitleBaseYear(){var path=Path.Combine(_dir,"year-boundary.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="26年新品计划";ws.Cell(2,1).Value="序号";ws.Cell(2,2).Value="项目";ws.Cell(2,3).Value="内容";ws.Cell(2,4).Value="日程";ws.Cell(3,4).Value="8月";ws.Cell(3,6).Value="27年";ws.Cell(4,4).Value=31;ws.Cell(5,1).Value=1;ws.Cell(5,2).Value="甲项目";ws.Cell(5,3).Value="评审";ws.Cell(5,4).Value="8月中旬评审";wb.SaveAs(path);}var row=Assert.Single(await new ImportService(new FakeOcr()).ParseAsync(path));Assert.Equal(new DateTime(2026,8,13,7,30,0),row.ScheduledAt);}
    [Fact] public void MultipleContextYearsDoNotResolveMonth(){var rows=ImportService.ParseText("2026年 甲项目\n2027年 乙项目\n8月中旬 客户走访");var row=Assert.Single(rows,x=>x.Content=="客户走访");Assert.Null(row.ScheduledAt);Assert.Equal(ImportStatus.需要确认,row.Status);}
    [Fact] public void AiPartialMonthUsesSameDefaultRules(){var row=Assert.Single(AiImportService.ParseCandidatesJson("{\"events\":[{\"date\":\"2026年8月下旬\",\"time\":\"\",\"content\":\"客户走访\",\"status\":\"certain\"}]}"));Assert.Equal(new DateTime(2026,8,19,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,8,18,7,30,0),row.ReminderAt);}
    [Fact] public void ClosedDateRangeExpandsEveryDayWithSharedReminderAndContent(){var rows=ImportService.ParseText("2026年7月20日至22日 客户打样");Assert.Equal(3,rows.Count);Assert.Equal([20,21,22],rows.Select(x=>x.ScheduledAt!.Value.Day));Assert.All(rows,x=>Assert.Equal(new DateTime(2026,7,19,7,30,0),x.ReminderAt));Assert.Single(rows.Select(x=>x.Content).Distinct());Assert.All(rows,x=>Assert.Contains("日期范围：2026/07/20—2026/07/22",x.Content));}
    [Fact] public void CrossYearRangeInfersNextYear(){var rows=ImportService.ParseText("2026年12月30日至1月2日 年度盘点");Assert.Equal(4,rows.Count);Assert.Equal(new DateTime(2026,12,30,7,30,0),rows[0].ScheduledAt);Assert.Equal(new DateTime(2027,1,2,7,30,0),rows[^1].ScheduledAt);Assert.All(rows,x=>Assert.Equal(new DateTime(2026,12,29,7,30,0),x.ReminderAt));}
    [Fact] public void FirstMonthWeekIsMondayThroughFirstSunday(){var rows=ImportService.ParseText("2026年8月第一周 项目评审");Assert.Equal(7,rows.Count);Assert.Equal(new DateTime(2026,7,27,7,30,0),rows[0].ScheduledAt);Assert.Equal(new DateTime(2026,8,2,7,30,0),rows[^1].ScheduledAt);Assert.All(rows,x=>Assert.Equal(new DateTime(2026,7,26,7,30,0),x.ReminderAt));Assert.Single(rows.Select(x=>x.Content).Distinct());}
    [Fact] public void SecondMonthWeekIsSevenDaysAfterFirstSunday(){var rows=ImportService.ParseText("2026年8月第2周 项目评审");Assert.Equal(new DateTime(2026,8,3,7,30,0),rows[0].ScheduledAt);Assert.Equal(new DateTime(2026,8,9,7,30,0),rows[^1].ScheduledAt);Assert.All(rows,x=>Assert.Equal(new DateTime(2026,8,2,7,30,0),x.ReminderAt));}
    [Fact] public void WeekWithoutYearRequiresConfirmation(){var row=Assert.Single(ImportService.ParseText("8月第三周 项目评审"));Assert.Null(row.ScheduledAt);Assert.Equal(ImportStatus.需要确认,row.Status);Assert.False(row.IsSelected);}
    [Fact] public void AiDateRangeExpandsThroughLocalRules(){var rows=AiImportService.ParseCandidatesJson("{\"events\":[{\"rangeStart\":\"2026/07/20\",\"rangeEnd\":\"2026/07/22\",\"content\":\"客户打样\",\"status\":\"certain\"}]}");Assert.Equal(3,rows.Count);Assert.All(rows,x=>Assert.Equal(new DateTime(2026,7,19,7,30,0),x.ReminderAt));}
    [Fact] public async Task DueRangeExcludesLeftAndIncludesRightBoundary(){var start=new DateTime(2026,8,1,8,0,0);foreach(var offset in new[]{0,1,2})await _repo.AddAsync(new(){ScheduledAt=start.AddHours(offset),ReminderAt=start.AddHours(offset),Content=$"R{offset}"});var rows=await _repo.GetDueRangeAsync(start,start.AddHours(1));Assert.Single(rows);Assert.Equal("R1",rows[0].Content);}
    [Fact] public async Task BatchMarksMissedRemindersTriggered(){var at=new DateTime(2026,8,1,8,0,0);var first=new ScheduleItem{ScheduledAt=at,ReminderAt=at,Content="甲"};var second=new ScheduleItem{ScheduledAt=at.AddMinutes(1),ReminderAt=at.AddMinutes(1),Content="乙"};await _repo.AddAsync(first);await _repo.AddAsync(second);await _repo.MarkTriggeredAsync(new[]{first.Id,second.Id});Assert.Empty(await _repo.GetDueAsync(at.AddHours(1)));}
    [Fact] public void MissedReminderSummaryContainsEverySchedule(){var text=ReminderSummaryFormatter.Build([new(){ScheduledAt=new DateTime(2026,8,2,9,0,0),ReminderAt=new DateTime(2026,8,1,7,30,0),Content="甲事项"},new(){ScheduledAt=new DateTime(2026,8,3,10,0,0),ReminderAt=new DateTime(2026,8,2,7,30,0),Content="乙事项"}]);Assert.Contains("甲事项",text);Assert.Contains("乙事项",text);Assert.Contains("2026/08/02 09:00",text);Assert.Contains("2026/08/02 07:30",text);}
    [Fact] public void ReminderCheckpointRoundTrips(){var path=Path.Combine(_dir,"checkpoint.txt");var store=new ReminderCheckpointStore(path);var value=new DateTime(2026,8,2,21,7,0,DateTimeKind.Local);store.Write(value);Assert.Equal(value,store.Read());}
    [Fact] public async Task PlanningCellWithWeekAndExactDateKeepsBothEvents(){var path=Path.Combine(_dir,"two-dates.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="26年计划";ws.Cell(2,1).Value="序号";ws.Cell(2,2).Value="项目";ws.Cell(2,3).Value="内容";ws.Cell(2,4).Value="日程";ws.Cell(3,4).Value="8月";ws.Cell(4,4).Value=31;ws.Cell(5,1).Value=1;ws.Cell(5,2).Value="甲项目";ws.Cell(5,3).Value="审核";ws.Cell(5,4).Value="8月第一周雄邦审核        8月25日奔驰审核";wb.SaveAs(path);}var rows=await new ImportService(new FakeOcr()).ParseAsync(path);Assert.Equal(8,rows.Count);Assert.Equal(7,rows.Count(x=>x.Content.Contains("识别片段：8月第一周雄邦审核")));Assert.Contains(rows,x=>x.ScheduledAt==new DateTime(2026,8,25,7,30,0)&&x.Content.Contains("识别片段：8月25日奔驰审核"));}
    [Fact] public void EveryAiProviderOffersTwoToFourModels(){Assert.Equal(4,AiProviderCatalog.All.Count);Assert.All(AiProviderCatalog.All,x=>Assert.InRange(x.Models.Count,2,4));}
    [Fact] public void AiJsonParsesExactDateTime(){var row=Assert.Single(AiImportService.ParseCandidatesJson("{\"events\":[{\"scheduledAt\":\"2026/08/02 21:07\",\"content\":\"新品评审\",\"status\":\"certain\"}]}"));Assert.Equal(new DateTime(2026,8,2,21,7,0),row.ScheduledAt);Assert.Equal(ImportStatus.可导入,row.Status);}
    [Fact] public void AiJsonDateOnlyUsesRequiredDefaults(){var row=Assert.Single(AiImportService.ParseCandidatesJson("{\"events\":[{\"date\":\"2026/08/02\",\"time\":\"\",\"content\":\"拜访客户\",\"status\":\"certain\"}]}"));Assert.Equal(new DateTime(2026,8,2,7,30,0),row.ScheduledAt);Assert.Equal(new DateTime(2026,8,1,7,30,0),row.ReminderAt);}
    [Fact] public void AiJsonDoesNotInventAmbiguousDate(){var row=Assert.Single(AiImportService.ParseCandidatesJson("```json\n{\"events\":[{\"date\":\"\",\"time\":\"\",\"content\":\"月底拜访\",\"status\":\"needs_confirmation\",\"reason\":\"日期不完整\"}]}\n```"));Assert.Null(row.ScheduledAt);Assert.False(row.IsSelected);Assert.Equal(ImportStatus.需要确认,row.Status);}
    [Fact] public void ParsesSemanticOverlapResponse(){var values=AiImportService.ParseOverlapIds("{\"overlaps\":[{\"candidateId\":\"N2\",\"existingId\":\"E8\",\"reason\":\"同一客户会议\"}]}");var match=Assert.Single(values);Assert.Equal("N2",match.Id);Assert.Contains("同一客户",match.Reason);}
    [Fact] public async Task SemanticOverlapIsNeverImported(){var result=await _repo.BatchImportAsync([new(){ScheduledAt=new DateTime(2026,8,2,7,30,0),Content="重合事项",Status=ImportStatus.高度重合,IsSelected=true}],DuplicateAction.仍然导入);Assert.Equal(1,result.Skipped);Assert.Empty(await _repo.GetAllAsync());}
    [Fact] public async Task AiPipelineCreatesExcelAndBlocksSemanticOverlap()
    {
        var input=Path.Combine(_dir,"ai-source.xlsx");using(var wb=new XLWorkbook()){var ws=wb.AddWorksheet("计划");ws.Cell(1,1).Value="客户";ws.Cell(1,2).Value="日期";ws.Cell(2,1).Value="星河新品评审";ws.Cell(2,2).Value="2026/08/02";wb.SaveAs(input);}
        var handler=new FakeAiHandler();var service=new AiImportService(new HttpClient(handler),new FakeOcr());var existing=new[]{new ScheduleItem{ScheduledAt=new DateTime(2026,8,2,7,30,0),ReminderAt=new DateTime(2026,8,1,7,30,0),Content="星河客户新品评审会"}};
        var result=await service.AnalyzeAsync(input,AiProvider.DeepSeek,"deepseek-v4-flash","test-key",existing);
        var row=Assert.Single(result.Candidates);Assert.Equal(ImportStatus.高度重合,row.Status);Assert.False(row.IsSelected);Assert.Equal(1,result.SemanticOverlapCount);Assert.True(File.Exists(result.ExcelPath));using var output=new XLWorkbook(result.ExcelPath);Assert.Equal("安排",output.Worksheet(1).Cell(1,4).GetString());File.Delete(result.ExcelPath);
    }
    private sealed class FakeAiHandler:System.Net.Http.HttpMessageHandler
    {
        private int _calls;
        protected override Task<System.Net.Http.HttpResponseMessage> SendAsync(System.Net.Http.HttpRequestMessage request,CancellationToken cancellationToken)
        {
            var content=Interlocked.Increment(ref _calls)==1?"{\"events\":[{\"date\":\"2026/08/02\",\"time\":\"\",\"content\":\"星河新品评审\",\"status\":\"certain\"}]}":"{\"overlaps\":[{\"candidateId\":\"N1\",\"existingId\":\"E1\",\"reason\":\"同日同一客户新品评审\"}]}";
            var envelope=System.Text.Json.JsonSerializer.Serialize(new{choices=new[]{new{message=new{content}}}});return Task.FromResult(new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK){Content=new StringContent(envelope)});
        }
    }
    private sealed class FakeOcr:IOcrService{public Task<string> ReadImageAsync(string path)=>Task.FromResult("");public Task<string> ReadPdfAsync(string path,IProgress<int>? progress=null)=>Task.FromResult("");}
}
