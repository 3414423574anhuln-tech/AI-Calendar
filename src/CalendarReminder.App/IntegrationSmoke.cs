using CalendarReminder.Core.Models;
using CalendarReminder.Core.Services;
using DocumentFormat.OpenXml.Packaging;
using System.Text.Json;
using UglyToad.PdfPig;

namespace CalendarReminder.App;

internal static class IntegrationSmoke
{
    public static async Task<int> RunAsync(ScheduleRepository repo,string? folder)
    {
        folder=string.IsNullOrWhiteSpace(folder)?Path.Combine(Path.GetTempPath(),"CalendarReminderSmoke"):Path.GetFullPath(folder);Directory.CreateDirectory(folder);var report=new Dictionary<string,object>();
        ScheduleItem? item=null;
        try
        {
            foreach(var stale in (await repo.GetAllAsync()).Where(x=>x.Content.StartsWith("集成验证日程："))) await repo.DeleteAsync(stale.Id);
            item=new ScheduleItem{ScheduledAt=DateTime.Now.AddDays(2).Date.AddHours(10).AddMinutes(15),Content="集成验证日程：准备项目报告"};await repo.AddAsync(item);report["created"]=true;item.ScheduledAt=item.ScheduledAt.AddMinutes(5);item.Content="集成验证日程：已修改";item.ReminderTriggered=true;await repo.UpdateAsync(item);var loaded=(await repo.GetAllAsync()).Single(x=>x.Id==item.Id);report["updated"]=loaded.Content==item.Content;report["reminderReset"]=!loaded.ReminderTriggered;
            var service=new ExportService();var start=loaded.ScheduledAt.AddMinutes(-1);var end=loaded.ScheduledAt.AddMinutes(1);var files=new List<string>();foreach(var f in Enum.GetValues<ExportFormat>()){var ext=f switch{ExportFormat.Excel=>"xlsx",ExportFormat.Word=>"docx",ExportFormat.PDF=>"pdf",ExportFormat.PowerPoint=>"pptx",ExportFormat.PNG=>"png",_=>"jpg"};var result=await service.ExportAsync([loaded],start,end,f,Path.Combine(folder,$"smoke.{ext}"));files.AddRange(result.Files);}
            using(var x=new ClosedXML.Excel.XLWorkbook(Path.Combine(folder,"smoke.xlsx")))report["xlsxRead"]=x.Worksheets.Any();using(var w=WordprocessingDocument.Open(Path.Combine(folder,"smoke.docx"),false))report["docxRead"]=w.MainDocumentPart?.Document.InnerText.Length>0;using(var p=PdfDocument.Open(Path.Combine(folder,"smoke.pdf")))report["pdfRead"]=p.NumberOfPages>0;using(var pp=PresentationDocument.Open(Path.Combine(folder,"smoke.pptx"),false))report["pptxRead"]=pp.PresentationPart?.SlideParts.Any()==true;report["imagesRead"]=files.Where(x=>x.EndsWith("png")||x.EndsWith("jpg")).All(x=>new FileInfo(x).Length>0);
            await repo.DeleteAsync(item.Id);report["deleted"]=!(await repo.GetAllAsync()).Any(x=>x.Id==item.Id);report["files"]=files;await File.WriteAllTextAsync(Path.Combine(folder,"smoke-result.json"),JsonSerializer.Serialize(report,new JsonSerializerOptions{WriteIndented=true}));return report.Values.OfType<bool>().All(x=>x)?0:2;
        }catch(Exception ex){AppLog.Error("集成冒烟测试",ex);report["error"]=ex.ToString();await File.WriteAllTextAsync(Path.Combine(folder,"smoke-result.json"),JsonSerializer.Serialize(report,new JsonSerializerOptions{WriteIndented=true}));return 1;}
        finally{if(item is not null)try{await repo.DeleteAsync(item.Id);}catch{} }
    }
}
