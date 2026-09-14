using CalendarReminder.Core.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Fonts;
using System.Drawing;
using System.Drawing.Imaging;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using P = DocumentFormat.OpenXml.Presentation;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace CalendarReminder.Core.Services;

public enum ExportFormat { Excel, Word, PDF, PowerPoint, PNG, JPG }

public sealed class ExportService
{
    public async Task<ExportResult> ExportAsync(IReadOnlyList<ScheduleItem> items, DateTime start, DateTime end, ExportFormat format, string path)
    {
        if (end < start) throw new ArgumentException("结束时间不得早于开始时间。");
        var selected = items.Where(x => x.ScheduledAt >= start && x.ScheduledAt <= end).OrderBy(x => x.ScheduledAt).ThenBy(x => x.CreatedAt).ToList();
        if (selected.Count == 0) throw new InvalidOperationException("所选时间范围内没有日程，不会生成空文件。");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
        try
        {
            return await Task.Run(() =>
            {
                var files = format switch
                {
                    ExportFormat.Excel => One(ExportExcel(selected, start, end, path)),
                    ExportFormat.Word => One(ExportWord(selected, start, end, path)),
                    ExportFormat.PDF => One(ExportPdf(selected, start, end, path)),
                    ExportFormat.PowerPoint => One(ExportPpt(selected, start, end, path)),
                    ExportFormat.PNG => ExportImages(selected, start, end, path, false),
                    ExportFormat.JPG => ExportImages(selected, start, end, path, true),
                    _ => throw new NotSupportedException()
                };
                return new ExportResult(files, selected.Count);
            });
        }
        catch (Exception ex) { AppLog.Error($"导出 {path}", ex); throw new IOException("导出失败。请确认保存位置可写、文件未被占用且磁盘空间充足。", ex); }
    }

    private static IReadOnlyList<string> One(string path) => [path];
    private static string Range(DateTime start, DateTime end) => $"时间范围：{DateTimeFormat.Format(start)} — {DateTimeFormat.Format(end)}";

    private static string ExportExcel(List<ScheduleItem> items, DateTime start, DateTime end, string path)
    {
        using var book = new XLWorkbook(); var sheet = book.AddWorksheet("日历安排");
        sheet.Cell(1, 1).Value = "日历安排"; sheet.Range(1, 1, 1, 3).Merge(); sheet.Cell(1, 1).Style.Font.SetBold().Font.SetFontSize(18);
        sheet.Cell(2, 1).Value = Range(start, end); sheet.Range(2, 1, 2, 3).Merge();
        sheet.Cell(3, 1).Value = $"生成时间：{DateTimeFormat.Format(DateTime.Now)}"; sheet.Range(3, 1, 3, 3).Merge();
        string[] headers = ["序号", "日期时间", "安排"]; for (var c = 0; c < 3; c++) { sheet.Cell(5, c + 1).Value = headers[c]; sheet.Cell(5, c + 1).Style.Font.SetBold(); sheet.Cell(5, c + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8EEFF"); }
        for (var i = 0; i < items.Count; i++) { sheet.Cell(i + 6, 1).Value = i + 1; sheet.Cell(i + 6, 2).Value = DateTimeFormat.Format(items[i].ScheduledAt); sheet.Cell(i + 6, 3).Value = items[i].Content; }
        sheet.Column(1).Width = 8; sheet.Column(2).Width = 22; sheet.Column(3).Width = 65; sheet.Column(3).Style.Alignment.WrapText = true; sheet.SheetView.FreezeRows(5);
        sheet.RangeUsed()!.Style.Border.OutsideBorder = XLBorderStyleValues.Thin; sheet.Range(5, 1, 5 + items.Count, 3).Style.Border.InsideBorder = XLBorderStyleValues.Hair;
        book.SaveAs(path); return path;
    }

    private static string ExportWord(List<ScheduleItem> items, DateTime start, DateTime end, string path)
    {
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document); var main = doc.AddMainDocumentPart(); main.Document = new W.Document(); var body = main.Document.AppendChild(new W.Body());
        body.Append(Paragraph("日历安排", 32, true, JustificationValues.Center)); body.Append(Paragraph(Range(start, end), 20)); body.Append(Paragraph($"生成时间：{DateTimeFormat.Format(DateTime.Now)}", 18));
        var table = new W.Table(new TableProperties(new TableBorders(new TopBorder { Val=BorderValues.Single }, new BottomBorder { Val=BorderValues.Single }, new LeftBorder { Val=BorderValues.Single }, new RightBorder { Val=BorderValues.Single }, new InsideHorizontalBorder { Val=BorderValues.Single }, new InsideVerticalBorder { Val=BorderValues.Single })), Row("日期时间", "安排", true));
        foreach (var item in items) table.Append(Row(DateTimeFormat.Format(item.ScheduledAt), item.Content, false)); body.Append(table);
        var section = new SectionProperties(new PageMargin { Top=720, Right=720, Bottom=720, Left=720 }); body.Append(section); main.Document.Save(); return path;
    }

    private static W.Paragraph Paragraph(string text, int size, bool bold=false, JustificationValues? align=null)
    {
        var props = new ParagraphProperties(); if (align.HasValue) props.Append(new Justification { Val = align.Value });
        var runProperties = new RunProperties(new RunFonts { EastAsia="Microsoft YaHei", Ascii="Microsoft YaHei" }, new FontSize { Val=size.ToString() }); if (bold) runProperties.Append(new Bold());
        return new W.Paragraph(props, new W.Run(runProperties, new W.Text(text) { Space=SpaceProcessingModeValues.Preserve }));
    }
    private static TableRow Row(string a, string b, bool bold) => new(new TableCell(Paragraph(a, 20, bold), new TableCellProperties(new TableCellWidth { Width="2600", Type=TableWidthUnitValues.Dxa })), new TableCell(Paragraph(b, 20, bold), new TableCellProperties(new TableCellWidth { Width="6500", Type=TableWidthUnitValues.Dxa })));

    private static string ExportPdf(List<ScheduleItem> items, DateTime start, DateTime end, string path)
    {
        if (GlobalFontSettings.FontResolver is null) GlobalFontSettings.FontResolver = new ChineseFontResolver();
        var doc = new PdfDocument(); doc.Info.Title = "日历安排"; var title = new XFont("Microsoft YaHei", 20, XFontStyleEx.Bold); var normal = new XFont("Microsoft YaHei", 10); var bold = new XFont("Microsoft YaHei", 10, XFontStyleEx.Bold);
        PdfPage? page = null; XGraphics? gfx = null; XTextFormatter? tf = null; double y = 0;
        void NewPage()
        {
            gfx?.Dispose(); page = doc.AddPage(); page.Size = PdfSharp.PageSize.A4; gfx = XGraphics.FromPdfPage(page); tf = new XTextFormatter(gfx); y = 45;
            gfx.DrawString("日历安排", title, XBrushes.DarkSlateBlue, new XRect(45, y, page.Width.Point-90, 30), XStringFormats.TopLeft); y += 38;
            gfx.DrawString(Range(start,end), normal, XBrushes.Black, 45, y); y += 20; gfx.DrawString($"生成时间：{DateTimeFormat.Format(DateTime.Now)}", normal, XBrushes.Black, 45, y); y += 30;
            gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(232,238,255)), 45,y,page.Width.Point-90,25); gfx.DrawString("日期时间",bold,XBrushes.Black,50,y+17); gfx.DrawString("安排",bold,XBrushes.Black,190,y+17); y+=30;
        }
        NewPage();
        foreach (var item in items)
        {
            var lines = Math.Max(1, (int)Math.Ceiling(item.Content.Length / 36d)); var h = Math.Max(30, lines * 18 + 8);
            if (y + h > page!.Height.Point - 45) NewPage();
            gfx!.DrawRectangle(XPens.LightGray, 45,y,page!.Width.Point-90,h); gfx.DrawLine(XPens.LightGray,185,y,185,y+h); gfx.DrawString(DateTimeFormat.Format(item.ScheduledAt),normal,XBrushes.Black,new XRect(50,y+7,130,h),XStringFormats.TopLeft); tf!.DrawString(item.Content,normal,XBrushes.Black,new XRect(190,y+5,page.Width.Point-240,h-8),XStringFormats.TopLeft); y+=h;
        }
        gfx?.Dispose(); doc.Save(path); return path;
    }

    private static string ExportPpt(List<ScheduleItem> items, DateTime start, DateTime end, string path)
    {
        using var doc = PresentationDocument.Create(path, PresentationDocumentType.Presentation); var pp = doc.AddPresentationPart(); pp.Presentation = new P.Presentation();
        var slideIds = pp.Presentation.AppendChild(new SlideIdList()); pp.Presentation.SlideSize = new SlideSize { Cx=12192000, Cy=6858000, Type=SlideSizeValues.Screen16x9 }; uint id = 256;
        AddSlide(pp, slideIds, id++, [("日历安排", 800000L, 1200000L, 10500000L, 1000000L, 32), (Range(start,end), 800000L, 2600000L, 10500000L, 600000L, 18), ($"生成时间：{DateTimeFormat.Format(DateTime.Now)}",800000L,3300000L,10500000L,600000L,16)]);
        const int perPage = 7;
        for (var offset=0; offset<items.Count; offset+=perPage)
        {
            var texts = new List<(string,long,long,long,long,int)> { ($"日历安排  {offset+1}—{Math.Min(offset+perPage,items.Count)}",600000,250000,11000000,500000,22) };
            for (var i=0;i<Math.Min(perPage,items.Count-offset);i++) { var item=items[offset+i]; var y=900000L+i*780000L; texts.Add((DateTimeFormat.Format(item.ScheduledAt),600000,y,2300000,650000,15)); texts.Add((item.Content,3000000,y,8300000,650000,15)); }
            AddSlide(pp,slideIds,id++,texts);
        }
        pp.Presentation.Save(); return path;
    }

    private static void AddSlide(PresentationPart pp, SlideIdList ids, uint id, IEnumerable<(string text,long x,long y,long w,long h,int size)> blocks)
    {
        var part=pp.AddNewPart<SlidePart>(); part.Slide=new P.Slide(new P.CommonSlideData(new P.ShapeTree(new P.NonVisualGroupShapeProperties(new P.NonVisualDrawingProperties{Id=1,Name=""},new P.NonVisualGroupShapeDrawingProperties(),new ApplicationNonVisualDrawingProperties()),new P.GroupShapeProperties(new A.TransformGroup(new A.Offset{X=0,Y=0},new A.Extents{Cx=0,Cy=0},new A.ChildOffset{X=0,Y=0},new A.ChildExtents{Cx=0,Cy=0})) )));
        uint sid=2; var tree=part.Slide.CommonSlideData!.ShapeTree!;
        foreach(var b in blocks) tree.Append(TextShape(sid++,b.text,b.x,b.y,b.w,b.h,b.size));
        part.Slide.Save(); ids.Append(new SlideId { Id=id, RelationshipId=pp.GetIdOfPart(part) });
    }
    private static P.Shape TextShape(uint id,string text,long x,long y,long w,long h,int size) => new(new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties{Id=id,Name=$"文本{id}"},new P.NonVisualShapeDrawingProperties(new A.ShapeLocks{NoGrouping=true}),new ApplicationNonVisualDrawingProperties()),new P.ShapeProperties(new A.Transform2D(new A.Offset{X=x,Y=y},new A.Extents{Cx=w,Cy=h}),new A.PresetGeometry(new A.AdjustValueList()){Preset=A.ShapeTypeValues.Rectangle},new A.NoFill()),new P.TextBody(new A.BodyProperties{Wrap=A.TextWrappingValues.Square},new A.ListStyle(),new A.Paragraph(new A.Run(new A.RunProperties{Language="zh-CN",FontSize=size*100},new A.Text(text)),new A.EndParagraphRunProperties{Language="zh-CN"})));

    private static IReadOnlyList<string> ExportImages(List<ScheduleItem> items, DateTime start, DateTime end, string path, bool jpg)
    {
        const int width=1400,height=1800,margin=80,row=120; var capacity=(height-330)/row; var pages=(int)Math.Ceiling(items.Count/(double)capacity); var result=new List<string>();
        for(var p=0;p<pages;p++)
        {
            var actual=pages==1?path:Path.Combine(Path.GetDirectoryName(path)!, $"{Path.GetFileNameWithoutExtension(path)}_{p+1:00}{Path.GetExtension(path)}");
            using var bmp=new Bitmap(width,height); bmp.SetResolution(150,150); using var g=Graphics.FromImage(bmp); g.Clear(System.Drawing.Color.White); g.TextRenderingHint=System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
            using var title=new System.Drawing.Font("Microsoft YaHei",28,FontStyle.Bold); using var normal=new System.Drawing.Font("Microsoft YaHei",16); using var bold=new System.Drawing.Font("Microsoft YaHei",16,FontStyle.Bold); using var pen=new Pen(System.Drawing.Color.FromArgb(220,225,235)); using var header=new SolidBrush(System.Drawing.Color.FromArgb(237,241,255));
            g.DrawString("日历安排",title,Brushes.MidnightBlue,margin,50); g.DrawString(Range(start,end),normal,Brushes.Black,margin,115); g.DrawString($"生成时间：{DateTimeFormat.Format(DateTime.Now)}    第 {p+1}/{pages} 页",normal,Brushes.DimGray,margin,155);
            var y=220; g.FillRectangle(header,margin,y,width-2*margin,55); g.DrawString("日期时间",bold,Brushes.Black,margin+15,y+10); g.DrawString("安排",bold,Brushes.Black,margin+300,y+10); y+=55;
            foreach(var item in items.Skip(p*capacity).Take(capacity)) { g.DrawRectangle(pen,margin,y,width-2*margin,row); g.DrawLine(pen,margin+285,y,margin+285,y+row); g.DrawString(DateTimeFormat.Format(item.ScheduledAt),normal,Brushes.Black,new RectangleF(margin+12,y+12,260,row-20)); g.DrawString(item.Content,normal,Brushes.Black,new RectangleF(margin+305,y+10,width-margin-335,row-15)); y+=row; }
            if(jpg) { var codec=ImageCodecInfo.GetImageEncoders().First(x=>x.FormatID==ImageFormat.Jpeg.Guid); using var parameters=new EncoderParameters(1); parameters.Param[0]=new EncoderParameter(System.Drawing.Imaging.Encoder.Quality,92L); bmp.Save(actual,codec,parameters); } else bmp.Save(actual,ImageFormat.Png); result.Add(actual);
        }
        return result;
    }
}

internal sealed class ChineseFontResolver : IFontResolver
{
    private const string Regular = "calendar-deng-regular";
    private const string Bold = "calendar-deng-bold";
    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic) => new(isBold ? Bold : Regular, false, isItalic);
    public byte[]? GetFont(string faceName)
    {
        var file = faceName == Bold ? "Dengb.ttf" : "Deng.ttf";
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", file);
        if (!File.Exists(path))
        {
            path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts", "simhei.ttf");
            if (!File.Exists(path)) throw new FileNotFoundException("系统中未找到可用于 PDF 中文导出的字体。", path);
        }
        return File.ReadAllBytes(path);
    }
}
