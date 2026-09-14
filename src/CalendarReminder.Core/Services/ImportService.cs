using CalendarReminder.Core.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace CalendarReminder.Core.Services;

public sealed class ImportService
{
    private static readonly TimeSpan DefaultScheduleTime = new(7, 30, 0);
    private static readonly Regex DateOnlyRegex = new(@"(?<y>\d{4})\s*(?:[/\-.年])\s*(?<m>\d{1,2})\s*(?:[/\-.月])\s*(?<d>\d{1,2})(?:\s*日)?(?:\s*(?<range>[~～\-—至])\s*(?<endDay>\d{1,2})\s*日?)?", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex PartialMonthRegex = new(@"(?:(?<year>\d{2,4})\s*年\s*)?(?<month>\d{1,2}|十[一二]?|[一二三四五六七八九])\s*月\s*(?<period>上旬|初旬|初期|初段|初|开头|开始|中旬|中期|中段|下旬|后期|后段|底|末期|月末|末|尾)?(?!\s*\d{1,2}\s*[日号])", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex YearOnlyRegex = new(@"(?<!\d)(?<year>\d{2,4})\s*年(?!\s*(?:\d{1,2}|十[一二]?|[一二三四五六七八九])\s*月)", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex ChineseRangeRegex = new(@"(?:(?<y1>\d{2,4})\s*年\s*)?(?<m1>\d{1,2})\s*月\s*(?<d1>\d{1,2})\s*(?:日|号)?\s*(?:~|～|—|至|到|-)\s*(?:(?<y2>\d{2,4})\s*年\s*)?(?:(?<m2>\d{1,2})\s*月\s*)?(?<d2>\d{1,2})\s*(?:日|号)", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex NumericRangeRegex = new(@"(?<y1>\d{4})\s*[/\-.]\s*(?<m1>\d{1,2})\s*[/\-.]\s*(?<d1>\d{1,2})\s*(?:~|～|—|至|到|-)\s*(?:(?<y2>\d{4})\s*[/\-.]\s*)?(?:(?<m2>\d{1,2})\s*[/\-.]\s*)?(?<d2>\d{1,2})", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex MonthWeekRegex = new(@"(?:(?<year>\d{2,4})\s*年\s*)?(?<month>\d{1,2}|十[一二]?|[一二三四五六七八九])\s*月\s*第?\s*(?<week>\d|[一二三四五六])\s*(?:个)?\s*周", RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly string[] DateHeaders = ["日期时间", "日期", "时间", "date", "datetime"];
    private static readonly string[] ContentHeaders = ["安排", "内容", "事项", "日程", "备注", "content", "subject"];
    private readonly IOcrService _ocr;
    public ImportService(IOcrService? ocr = null) => _ocr = ocr ?? new WindowsOcrService();

    public async Task<List<ImportCandidate>> ParseAsync(string path, IProgress<int>? progress = null)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("选择的文件不存在。", path);
        var extension = Path.GetExtension(path).ToLowerInvariant();
        try
        {
            progress?.Report(5);
            var rows = extension switch
            {
                ".xls" or ".xlsx" => await Task.Run(() => ParseExcel(path)),
                ".docx" => await Task.Run(() => ParseWord(path)),
                ".pptx" => await Task.Run(() => ParsePowerPoint(path)),
                ".pdf" => await ParsePdfAsync(path, progress),
                ".png" or ".jpg" or ".jpeg" => ParseText(await _ocr.ReadImageAsync(path)),
                _ => throw new NotSupportedException("不支持此文件类型。")
            };
            Normalize(rows); progress?.Report(100); return rows;
        }
        catch (Exception ex) when (ex is not NotSupportedException and not FileNotFoundException)
        { AppLog.Error($"导入文件 {path}", ex); throw new InvalidDataException("文件无法解析，可能已损坏、被占用或格式与扩展名不符。", ex); }
    }

    public static List<ImportCandidate> ParseText(string text) => ParseText(text, FindContextYear(text));

    private static List<ImportCandidate> ParseText(string text, int? contextYear)
    {
        var result = new List<ImportCandidate>();
        foreach (var raw in text.Replace("\r", "").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (DateTimeFormat.TryExtract(raw, out var at, out var content))
                result.Add(Create(at, content));
            else if (TryCreateDateRangeCandidates(raw, contextYear, null, out var rangeRows))
                result.AddRange(rangeRows);
            else if (TryCreateWeekCandidates(raw, contextYear, null, out var weekRows))
                result.AddRange(weekRows);
            else if (TryExtractDateOnly(raw, out var date, out content, out var isRange))
                result.Add(CreateWithoutSpecificTime(date, content, isRange));
            else if (TryExtractPartialDate(raw, contextYear, out var partial))
                result.Add(partial);
            else if (raw.Any(char.IsDigit))
                result.Add(new ImportCandidate { Content = raw, Status = ImportStatus.需要确认, ErrorReason = "未识别出完整的年月日和时间。" });
        }
        return result;
    }

    private static List<ImportCandidate> ParseExcel(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var data = reader.AsDataSet(new ExcelDataSetConfiguration { ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = false } });
        var result = new List<ImportCandidate>();
        foreach (DataTable table in data.Tables)
        {
            var before = result.Count; var contextYear = FindPlanningYear(table);
            int header = -1, dateColumn = -1, contentColumn = -1;
            for (var r = 0; r < Math.Min(10, table.Rows.Count); r++)
            {
                for (var c = 0; c < table.Columns.Count; c++)
                {
                    var value = table.Rows[r][c]?.ToString()?.Trim().ToLowerInvariant() ?? "";
                    if (DateHeaders.Contains(value)) { header = r; dateColumn = c; }
                    if (ContentHeaders.Contains(value)) { header = r; contentColumn = c; }
                }
                if (dateColumn >= 0 || contentColumn >= 0) break;
            }
            for (var r = Math.Max(0, header + 1); r < table.Rows.Count; r++)
            {
                if (dateColumn >= 0)
                {
                    var dateCell = table.Rows[r][dateColumn];
                    var content = contentColumn >= 0 ? table.Rows[r][contentColumn]?.ToString()?.Trim() ?? "" : "";
                    if (TryCreateDateRangeCandidates(dateCell?.ToString() ?? "", contextYear, content, out var rangeRows)) result.AddRange(rangeRows);
                    else if (TryCreateWeekCandidates(dateCell?.ToString() ?? "", contextYear, content, out var weekRows)) result.AddRange(weekRows);
                    else if (TryCellDate(dateCell, out var at, out var missingTime, out var ambiguousDate))
                        result.Add(missingTime ? CreateWithoutSpecificTime(at, content, ambiguousDate) : Create(at, content));
                    else if (TryExtractPartialDate($"{dateCell} {content}", contextYear, out var partial)) result.Add(partial);
                    else if (!string.IsNullOrWhiteSpace(dateCell?.ToString())) result.Add(new ImportCandidate { Content = content, Status = ImportStatus.需要确认, ErrorReason = $"无法识别日期：{dateCell}" });
                }
                else
                {
                    foreach (var cell in table.Rows[r].ItemArray)
                        if (cell is not null && DateTimeFormat.TryExtract(cell.ToString()!, out var at, out var content)) { result.Add(Create(at, content)); break; }
                }
            }
            if (result.Count == before) result.AddRange(ParsePlanningTable(table));
        }
        return result;
    }

    private static List<ImportCandidate> ParsePlanningTable(DataTable table)
    {
        var result = new List<ImportCandidate>();
        var year = FindPlanningYear(table);
        var titleMonth = FindPlanningMonth(table);
        var headerRow = FindHeaderRow(table);
        if (headerRow < 0) return result;

        var timelineColumn = -1;
        for (var c = 0; c < table.Columns.Count; c++)
        {
            var header = Cell(table, headerRow, c);
            if (header.Contains("日程", StringComparison.Ordinal) || header.Contains("计划", StringComparison.Ordinal))
            { timelineColumn = c; break; }
        }

        if (timelineColumn >= 0)
        {
            var monthRow = headerRow + 1;
            var weekRow = headerRow + 2;
            var months = BuildMonthMap(table, monthRow, timelineColumn);
            var years = BuildYearMap(table, monthRow, timelineColumn, year);
            for (var r = headerRow + 3; r < table.Rows.Count; r++)
            {
                var baseContent = BuildRowContent(table, headerRow, r, timelineColumn);
                if (string.IsNullOrWhiteSpace(baseContent)) continue;
                var foundEvent = false;
                for (var c = timelineColumn; c < table.Columns.Count; c++)
                {
                    var eventText = Cell(table, r, c);
                    if (string.IsNullOrWhiteSpace(eventText) || double.TryParse(eventText, out _)) continue;
                    foundEvent = true;
                    var month = months.TryGetValue(c, out var mappedMonth) ? mappedMonth : titleMonth;
                    var week = int.TryParse(Cell(table, weekRow, c), out var parsedWeek) ? parsedWeek : (int?)null;
                    var eventYear = years.TryGetValue(c, out var mappedYear) ? mappedYear : year;
                    result.AddRange(CreatePlanningCandidates(eventYear, month, week, eventText, CombineContent(baseContent, eventText)));
                }
                if (!foundEvent)
                    result.Add(new ImportCandidate { Content = baseContent, Status = ImportStatus.需要确认, ErrorReason = "未找到明确的日程日期，请手动选择日期和时间。" });
            }
            return result;
        }

        for (var r = headerRow + 1; r < table.Rows.Count; r++)
        {
            var content = BuildRowContent(table, headerRow, r, table.Columns.Count);
            if (string.IsNullOrWhiteSpace(content)) continue;
            var sourceText = string.Join(" ", table.Rows[r].ItemArray.Select(x => x?.ToString()).Where(x => !string.IsNullOrWhiteSpace(x)));
            result.AddRange(CreatePlanningCandidates(year, titleMonth, null, sourceText, content));
        }
        return result;
    }

    private static IReadOnlyList<ImportCandidate> CreatePlanningCandidates(int? year, int? columnMonth, int? week, string dateText, string content)
    {
        var datedSegments = Regex.Split(dateText, @"(?:\r?\n|\s{2,}|[；;])")
            .Select(x => x.Trim()).Where(x => !string.IsNullOrWhiteSpace(x) && Regex.IsMatch(x, @"(?:\d{1,4}\s*[年月日号]|[一二三四五六]\s*(?:个)?周)"))
            .ToList();
        if (datedSegments.Count > 1)
        {
            var combined = new List<ImportCandidate>();
            foreach (var segment in datedSegments) combined.AddRange(CreatePlanningCandidatesSingle(year, columnMonth, week, segment, $"{content}；识别片段：{segment}"));
            return combined;
        }
        return CreatePlanningCandidatesSingle(year, columnMonth, week, dateText, content);
    }

    private static IReadOnlyList<ImportCandidate> CreatePlanningCandidatesSingle(int? year, int? columnMonth, int? week, string dateText, string content)
    {
        if (TryCreateDateRangeCandidates(dateText, year, content, out var rangeRows)) return rangeRows;
        if (TryCreateWeekCandidates(dateText, year, content, out var weekRows)) return weekRows;
        return [CreatePlanningCandidate(year, columnMonth, week, dateText, content)];
    }

    private static ImportCandidate CreatePlanningCandidate(int? year, int? columnMonth, int? week, string dateText, string content)
    {
        var dateMatch = Regex.Match(dateText, @"(?:(?<year>\d{2,4})年)?\s*(?<month>\d{1,2})月\s*(?<day>\d{1,2})(?:\s*(?<range>[~～\-—至])\s*(?<endDay>\d{1,2}))?\s*日?");
        if (!dateMatch.Success && columnMonth.HasValue)
            dateMatch = Regex.Match(dateText, @"(?<day>\d{1,2})\s*日");

        if (dateMatch.Success)
        {
            var parsedYear = ParseYear(dateMatch.Groups["year"].Value) ?? year;
            var parsedMonth = dateMatch.Groups["month"].Success ? int.Parse(dateMatch.Groups["month"].Value) : columnMonth;
            var parsedDay = int.Parse(dateMatch.Groups["day"].Value);
            if (parsedYear.HasValue && parsedMonth.HasValue && TryDate(parsedYear.Value, parsedMonth.Value, parsedDay, out var date))
            {
                var isRange = dateMatch.Groups["range"].Success;
                return new ImportCandidate
                {
                    ScheduledAt = date.Add(DefaultScheduleTime),
                    ReminderAt = DefaultReminder(date),
                    Content = content,
                    Status = isRange ? ImportStatus.需要确认 : ImportStatus.可导入,
                    ErrorReason = isRange
                        ? $"识别到日期范围“{dateMatch.Value.Trim()}”，已暂用起始日期；时间默认 07:30，提醒为前一天 07:30，请确认具体日期。"
                        : "未识别到具体时间，已默认日程 07:30、提醒为前一天 07:30。"
                };
            }
        }

        if (TryExtractPartialDate(dateText, year, out var partial))
        {
            partial.Content = content;
            return partial;
        }

        var explicitMonth = Regex.Match(dateText, @"(?<!\d)(?<month>\d{1,2})\s*月");
        var approximateMonth = explicitMonth.Success && int.TryParse(explicitMonth.Groups["month"].Value, out var textMonth) && textMonth is >= 1 and <= 12 ? textMonth : columnMonth;
        var context = year.HasValue ? $"{year}年" : "";
        if (approximateMonth.HasValue) context += $"{approximateMonth}月";
        if (week.HasValue) context += $"第{week}周";
        return new ImportCandidate
        {
            Content = content,
            Status = ImportStatus.需要确认,
            ErrorReason = string.IsNullOrEmpty(context)
                ? "未识别到完整日期和时间，请手动确认。"
                : $"仅能确定大致时间（{context}），请手动选择具体日期和时间。"
        };
    }

    private static int FindHeaderRow(DataTable table)
    {
        for (var r = 0; r < Math.Min(12, table.Rows.Count); r++)
        {
            var values = table.Rows[r].ItemArray.Select(x => x?.ToString()?.Trim() ?? "").ToArray();
            if (values.Any(x => x == "序号") && values.Any(x => x is "内容" or "项目" or "客户")) return r;
        }
        return -1;
    }

    private static int? FindPlanningYear(DataTable table)
    {
        // The title establishes the base year. Later year labels usually mark a future column boundary.
        var titleYears = new HashSet<int>();
        for (var r = 0; r < Math.Min(2, table.Rows.Count); r++)
            for (var c = 0; c < table.Columns.Count; c++)
            {
                var match = Regex.Match(Cell(table, r, c), @"(?<!\d)(?<year>\d{2,4})\s*年");
                var year = match.Success ? ParseYear(match.Groups["year"].Value) : null;
                if (year.HasValue) titleYears.Add(year.Value);
            }
        if (titleYears.Count == 1) return titleYears.Single();
        if (titleYears.Count > 1) return null;

        var contextYears = new HashSet<int>();
        for (var r = 2; r < Math.Min(5, table.Rows.Count); r++)
            for (var c = 0; c < table.Columns.Count; c++)
            {
                var match = Regex.Match(Cell(table, r, c), @"(?<!\d)(?<year>\d{2,4})\s*年");
                var year = match.Success ? ParseYear(match.Groups["year"].Value) : null;
                if (year.HasValue) contextYears.Add(year.Value);
            }
        return contextYears.Count == 1 ? contextYears.Single() : null;
    }

    private static int? FindContextYear(string text)
    {
        var years = Regex.Matches(text, @"(?<!\d)(?<year>\d{2,4})\s*年")
            .Select(x => ParseYear(x.Groups["year"].Value)).Where(x => x.HasValue).Select(x => x!.Value).Distinct().ToList();
        return years.Count == 1 ? years[0] : null;
    }

    private static int? FindPlanningMonth(DataTable table)
    {
        for (var r = 0; r < Math.Min(5, table.Rows.Count); r++)
            for (var c = 0; c < table.Columns.Count; c++)
            {
                var match = Regex.Match(Cell(table, r, c), @"(?<!\d)(?<month>\d{1,2})\s*月");
                if (match.Success && int.TryParse(match.Groups["month"].Value, out var month) && month is >= 1 and <= 12) return month;
            }
        return null;
    }

    private static Dictionary<int, int> BuildMonthMap(DataTable table, int monthRow, int startColumn)
    {
        var result = new Dictionary<int, int>(); int? current = null;
        for (var c = startColumn; c < table.Columns.Count; c++)
        {
            var match = Regex.Match(Cell(table, monthRow, c), @"(?<month>\d{1,2})\s*月");
            if (match.Success && int.TryParse(match.Groups["month"].Value, out var month) && month is >= 1 and <= 12) current = month;
            if (current.HasValue) result[c] = current.Value;
        }
        return result;
    }

    private static Dictionary<int, int> BuildYearMap(DataTable table, int row, int startColumn, int? initialYear)
    {
        var result = new Dictionary<int, int>(); var current = initialYear;
        for (var c = startColumn; c < table.Columns.Count; c++)
        {
            var match = Regex.Match(Cell(table, row, c), @"(?<!\d)(?<year>\d{2,4})\s*年");
            if (match.Success) current = ParseYear(match.Groups["year"].Value);
            if (current.HasValue) result[c] = current.Value;
        }
        return result;
    }

    private static string BuildRowContent(DataTable table, int headerRow, int row, int endColumn)
    {
        var parts = new List<string>();
        for (var c = 0; c < Math.Min(endColumn, table.Columns.Count); c++)
        {
            var header = Cell(table, headerRow, c);
            var value = Cell(table, row, c);
            if (string.IsNullOrWhiteSpace(value) || header == "序号" || double.TryParse(value, out _)) continue;
            if (header is "客户" or "终端客户" or "品名" or "项目" or "内容" or "状态" or "问题点" or "说明" or "事项" or "日程")
                parts.Add(string.IsNullOrWhiteSpace(header) ? value : $"{header}：{value}");
        }
        return string.Join("；", parts.Distinct());
    }

    private static string CombineContent(string baseContent, string eventText) => $"{baseContent}；计划：{eventText}";

    private static int? ParseYear(string value)
    {
        if (!int.TryParse(value, out var year)) return null;
        return year < 100 ? 2000 + year : year;
    }

    private static bool TryDate(int year, int month, int day, out DateTime date)
    {
        try { date = new DateTime(year, month, day, 0, 0, 0, DateTimeKind.Local); return true; }
        catch (ArgumentOutOfRangeException) { date = default; return false; }
    }

    private static string Cell(DataTable table, int row, int column)
        => row >= 0 && row < table.Rows.Count && column >= 0 && column < table.Columns.Count
            ? table.Rows[row][column]?.ToString()?.Trim() ?? ""
            : "";

    private static bool TryCellDate(object? cell, out DateTime at, out bool missingTime, out bool ambiguousDate)
    {
        missingTime = false; ambiguousDate = false;
        if (cell is DateTime dt) { at = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, 0); missingTime = dt.TimeOfDay == TimeSpan.Zero; return true; }
        if (cell is double serial && serial is > 1 and < 2958466) { at = DateTime.FromOADate(serial); at = new DateTime(at.Year, at.Month, at.Day, at.Hour, at.Minute, 0); missingTime = Math.Abs(serial-Math.Truncate(serial)) < 0.0000001; return true; }
        return TryImportDate(cell?.ToString(), out at, out missingTime, out ambiguousDate);
    }

    private static List<ImportCandidate> ParseWord(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false); var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("Word 文档没有正文。");
        var result = new List<ImportCandidate>(); var contextYear = FindContextYear(body.InnerText);
        foreach (var table in body.Descendants<Table>())
        {
            foreach (var row in table.Elements<TableRow>())
            {
                var cells = row.Elements<TableCell>().Select(c => c.InnerText.Trim()).ToList();
                result.AddRange(ParseText(string.Join(" ", cells), contextYear));
            }
        }
        foreach (var p in body.Elements<Paragraph>()) result.AddRange(ParseText(p.InnerText, contextYear));
        return result;
    }

    private static List<ImportCandidate> ParsePowerPoint(string path)
    {
        using var doc = PresentationDocument.Open(path, false); var lines = new List<string>();
        foreach (var part in doc.PresentationPart?.SlideParts ?? [])
        {
            lines.AddRange(part.Slide.Descendants<A.Paragraph>().Select(x => string.Concat(x.Descendants<A.Text>().Select(t => t.Text))).Where(x => !string.IsNullOrWhiteSpace(x)));
            foreach (var row in part.Slide.Descendants<A.TableRow>()) lines.Add(string.Join(" ", row.Descendants<A.TableCell>().Select(x => x.InnerText)));
        }
        return ParseText(string.Join("\n", lines));
    }

    private async Task<List<ImportCandidate>> ParsePdfAsync(string path, IProgress<int>? progress)
    {
        var lines = new List<string>(); int pageCount;
        using (var pdf = PdfDocument.Open(path)) { pageCount = pdf.NumberOfPages; foreach (var page in pdf.GetPages()) lines.Add(page.Text); }
        var text = string.Join("\n", lines); if (text.Count(char.IsLetterOrDigit) >= Math.Max(12, pageCount * 6)) return ParseText(text);
        progress?.Report(30);
        var ocr = await _ocr.ReadPdfAsync(path, progress); return ParseText(ocr);
    }

    private static bool TryImportDate(string? value, out DateTime at, out bool missingTime, out bool ambiguousDate)
    {
        at = default; missingTime = false; ambiguousDate = false;
        if (DateTimeFormat.TryParse(value, out at)) return true;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var match = DateOnlyRegex.Match(value.Replace('：', ':'));
        if (!match.Success || !int.TryParse(match.Groups["y"].Value, out var year) || !int.TryParse(match.Groups["m"].Value, out var month) || !int.TryParse(match.Groups["d"].Value, out var day) || !TryDate(year, month, day, out at)) return false;
        missingTime = true; ambiguousDate = match.Groups["range"].Success; return true;
    }

    private static bool TryExtractDateOnly(string input, out DateTime date, out string remaining, out bool ambiguousDate)
    {
        date = default; remaining = input.Trim(); ambiguousDate = false;
        var match = DateOnlyRegex.Match(input.Replace('：', ':'));
        if (!match.Success || !TryImportDate(match.Value, out date, out _, out ambiguousDate)) return false;
        remaining = (input[..match.Index] + " " + input[(match.Index + match.Length)..]).Trim(' ', '-', '—', ':', '：', '|');
        return true;
    }

    private static bool TryExtractPartialDate(string input, int? contextYear, out ImportCandidate candidate)
    {
        candidate = null!; var normalized = input.Replace('：', ':');
        var monthMatch = PartialMonthRegex.Match(normalized);
        if (monthMatch.Success && TryParseMonth(monthMatch.Groups["month"].Value, out var month))
        {
            var year = ParseYear(monthMatch.Groups["year"].Value) ?? contextYear;
            var remaining = RemoveMatch(input, monthMatch);
            if (!year.HasValue)
            {
                candidate = new ImportCandidate { Content = remaining, Status = ImportStatus.需要确认, IsSelected = false, ErrorReason = $"识别到{month}月，但文件中没有可确定的年份，请手动选择日期。" };
                return true;
            }
            var period = monthMatch.Groups["period"].Value;
            var day = PartialMonthDay(period); var date = new DateTime(year.Value, month, day);
            var label = string.IsNullOrWhiteSpace(period) ? $"{month}月" : $"{month}月{period}";
            candidate = CreateInferred(date, remaining, $"仅识别到“{label}”，已默认日程为 {date:yyyy/MM/dd} 07:30、提醒为 {date.AddDays(-1):yyyy/MM/dd} 07:30。");
            return true;
        }

        var yearMatch = YearOnlyRegex.Match(normalized);
        if (!yearMatch.Success || !int.TryParse(yearMatch.Groups["year"].Value, out var rawYear)) return false;
        var parsedYear = ParseYear(rawYear.ToString()); if (!parsedYear.HasValue) return false;
        var firstDay = new DateTime(parsedYear.Value, 1, 1); var yearContent = RemoveMatch(input, yearMatch);
        candidate = CreateInferred(firstDay, yearContent, $"仅识别到“{parsedYear}年”，已默认日程为 {parsedYear}/01/01 07:30、提醒为 {parsedYear.Value - 1}/12/31 07:30。");
        return true;
    }

    private static bool TryCreateDateRangeCandidates(string input, int? contextYear, string? contentOverride, out List<ImportCandidate> candidates)
    {
        candidates = []; var match = ChineseRangeRegex.Match(input);
        if (!match.Success) match = NumericRangeRegex.Match(input);
        if (!match.Success) return false;

        var startYear = ParseYear(match.Groups["y1"].Value) ?? contextYear;
        if (!int.TryParse(match.Groups["m1"].Value, out var startMonth) || !int.TryParse(match.Groups["d1"].Value, out var startDay)) return false;
        var baseContent = contentOverride ?? RemoveMatch(input, match);
        if (!startYear.HasValue)
        {
            candidates.Add(new ImportCandidate { Content = baseContent, Status = ImportStatus.需要确认, IsSelected = false, ErrorReason = $"识别到日期范围“{match.Value.Trim()}”，但无法确定年份。" });
            return true;
        }

        var endYear = ParseYear(match.Groups["y2"].Value) ?? startYear;
        var endMonth = int.TryParse(match.Groups["m2"].Value, out var parsedEndMonth) ? parsedEndMonth : startMonth;
        if (!TryDate(startYear.Value, startMonth, startDay, out var start)) return InvalidRange(match.Value, baseContent, "范围起始日期无效。", out candidates);
        if (!match.Groups["y2"].Success && match.Groups["m2"].Success && endMonth < startMonth) endYear = startYear + 1;
        if (!endYear.HasValue || !int.TryParse(match.Groups["d2"].Value, out var endDay) || !TryDate(endYear.Value, endMonth, endDay, out var end)) return InvalidRange(match.Value, baseContent, "范围结束日期无效。", out candidates);
        if (end < start) return InvalidRange(match.Value, baseContent, "范围结束日期早于起始日期。", out candidates);
        candidates = CreateRangeCandidates(start, end, baseContent, match.Value.Trim());
        return true;
    }

    private static bool TryCreateWeekCandidates(string input, int? contextYear, string? contentOverride, out List<ImportCandidate> candidates)
    {
        candidates = []; var match = MonthWeekRegex.Match(input); if (!match.Success) return false;
        var year = ParseYear(match.Groups["year"].Value) ?? contextYear;
        if (!TryParseMonth(match.Groups["month"].Value, out var month) || !TryParseWeek(match.Groups["week"].Value, out var week)) return false;
        var baseContent = contentOverride ?? RemoveMatch(input, match);
        if (!year.HasValue)
        {
            candidates.Add(new ImportCandidate { Content = baseContent, Status = ImportStatus.需要确认, IsSelected = false, ErrorReason = $"识别到{month}月第{week}周，但文件中没有可确定的年份。" });
            return true;
        }
        var firstDay = new DateTime(year.Value, month, 1);
        var firstSunday = firstDay.AddDays(((int)DayOfWeek.Sunday - (int)firstDay.DayOfWeek + 7) % 7);
        var start = firstSunday.AddDays(-6 + (week - 1) * 7); var end = start.AddDays(6);
        candidates = CreateRangeCandidates(start, end, baseContent, $"{year}年{month}月第{week}周");
        return true;
    }

    private static List<ImportCandidate> CreateRangeCandidates(DateTime start, DateTime end, string content, string sourceLabel)
    {
        var trimmed = content.Trim(); var reminder = start.Date.AddDays(-1).Add(DefaultScheduleTime);
        var normalizedRange = $"{start:yyyy/MM/dd}—{end:yyyy/MM/dd}";
        var arranged = string.IsNullOrWhiteSpace(trimmed) ? "" : $"日期范围：{normalizedRange}（原文：{sourceLabel}）；{trimmed}";
        var count = (end.Date - start.Date).Days + 1; var result = new List<ImportCandidate>(count);
        for (var date = start.Date; date <= end.Date; date = date.AddDays(1))
            result.Add(new ImportCandidate
            {
                ScheduledAt = date.Add(DefaultScheduleTime), ReminderAt = reminder, Content = arranged,
                Status = string.IsNullOrWhiteSpace(trimmed) ? ImportStatus.内容为空 : ImportStatus.可导入,
                ErrorReason = string.IsNullOrWhiteSpace(trimmed) ? "安排内容为空。" : $"闭区间 {normalizedRange} 已展开为 {count} 条日程；全部提醒时间为 {reminder:yyyy/MM/dd HH:mm}。",
                IsSelected = !string.IsNullOrWhiteSpace(trimmed)
            });
        return result;
    }

    private static bool InvalidRange(string source, string content, string reason, out List<ImportCandidate> candidates)
    {
        candidates = [new ImportCandidate { Content = content, Status = ImportStatus.需要确认, IsSelected = false, ErrorReason = $"日期范围“{source.Trim()}”无效：{reason}" }];
        return true;
    }

    private static bool TryParseWeek(string value, out int week)
    {
        if (int.TryParse(value, out week)) return week is >= 1 and <= 6;
        week = value switch { "一" => 1, "二" => 2, "三" => 3, "四" => 4, "五" => 5, "六" => 6, _ => 0 };
        return week != 0;
    }

    private static ImportCandidate CreateInferred(DateTime date, string content, string reason)
    {
        var trimmed = content.Trim();
        return new ImportCandidate
        {
            ScheduledAt = date.Date.Add(DefaultScheduleTime), ReminderAt = date.Date.AddDays(-1).Add(DefaultScheduleTime), Content = trimmed,
            Status = string.IsNullOrWhiteSpace(trimmed) ? ImportStatus.内容为空 : ImportStatus.可导入,
            ErrorReason = string.IsNullOrWhiteSpace(trimmed) ? "安排内容为空。" : reason,
            IsSelected = !string.IsNullOrWhiteSpace(trimmed)
        };
    }

    private static string RemoveMatch(string input, Match match)
        => (input[..match.Index] + " " + input[(match.Index + match.Length)..]).Trim(' ', '-', '—', ':', '：', '|', '，', ',');

    private static int PartialMonthDay(string period)
    {
        if (period is "中旬" or "中期" or "中段") return 13;
        if (period is "下旬" or "后期" or "后段") return 19;
        if (period is "底" or "末期" or "月末" or "末" or "尾") return 25;
        return 1;
    }

    private static bool TryParseMonth(string value, out int month)
    {
        if (int.TryParse(value, out month)) return month is >= 1 and <= 12;
        month = value switch { "一" => 1, "二" => 2, "三" => 3, "四" => 4, "五" => 5, "六" => 6, "七" => 7, "八" => 8, "九" => 9, "十" => 10, "十一" => 11, "十二" => 12, _ => 0 };
        return month != 0;
    }

    private static DateTime DefaultReminder(DateTime date)
        => (date.Date > DateTime.MinValue.Date ? date.Date.AddDays(-1) : date.Date).Add(DefaultScheduleTime);

    private static ImportCandidate CreateWithoutSpecificTime(DateTime date, string content, bool ambiguousDate)
    {
        var trimmed = content.Trim();
        return new ImportCandidate
        {
            ScheduledAt = date.Date.Add(DefaultScheduleTime), ReminderAt = DefaultReminder(date), Content = trimmed,
            Status = string.IsNullOrWhiteSpace(trimmed) ? ImportStatus.内容为空 : ambiguousDate ? ImportStatus.需要确认 : ImportStatus.可导入,
            ErrorReason = string.IsNullOrWhiteSpace(trimmed) ? "安排内容为空。" : ambiguousDate
                ? "日期范围存在歧义，已暂用起始日期；时间默认 07:30，提醒为前一天 07:30，请确认具体日期。"
                : "未识别到具体时间，已默认日程 07:30、提醒为前一天 07:30。"
        };
    }

    private static ImportCandidate Create(DateTime at, string content) => new() { ScheduledAt = at, ReminderAt = at, Content = content.Trim(), Status = string.IsNullOrWhiteSpace(content) ? ImportStatus.内容为空 : ImportStatus.可导入, ErrorReason = string.IsNullOrWhiteSpace(content) ? "安排内容为空。" : "" };
    private static void Normalize(List<ImportCandidate> rows)
    {
        foreach (var row in rows)
        {
            row.Content = row.Content.Trim();
            if (row.ScheduledAt is null && row.Status != ImportStatus.需要确认) { row.Status = ImportStatus.日期无效; row.ErrorReason = "日期无效。"; }
            else if (string.IsNullOrWhiteSpace(row.Content)) { row.Status = ImportStatus.内容为空; row.ErrorReason = "安排内容为空。"; }
            else if (row.ReminderAt is null && row.ScheduledAt.HasValue) row.ReminderAt = row.ScheduledAt;
        }
    }
}
