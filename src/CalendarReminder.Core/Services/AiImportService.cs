using CalendarReminder.Core.Models;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using ExcelDataReader;
using System.Data;
using System.Net;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using UglyToad.PdfPig;
using A = DocumentFormat.OpenXml.Drawing;

namespace CalendarReminder.Core.Services;

public enum AiProvider { DeepSeek, Kimi, OpenAI, Anthropic }
public sealed record AiModelOption(string Id, string DisplayName);
public sealed record AiProviderDefinition(AiProvider Provider, string DisplayName, string Endpoint, IReadOnlyList<AiModelOption> Models)
{
    public override string ToString() => DisplayName;
}

public static class AiProviderCatalog
{
    public static IReadOnlyList<AiProviderDefinition> All { get; } =
    [
        new(AiProvider.DeepSeek, "DeepSeek", "https://api.deepseek.com/chat/completions",
            [new("deepseek-v4-flash", "DeepSeek V4 Flash"), new("deepseek-v4-pro", "DeepSeek V4 Pro")]),
        new(AiProvider.Kimi, "Kimi（月之暗面）", "https://api.moonshot.cn/v1/chat/completions",
            [new("kimi-k3", "Kimi K3"), new("kimi-k2.6", "Kimi K2.6"), new("kimi-k2.7-code", "Kimi K2.7 Code")]),
        new(AiProvider.OpenAI, "OpenAI", "https://api.openai.com/v1/chat/completions",
            [new("gpt-5.6-sol", "GPT-5.6 Sol"), new("gpt-5.6-terra", "GPT-5.6 Terra"), new("gpt-5.6-luna", "GPT-5.6 Luna")]),
        new(AiProvider.Anthropic, "Anthropic（Claude）", "https://api.anthropic.com/v1/messages",
            [new("claude-fable-5", "Claude Fable 5"), new("claude-opus-5", "Claude Opus 5"), new("claude-sonnet-5", "Claude Sonnet 5"), new("claude-haiku-4-5-20251001", "Claude Haiku 4.5")])
    ];
    public static AiProviderDefinition Get(AiProvider provider) => All.Single(x => x.Provider == provider);
}

public sealed class AiApiSettings
{
    public AiProvider Provider { get; set; } = AiProvider.DeepSeek;
    public string Model { get; set; } = "deepseek-v4-flash";
}

public sealed class AiApiSettingsService
{
    public AiApiSettings Load()
    {
        try
        {
            if (!File.Exists(AppPaths.AiSettingsPath)) return new();
            var value = JsonSerializer.Deserialize<AiApiSettings>(File.ReadAllText(AppPaths.AiSettingsPath));
            if (value is null || !AiProviderCatalog.Get(value.Provider).Models.Any(x => x.Id == value.Model)) return new();
            return value;
        }
        catch (Exception ex) { AppLog.Error("读取 API 配置", ex); return new(); }
    }
    public void Save(AiApiSettings value)
    {
        AppPaths.Ensure();
        File.WriteAllText(AppPaths.AiSettingsPath, JsonSerializer.Serialize(value, new JsonSerializerOptions { WriteIndented = true }));
    }
}

public sealed class WindowsCredentialStore
{
    private const uint Generic = 1, PersistLocalMachine = 2;
    private static string Target(AiProvider provider) => $"CalendarReminder.ApiKey.{provider}";

    public void Save(AiProvider provider, string secret)
    {
        if (string.IsNullOrWhiteSpace(secret)) throw new ArgumentException("API Key 不能为空。");
        var bytes = Encoding.Unicode.GetBytes(secret.Trim());
        if (bytes.Length > 5120) throw new ArgumentException("API Key 长度异常。");
        var pointer = Marshal.AllocCoTaskMem(bytes.Length);
        try
        {
            Marshal.Copy(bytes, 0, pointer, bytes.Length);
            var credential = new Credential { Type = Generic, TargetName = Target(provider), CredentialBlobSize = (uint)bytes.Length, CredentialBlob = pointer, Persist = PersistLocalMachine, UserName = Environment.UserName };
            if (!CredWrite(ref credential, 0)) throw new InvalidOperationException($"Windows 凭据保存失败（{Marshal.GetLastWin32Error()}）。");
        }
        finally
        {
            Marshal.Copy(new byte[bytes.Length], 0, pointer, bytes.Length);
            Marshal.FreeCoTaskMem(pointer);
            System.Security.Cryptography.CryptographicOperations.ZeroMemory(bytes);
        }
    }

    public string? Read(AiProvider provider)
    {
        if (!CredRead(Target(provider), Generic, 0, out var pointer))
        {
            var error = Marshal.GetLastWin32Error();
            if (error == 1168) return null;
            throw new InvalidOperationException($"Windows 凭据读取失败（{error}）。");
        }
        try
        {
            var credential = Marshal.PtrToStructure<Credential>(pointer);
            return credential.CredentialBlob == IntPtr.Zero ? null : Marshal.PtrToStringUni(credential.CredentialBlob, (int)credential.CredentialBlobSize / 2);
        }
        finally { CredFree(pointer); }
    }

    public void Delete(AiProvider provider)
    {
        if (!CredDelete(Target(provider), Generic, 0) && Marshal.GetLastWin32Error() != 1168)
            throw new InvalidOperationException($"Windows 凭据删除失败（{Marshal.GetLastWin32Error()}）。");
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct Credential
    {
        public uint Flags, Type;
        public string TargetName;
        public string? Comment;
        public long LastWritten;
        public uint CredentialBlobSize;
        public IntPtr CredentialBlob;
        public uint Persist, AttributeCount;
        public IntPtr Attributes;
        public string? TargetAlias;
        public string UserName;
    }
    [DllImport("advapi32.dll", EntryPoint = "CredWriteW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredWrite(ref Credential credential, uint flags);
    [DllImport("advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredRead(string target, uint type, uint flags, out IntPtr credential);
    [DllImport("advapi32.dll", EntryPoint = "CredDeleteW", CharSet = CharSet.Unicode, SetLastError = true)] private static extern bool CredDelete(string target, uint type, uint flags);
    [DllImport("advapi32.dll")] private static extern void CredFree(IntPtr credential);
}

public sealed record AiImportResult(string ExcelPath, List<ImportCandidate> Candidates, int ExistingScheduleCount, int SemanticOverlapCount);

public sealed class AiImportService
{
    private const int SourceChunkSize = 55_000;
    private const int ExistingBatchSize = 120;
    private readonly HttpClient _http;
    private readonly IOcrService _ocr;
    public AiImportService(HttpClient? http = null, IOcrService? ocr = null)
    {
        _http = http ?? new HttpClient { Timeout = TimeSpan.FromMinutes(4) };
        _ocr = ocr ?? new WindowsOcrService();
    }

    public async Task TestConnectionAsync(AiProvider provider, string model, string apiKey, CancellationToken token = default)
        => _ = await CompleteJsonAsync(provider, model, apiKey, "你是连接测试助手。", "只返回 JSON：{\"ok\":true}", token);

    public async Task<AiImportResult> AnalyzeAsync(string path, AiProvider provider, string model, string apiKey,
        IReadOnlyList<ScheduleItem> existing, IProgress<string>? progress = null, CancellationToken token = default)
    {
        if (!File.Exists(path)) throw new FileNotFoundException("选择的文件不存在。", path);
        if (!AiProviderCatalog.Get(provider).Models.Any(x => x.Id == model)) throw new InvalidOperationException("所选模型不属于该供应商。");
        progress?.Report("正在本地读取文件…");
        var source = await ExtractSourceAsync(path, token);
        if (string.IsNullOrWhiteSpace(source)) throw new InvalidDataException("文件中没有可供 AI 分析的文字。");

        var candidates = new List<ImportCandidate>();
        var chunks = SplitText(source, SourceChunkSize).ToList();
        for (var i = 0; i < chunks.Count; i++)
        {
            progress?.Report($"正在进行 AI 识别（{i + 1}/{chunks.Count}）…");
            var json = await CompleteJsonAsync(provider, model, apiKey, ExtractionSystemPrompt,
                $"今天是 {DateTime.Today:yyyy/MM/dd}。请分析以下文件片段并严格返回 JSON。不要猜测缺失的年月日。\n\n{chunks[i]}", token);
            candidates.AddRange(ParseCandidatesJson(json));
        }
        candidates = candidates.GroupBy(x => new { x.ScheduledAt, Content = x.Content.Trim() }).Select(x => x.First()).ToList();
        if (candidates.Count == 0) throw new InvalidDataException("大模型未返回可预览的日程记录。");

        var overlapCount = 0;
        if (existing.Count > 0)
        {
            var candidateBatches = candidates.Select((x, i) => (Row: x, Id: $"N{i + 1}")).Chunk(50).ToList();
            var existingBatches = existing.Select((x, i) => (Row: x, Id: $"E{i + 1}")).Chunk(ExistingBatchSize).ToList();
            var total = candidateBatches.Count * existingBatches.Count; var done = 0;
            foreach (var newBatch in candidateBatches)
                foreach (var oldBatch in existingBatches)
                {
                    progress?.Report($"正在与全部已有日程语义比对（{++done}/{total}）…");
                    var request = BuildOverlapRequest(newBatch, oldBatch);
                    var response = await CompleteJsonAsync(provider, model, apiKey, OverlapSystemPrompt, request, token);
                    foreach (var match in ParseOverlapIds(response))
                    {
                        var tuple = newBatch.FirstOrDefault(x => x.Id == match.Id);
                        if (tuple.Row is null || tuple.Row.Status == ImportStatus.高度重合) continue;
                        tuple.Row.Status = ImportStatus.高度重合; tuple.Row.IsSelected = false; tuple.Row.IsDuplicate = true;
                        tuple.Row.ErrorReason = "AI 判定与已有日程高度重合，已禁止导入。" + (string.IsNullOrWhiteSpace(match.Reason) ? "" : $" {match.Reason}");
                        overlapCount++;
                    }
                }
        }

        progress?.Report("正在生成规范 Excel…");
        var excelPath = WriteNormalizedExcel(candidates);
        progress?.Report("AI 识别完成，正在打开预览…");
        return new(excelPath, candidates, existing.Count, overlapCount);
    }

    public static List<ImportCandidate> ParseCandidatesJson(string input)
    {
        using var doc = JsonDocument.Parse(CleanJson(input));
        var root = doc.RootElement;
        var events = root.ValueKind == JsonValueKind.Array ? root : root.TryGetProperty("events", out var value) ? value : default;
        if (events.ValueKind != JsonValueKind.Array) throw new InvalidDataException("大模型响应中缺少 events 数组。");
        var result = new List<ImportCandidate>();
        foreach (var item in events.EnumerateArray())
        {
            var content = GetString(item, "content").Trim();
            var scheduledText = GetString(item, "scheduledAt");
            var dateText = GetString(item, "date");
            var timeText = GetString(item, "time");
            var reason = GetString(item, "reason");
            var rangeStart = GetString(item, "rangeStart"); var rangeEnd = GetString(item, "rangeEnd"); var weekText = GetString(item, "week");
            var rangeExpression = !string.IsNullOrWhiteSpace(rangeStart) && !string.IsNullOrWhiteSpace(rangeEnd) ? $"{rangeStart}至{rangeEnd}" : !string.IsNullOrWhiteSpace(weekText) ? weekText : dateText;
            var expanded = ImportService.ParseText($"{rangeExpression} {content}");
            if (expanded.Count > 1 || (!string.IsNullOrWhiteSpace(rangeStart) && expanded.Count == 1) || (!string.IsNullOrWhiteSpace(weekText) && expanded.Count == 1))
            {
                foreach (var row in expanded) row.ErrorReason = JoinReason(reason, row.ErrorReason);
                result.AddRange(expanded); continue;
            }
            DateTime? scheduled = null; DateTime? reminder = null;
            if (DateTimeFormat.TryParse(scheduledText, out var exact)) { scheduled = exact; reminder = exact; }
            else if (DateTime.TryParseExact(dateText, ["yyyy/MM/dd", "yyyy-M-d", "yyyy/MM/d"], null, System.Globalization.DateTimeStyles.None, out var date))
            {
                if (TimeSpan.TryParseExact(timeText, ["hh\\:mm", "h\\:mm"], null, out var time)) { scheduled = date.Date.Add(time); reminder = scheduled; }
                else { scheduled = date.Date.AddHours(7).AddMinutes(30); reminder = date.Date.AddDays(-1).AddHours(7).AddMinutes(30); reason = JoinReason(reason, "未识别到具体时间，已默认日程 07:30、提醒为前一天 07:30。"); }
            }
            else if (!string.IsNullOrWhiteSpace(dateText))
            {
                var inferred = ImportService.ParseText($"{dateText} {content}").FirstOrDefault();
                if (inferred is not null)
                {
                    scheduled = inferred.ScheduledAt; reminder = inferred.ReminderAt;
                    reason = JoinReason(reason, inferred.ErrorReason);
                }
            }
            var statusText = GetString(item, "status");
            var status = scheduled is null || statusText.Contains("confirm", StringComparison.OrdinalIgnoreCase) || statusText.Contains("确认") ? ImportStatus.需要确认 : ImportStatus.可导入;
            if (string.IsNullOrWhiteSpace(content)) status = ImportStatus.内容为空;
            result.Add(new ImportCandidate { ScheduledAt = scheduled, ReminderAt = reminder, Content = content, Status = status, ErrorReason = reason, IsSelected = status == ImportStatus.可导入 });
        }
        return result;
    }

    private async Task<string> ExtractSourceAsync(string path, CancellationToken token)
    {
        var ext = Path.GetExtension(path).ToLowerInvariant();
        return ext switch
        {
            ".xls" or ".xlsx" => await Task.Run(() => ExtractExcel(path), token),
            ".docx" => await Task.Run(() => ExtractWord(path), token),
            ".pptx" => await Task.Run(() => ExtractPowerPoint(path), token),
            ".pdf" => await ExtractPdfAsync(path),
            ".png" or ".jpg" or ".jpeg" => await _ocr.ReadImageAsync(path),
            _ => throw new NotSupportedException("AI 导入不支持此文件类型。")
        };
    }

    private static string ExtractExcel(string path)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = ExcelReaderFactory.CreateReader(stream);
        var data = reader.AsDataSet(); var text = new StringBuilder();
        foreach (DataTable table in data.Tables)
        {
            text.AppendLine($"[工作表:{table.TableName}]");
            for (var r = 0; r < table.Rows.Count; r++)
            {
                var values = table.Rows[r].ItemArray.Select((x, c) => string.IsNullOrWhiteSpace(x?.ToString()) ? null : $"C{c + 1}={x}").Where(x => x is not null);
                var line = string.Join("\t", values!); if (!string.IsNullOrWhiteSpace(line)) text.AppendLine($"R{r + 1}\t{line}");
            }
        }
        return text.ToString();
    }
    private static string ExtractWord(string path)
    {
        using var doc = WordprocessingDocument.Open(path, false);
        var body = doc.MainDocumentPart?.Document.Body ?? throw new InvalidDataException("Word 文档没有正文。");
        return string.Join("\n", body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Select(x => x.InnerText).Where(x => !string.IsNullOrWhiteSpace(x)));
    }
    private static string ExtractPowerPoint(string path)
    {
        using var doc = PresentationDocument.Open(path, false); var text = new StringBuilder(); var page = 0;
        foreach (var slide in doc.PresentationPart?.SlideParts ?? [])
        {
            text.AppendLine($"[幻灯片:{++page}]");
            foreach (var line in slide.Slide.Descendants<A.Paragraph>().Select(x => string.Concat(x.Descendants<A.Text>().Select(t => t.Text))).Where(x => !string.IsNullOrWhiteSpace(x))) text.AppendLine(line);
        }
        return text.ToString();
    }
    private async Task<string> ExtractPdfAsync(string path)
    {
        using var pdf = PdfDocument.Open(path); var text = string.Join("\n", pdf.GetPages().Select(x => x.Text));
        return text.Count(char.IsLetterOrDigit) >= Math.Max(12, pdf.NumberOfPages * 6) ? text : await _ocr.ReadPdfAsync(path);
    }

    private async Task<string> CompleteJsonAsync(AiProvider provider, string model, string apiKey, string system, string user, CancellationToken token)
    {
        var definition = AiProviderCatalog.Get(provider);
        using var request = new HttpRequestMessage(HttpMethod.Post, definition.Endpoint);
        object payload;
        if (provider == AiProvider.Anthropic)
        {
            request.Headers.Add("x-api-key", apiKey.Trim()); request.Headers.Add("anthropic-version", "2023-06-01");
            payload = new { model, max_tokens = 8192, system, messages = new[] { new { role = "user", content = user } } };
        }
        else
        {
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey.Trim());
            payload = new { model, messages = new[] { new { role = "system", content = system }, new { role = "user", content = user } }, response_format = new { type = "json_object" } };
        }
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");
        HttpResponseMessage response;
        try { response = await _http.SendAsync(request, token); }
        catch (TaskCanceledException) when (!token.IsCancellationRequested) { throw new TimeoutException("连接模型服务超时，请检查网络后重试。"); }
        catch (HttpRequestException ex) { throw new InvalidOperationException("无法连接模型服务，请检查网络、代理和供应商地址。", ex); }
        using (response)
        {
            var body = await response.Content.ReadAsStringAsync(token);
            if (!response.IsSuccessStatusCode) throw CreateHttpError(response.StatusCode);
            using var doc = JsonDocument.Parse(body);
            if (provider == AiProvider.Anthropic)
            {
                foreach (var part in doc.RootElement.GetProperty("content").EnumerateArray()) if (GetString(part, "type") == "text") return GetString(part, "text");
            }
            else return doc.RootElement.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? "";
        }
        throw new InvalidDataException("模型返回了空响应。");
    }

    private static Exception CreateHttpError(HttpStatusCode status) => status switch
    {
        HttpStatusCode.BadRequest => new InvalidOperationException("模型拒绝了请求，请确认模型名称及账号权限。"),
        HttpStatusCode.Unauthorized => new InvalidOperationException("API Key 无效或已过期。"),
        HttpStatusCode.Forbidden => new InvalidOperationException("账号无权使用该模型，请检查供应商控制台权限。"),
        HttpStatusCode.NotFound => new InvalidOperationException("模型或 API 地址不存在，可能已下线。"),
        (HttpStatusCode)429 => new InvalidOperationException("请求过于频繁或账户额度不足，请稍后重试并检查余额。"),
        >= HttpStatusCode.InternalServerError => new InvalidOperationException("模型服务暂时异常，请稍后重试。"),
        _ => new InvalidOperationException($"模型服务请求失败（HTTP {(int)status}）。")
    };

    private static string WriteNormalizedExcel(IReadOnlyList<ImportCandidate> rows)
    {
        AppPaths.Ensure(); var path = Path.Combine(AppPaths.AiImportDirectory, $"AI识别_{DateTime.Now:yyyyMMdd_HHmmss_fff}.xlsx");
        using var wb = new XLWorkbook(); var ws = wb.AddWorksheet("AI识别结果");
        var headers = new[] { "序号", "日期时间", "提醒时间", "安排", "识别状态", "说明" };
        for (var c = 0; c < headers.Length; c++) ws.Cell(1, c + 1).Value = headers[c];
        for (var r = 0; r < rows.Count; r++)
        {
            var row = rows[r]; ws.Cell(r + 2, 1).Value = r + 1; ws.Cell(r + 2, 2).Value = row.DisplayTime; ws.Cell(r + 2, 3).Value = row.DisplayReminderTime;
            ws.Cell(r + 2, 4).Value = row.Content; ws.Cell(r + 2, 5).Value = row.Status.ToString(); ws.Cell(r + 2, 6).Value = row.ErrorReason;
        }
        ws.Row(1).Style.Font.Bold = true; ws.SheetView.FreezeRows(1); ws.Columns().AdjustToContents(); ws.Column(4).Width = Math.Min(70, Math.Max(25, ws.Column(4).Width)); ws.Column(6).Width = Math.Min(55, Math.Max(20, ws.Column(6).Width)); ws.Style.Alignment.WrapText = true;
        wb.SaveAs(path); return path;
    }

    private static string BuildOverlapRequest(IEnumerable<(ImportCandidate Row, string Id)> candidates, IEnumerable<(ScheduleItem Row, string Id)> existing)
    {
        var fresh = candidates.Select(x => new { id = x.Id, scheduledAt = x.Row.DisplayTime, reminderAt = x.Row.DisplayReminderTime, content = x.Row.Content });
        var old = existing.Select(x => new { id = x.Id, scheduledAt = DateTimeFormat.Format(x.Row.ScheduledAt), reminderAt = DateTimeFormat.Format(x.Row.ReminderAt), content = x.Row.Content });
        return "待导入：\n" + JsonSerializer.Serialize(fresh) + "\n已有日程：\n" + JsonSerializer.Serialize(old);
    }
    public static IReadOnlyList<(string Id, string Reason)> ParseOverlapIds(string input)
    {
        using var doc = JsonDocument.Parse(CleanJson(input)); if (!doc.RootElement.TryGetProperty("overlaps", out var values) || values.ValueKind != JsonValueKind.Array) return [];
        return values.EnumerateArray().Select(x => (GetString(x, "candidateId"), GetString(x, "reason"))).Where(x => !string.IsNullOrWhiteSpace(x.Item1)).ToList();
    }
    private static IEnumerable<string> SplitText(string text, int size)
    {
        if (text.Length <= size) { yield return text; yield break; }
        var position = 0;
        while (position < text.Length) { var length = Math.Min(size, text.Length - position); if (position + length < text.Length) { var newline = text.LastIndexOf('\n', position + length - 1, length); if (newline > position) length = newline - position + 1; } yield return text.Substring(position, length); position += length; }
    }
    private static string CleanJson(string text)
    {
        var value = text.Trim(); if (value.StartsWith("```")) { var first = value.IndexOf('\n'); var last = value.LastIndexOf("```", StringComparison.Ordinal); if (first >= 0 && last > first) value = value[(first + 1)..last].Trim(); }
        var start = Math.Min(value.IndexOf('{') is var o && o >= 0 ? o : int.MaxValue, value.IndexOf('[') is var a && a >= 0 ? a : int.MaxValue); return start == int.MaxValue ? value : value[start..];
    }
    private static string GetString(JsonElement element, string name) => element.TryGetProperty(name, out var value) && value.ValueKind != JsonValueKind.Null ? value.ToString() : "";
    private static string JoinReason(string a, string b) => string.IsNullOrWhiteSpace(a) ? b : $"{a} {b}";

    private const string ExtractionSystemPrompt = """
你是中文日历数据整理器。识别真实的行程、会议、拜访、截止事项和明确计划，合并跨单元格语义，忽略纯标题、表头、序号和没有安排意义的数字。只返回 JSON 对象：
{"events":[{"scheduledAt":"yyyy/MM/dd HH:mm 或空","date":"yyyy/MM/dd 或部分日期文字或空","time":"HH:mm 或空","rangeStart":"范围起始 yyyy/MM/dd 或空","rangeEnd":"范围结束 yyyy/MM/dd 或空","week":"yyyy年M月第N周或空","content":"完整安排","status":"certain 或 needs_confirmation","reason":"说明"}]}
规则：24 小时制；同一事项不要重复；必须在 JSON 中返回。明确年月日但没有时间时保留 date、time 留空。日期范围必须填写 rangeStart 和 rangeEnd，闭区间两端都保留，不要自行拆分。某月第N周填写 week。仅有年份时把 date 设为该年1月1日；仅有月份且能从文件确定年份时：月份/上旬/月初取1日，中旬取13日，下旬取19日，月底/月末取25日，并将 status 设为 certain；如果月份缺少年份则 date 保留原始“8月”等文字并设 needs_confirmation。不得猜测文件中不存在的年份。
""";
    private const string OverlapSystemPrompt = """
你是日程语义去重器。比较待导入日程与已有日程的 scheduledAt、reminderAt 和 content。只有当时间相同或高度接近，并且事项描述指向同一活动、同一客户/项目或实质相同任务时，才判定高度重合；仅主题相似、时间明显不同、周期性事项的不同实例不得判重。只返回 JSON：{"overlaps":[{"candidateId":"N1","existingId":"E1","reason":"简短依据"}]}。没有重合时返回空数组。不得修改 ID。
""";
}
