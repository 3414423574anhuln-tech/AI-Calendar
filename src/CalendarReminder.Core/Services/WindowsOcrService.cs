using System.Text;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage;
using Windows.Storage.Streams;

namespace CalendarReminder.Core.Services;

public interface IOcrService
{
    Task<string> ReadImageAsync(string path);
    Task<string> ReadPdfAsync(string path, IProgress<int>? progress = null);
}

public sealed class WindowsOcrService : IOcrService
{
    public async Task<string> ReadImageAsync(string path)
    {
        var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(path));
        using var stream = await file.OpenReadAsync();
        var decoder = await BitmapDecoder.CreateAsync(stream);
        using var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
        return await RecognizeAsync(bitmap);
    }

    public async Task<string> ReadPdfAsync(string path, IProgress<int>? progress = null)
    {
        var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(path));
        using var input = await file.OpenReadAsync();
        var pdf = await PdfDocument.LoadFromStreamAsync(input); var result = new StringBuilder();
        for (uint i = 0; i < pdf.PageCount; i++)
        {
            using var page = pdf.GetPage(i); using var rendered = new InMemoryRandomAccessStream();
            await page.RenderToStreamAsync(rendered); rendered.Seek(0);
            var decoder = await BitmapDecoder.CreateAsync(rendered);
            using var bitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
            result.AppendLine(await RecognizeAsync(bitmap)); progress?.Report(30 + (int)(65 * (i + 1) / pdf.PageCount));
        }
        return result.ToString();
    }

    private static async Task<string> RecognizeAsync(SoftwareBitmap bitmap)
    {
        var engine = OcrEngine.TryCreateFromLanguage(new Windows.Globalization.Language("zh-Hans")) ?? OcrEngine.TryCreateFromUserProfileLanguages();
        if (engine is null) throw new NotSupportedException("本机未安装可用的 OCR 语言包，请在 Windows 语言设置中安装中文简体的基本键入功能。");
        var result = await engine.RecognizeAsync(bitmap); return result.Text;
    }
}
