using System.Globalization;
using System.Text.RegularExpressions;

namespace CalendarReminder.Core;

public static partial class DateTimeFormat
{
    public const string DisplayPattern = "yyyy/MM/dd HH:mm";
    public static string Format(DateTime value) => value.ToString(DisplayPattern, CultureInfo.InvariantCulture);

    private static readonly string[] ExactFormats =
    [
        "yyyy/MM/dd HH:mm", "yyyy/M/d H:mm", "yyyy-MM-dd HH:mm", "yyyy-M-d H:mm",
        "yyyy.MM.dd HH:mm", "yyyy.M.d H:mm", "yyyy年M月d日 H:mm", "yyyy年M月d日 H点m分",
        "yyyy年MM月dd日 HH点mm分", "yyyy年M月d日H点m分", "yyyy年M月d日H:mm"
    ];

    public static bool TryParse(string? input, out DateTime value)
    {
        value = default;
        if (string.IsNullOrWhiteSpace(input)) return false;
        var normalized = input.Trim().Replace('：', ':');
        if (DateTime.TryParseExact(normalized, ExactFormats, CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces, out value)) return true;
        var match = DateRegex().Match(normalized);
        if (!match.Success) return false;
        return int.TryParse(match.Groups["y"].Value, out var y) &&
               int.TryParse(match.Groups["m"].Value, out var m) &&
               int.TryParse(match.Groups["d"].Value, out var d) &&
               int.TryParse(match.Groups["h"].Value, out var h) &&
               int.TryParse(match.Groups["min"].Value, out var min) &&
               TryCreate(y, m, d, h, min, out value);
    }

    public static bool TryExtract(string input, out DateTime value, out string remaining)
    {
        value = default; remaining = input.Trim();
        var match = DateRegex().Match(input.Replace('：', ':'));
        if (!match.Success || !TryParse(match.Value, out value)) return false;
        remaining = (input[..match.Index] + " " + input[(match.Index + match.Length)..]).Trim(' ', '-', '—', ':', '：', '|');
        return true;
    }

    private static bool TryCreate(int y, int m, int d, int h, int min, out DateTime value)
    {
        value = default;
        try { value = new DateTime(y, m, d, h, min, 0, DateTimeKind.Local); return true; }
        catch (ArgumentOutOfRangeException) { return false; }
    }

    [GeneratedRegex(@"(?<y>\d{4})\s*(?:[/\-.年])\s*(?<m>\d{1,2})\s*(?:[/\-.月])\s*(?<d>\d{1,2})\s*(?:日)?\s+(?<h>\d{1,2})\s*(?:[:点时])\s*(?<min>\d{1,2})\s*(?:分)?", RegexOptions.CultureInvariant)]
    private static partial Regex DateRegex();
}
