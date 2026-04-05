namespace GenerateDeliveryReports.Data.Extensions;

public static class StringExtensions
{
    public static string ToEmptyStringIfNull(this object? value)
    {
        return value == null ? string.Empty : value.ToString()!;
    }

    public static string ToEmptyStringIfNull(this string? value)
    {
        return string.IsNullOrEmpty(value) ? string.Empty : value;
    }

    public static string[] RemoveEmptyLines(this string[]? lines)
    {
        if (lines == null || lines.Length == 0)
            return lines ?? [];

        return lines.Where(line => !string.IsNullOrEmpty(line)).ToArray();
    }
}
