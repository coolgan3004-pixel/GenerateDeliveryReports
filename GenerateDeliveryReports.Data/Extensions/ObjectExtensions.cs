namespace GenerateDeliveryReports.Data.Extensions;

public static class ObjectExtensions
{
    public static int ToInt(this object? obj)
    {
        if (obj == null) return 0;
        return int.TryParse(obj.ToString(), out int result) ? result : 0;
    }

    public static long ToLong(this object? obj)
    {
        if (obj == null) return 0;
        return long.TryParse(obj.ToString(), out long result) ? result : 0;
    }

    public static double ToDouble(this object? obj)
    {
        if (obj == null) return 0;
        return double.TryParse(obj.ToString(), out double result) ? result : 0;
    }
}
