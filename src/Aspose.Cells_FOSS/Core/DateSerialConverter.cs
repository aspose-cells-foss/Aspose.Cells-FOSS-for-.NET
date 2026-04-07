namespace Aspose.Cells_FOSS.Core;

public static class DateSerialConverter
{
    private static readonly DateTime Windows1900Epoch = new DateTime(1899, 12, 31);
    private static readonly DateTime Mac1904Epoch = new DateTime(1904, 1, 1);
    private const double TicksPerDay = TimeSpan.TicksPerDay;

    public static double ToSerial(DateTime value, DateSystem dateSystem)
    {
        var baseDate = dateSystem == DateSystem.Mac1904 ? Mac1904Epoch : Windows1900Epoch;
        var serial = (value.Ticks - baseDate.Ticks) / TicksPerDay;
        if (dateSystem == DateSystem.Windows1900 && serial >= 60d)
        {
            serial += 1d;
        }

        return serial;
    }

    public static DateTime FromSerial(double serial, DateSystem dateSystem)
    {
        if (dateSystem == DateSystem.Windows1900 && serial >= 60d)
        {
            serial -= 1d;
        }

        var baseDate = dateSystem == DateSystem.Mac1904 ? Mac1904Epoch : Windows1900Epoch;
        var ticks = (long)Math.Round(serial * TicksPerDay, MidpointRounding.AwayFromZero);
        return new DateTime(baseDate.Ticks + ticks, DateTimeKind.Unspecified);
    }
}

