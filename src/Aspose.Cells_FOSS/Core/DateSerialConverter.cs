using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Provides date serial converter operations.
    /// </summary>
    public static class DateSerialConverter
    {
        private static readonly DateTime Windows1900Epoch = new DateTime(1899, 12, 31);
        private static readonly DateTime Mac1904Epoch = new DateTime(1904, 1, 1);
        private const double TicksPerDay = TimeSpan.TicksPerDay;

        /// <summary>
        /// Performs to serial.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="dateSystem">The date system.</param>
        /// <returns>The double.</returns>
        public static double ToSerial(DateTime value, DateSystem dateSystem)
        {
            var baseDate = dateSystem == DateSystem.Mac1904 ? Mac1904Epoch : Windows1900Epoch;
            var serial = (value.Ticks - baseDate.Ticks) / TicksPerDay;
            // Excel's 1900 date system preserves the historic leap-year bug, so serials
            // on or after the phantom 1900-02-29 are shifted by one day.
            if (dateSystem == DateSystem.Windows1900 && serial >= 60d)
            {
                serial += 1d;
            }

            return serial;
        }

        /// <summary>
        /// Performs from serial.
        /// </summary>
        /// <param name="serial">The serial.</param>
        /// <param name="dateSystem">The date system.</param>
        /// <returns>The date time.</returns>
        public static DateTime FromSerial(double serial, DateSystem dateSystem)
        {
            // Reverse the 1900-system compatibility adjustment before converting back
            // to a .NET DateTime.
            if (dateSystem == DateSystem.Windows1900 && serial >= 60d)
            {
                serial -= 1d;
            }

            var baseDate = dateSystem == DateSystem.Mac1904 ? Mac1904Epoch : Windows1900Epoch;
            var ticks = (long)Math.Round(serial * TicksPerDay, MidpointRounding.AwayFromZero);
            return new DateTime(baseDate.Ticks + ticks, DateTimeKind.Unspecified);
        }
    }
}
