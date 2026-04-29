using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Pictures
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "pictures-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Pictures";

            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Description");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue("Sample product with logo");
            sheet.Cells["A4"].PutValue("Item 2");
            sheet.Cells["B4"].PutValue("Another product");

            var sampleImageData = CreateSamplePngImage();

            var picture1Index = sheet.Pictures.Add(1, 2, 3, 3, sampleImageData);
            var picture1 = sheet.Pictures[picture1Index];
            picture1.Name = "Product Logo 1";

            var picture2Index = sheet.Pictures.Add(3, 2, 5, 3, sampleImageData);
            var picture2 = sheet.Pictures[picture2Index];
            picture2.Name = "Product Logo 2";
            picture2.UpperLeftColumn = 4;
            picture2.LowerRightColumn = 5;

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Pictures"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Picture count: " + loadedSheet.Pictures.Count);
            Console.WriteLine("First picture: " + loadedSheet.Pictures[0].Name + " (Type: " + loadedSheet.Pictures[0].ImageType + ")");
            Console.WriteLine("First picture anchor: R" + loadedSheet.Pictures[0].UpperLeftRow + "C" + loadedSheet.Pictures[0].UpperLeftColumn + " to R" + loadedSheet.Pictures[0].LowerRightRow + "C" + loadedSheet.Pictures[0].LowerRightColumn);
            Console.WriteLine("Second picture: " + loadedSheet.Pictures[1].Name + " (Type: " + loadedSheet.Pictures[1].ImageType + ")");
        }

        private static byte[] CreateSamplePngImage()
        {
            using (var ms = new MemoryStream())
            {
                var pngHeader = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                ms.Write(pngHeader, 0, pngHeader.Length);

                var ihdr = new byte[] { 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00 };
                WriteChunk(ms, "IHDR", ihdr);

                var idat = new byte[] { 0x78, 0x9C, 0x62, 0x00, 0x02, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D };
                WriteChunk(ms, "IDAT", idat);

                WriteChunk(ms, "IEND", new byte[0]);

                return ms.ToArray();
            }
        }

        private static void WriteChunk(MemoryStream ms, string type, byte[] data)
        {
            var lengthBytes = BitConverter.GetBytes(data.Length);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(lengthBytes);
            }

            ms.Write(lengthBytes, 0, 4);

            var typeBytes = System.Text.Encoding.ASCII.GetBytes(type);
            ms.Write(typeBytes, 0, 4);

            if (data.Length > 0)
            {
                ms.Write(data, 0, data.Length);
            }

            var crc = CalculateCrc(typeBytes, data);
            var crcBytes = BitConverter.GetBytes(crc);
            if (BitConverter.IsLittleEndian)
            {
                Array.Reverse(crcBytes);
            }
            ms.Write(crcBytes, 0, 4);
        }

        private static uint CalculateCrc(byte[] typeBytes, byte[] data)
        {
            var crc = 0xFFFFFFFF;
            var buffer = new byte[typeBytes.Length + data.Length];
            Buffer.BlockCopy(typeBytes, 0, buffer, 0, typeBytes.Length);
            Buffer.BlockCopy(data, 0, buffer, typeBytes.Length, data.Length);

            for (var i = 0; i < buffer.Length; i++)
            {
                crc ^= buffer[i];
                for (var j = 0; j < 8; j++)
                {
                    if ((crc & 1) != 0)
                    {
                        crc = (crc >> 1) ^ 0xEDB88320;
                    }
                    else
                    {
                        crc >>= 1;
                    }
                }
            }

            return ~crc;
        }
    }
}
