namespace Aspose.Cells_FOSS.Rendering
{
    /// <summary>Maps an OOXML paper-size code to portrait dimensions in points.</summary>
    internal static class PaperSizes
    {
        public static void GetDimensionsPoints(int paperSizeCode, out double widthPt, out double heightPt)
        {
            // Unspecified paper uses A4 unless saved printer settings supplied a paper code.
            double widthIn, heightIn;
            switch (paperSizeCode)
            {
                case 1:  widthIn = 8.5;  heightIn = 11.0;  break; // Letter
                case 3:  widthIn = 11.0; heightIn = 17.0;  break; // Tabloid
                case 5:  widthIn = 8.5;  heightIn = 14.0;  break; // Legal
                case 8:  widthIn = 11.69; heightIn = 16.54; break; // A3
                case 9:  widthIn = 210d / 25.4d; heightIn = 297d / 25.4d; break; // A4
                case 11: widthIn = 5.83; heightIn = 8.27;  break; // A5
                case 13: widthIn = 7.17; heightIn = 10.12; break; // B5 (JIS)
                default: widthIn = 210d / 25.4d; heightIn = 297d / 25.4d; break;
            }
            widthPt = RenderUnits.InchesToPoints(widthIn);
            heightPt = RenderUnits.InchesToPoints(heightIn);
        }
    }
}
