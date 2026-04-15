using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents workbook view model.
    /// </summary>
    public sealed class WorkbookViewModel
    {
        /// <summary>
        /// Gets or sets the x window.
        /// </summary>
        public int? XWindow { get; set; }
        /// <summary>
        /// Gets or sets the y window.
        /// </summary>
        public int? YWindow { get; set; }
        /// <summary>
        /// Gets or sets the window width.
        /// </summary>
        public int? WindowWidth { get; set; }
        /// <summary>
        /// Gets or sets the window height.
        /// </summary>
        public int? WindowHeight { get; set; }
        /// <summary>
        /// Gets or sets the first sheet.
        /// </summary>
        public int? FirstSheet { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether show horizontal scroll.
        /// </summary>
        public bool? ShowHorizontalScroll { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether show vertical scroll.
        /// </summary>
        public bool? ShowVerticalScroll { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether show sheet tabs.
        /// </summary>
        public bool? ShowSheetTabs { get; set; }
        /// <summary>
        /// Gets or sets the tab ratio.
        /// </summary>
        public int? TabRatio { get; set; }
        /// <summary>
        /// Gets or sets the visibility.
        /// </summary>
        public string Visibility { get; set; } = string.Empty;
        /// <summary>
        /// Gets or sets a value indicating whether minimized.
        /// </summary>
        public bool Minimized { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether auto filter date grouping.
        /// </summary>
        public bool AutoFilterDateGrouping { get; set; } = true;

        /// <summary>
        /// Copies values from the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        public void CopyFrom(WorkbookViewModel source)
        {
            XWindow = source.XWindow;
            YWindow = source.YWindow;
            WindowWidth = source.WindowWidth;
            WindowHeight = source.WindowHeight;
            FirstSheet = source.FirstSheet;
            ShowHorizontalScroll = source.ShowHorizontalScroll;
            ShowVerticalScroll = source.ShowVerticalScroll;
            ShowSheetTabs = source.ShowSheetTabs;
            TabRatio = source.TabRatio;
            Visibility = source.Visibility;
            Minimized = source.Minimized;
            AutoFilterDateGrouping = source.AutoFilterDateGrouping;
        }

        /// <summary>
        /// Performs has stored state.
        /// </summary>
        /// <param name="activeSheetIndex">The active sheet index.</param>
        /// <returns><see langword="true"/> if the condition is met; otherwise, <see langword="false"/>.</returns>
        public bool HasStoredState(int activeSheetIndex)
        {
            return activeSheetIndex > 0
                || XWindow.HasValue
                || YWindow.HasValue
                || WindowWidth.HasValue
                || WindowHeight.HasValue
                || FirstSheet.HasValue
                || ShowHorizontalScroll.HasValue
                || ShowVerticalScroll.HasValue
                || ShowSheetTabs.HasValue
                || TabRatio.HasValue
                || !string.IsNullOrEmpty(Visibility)
                || Minimized
                || !AutoFilterDateGrouping;
        }
    }
}
