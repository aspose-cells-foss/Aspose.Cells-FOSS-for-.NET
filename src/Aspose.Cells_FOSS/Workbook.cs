using System.IO;
using System.Collections.Generic;
using System;
using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS
{
    /// <summary>
    /// Represents the root spreadsheet object used to create, load, modify, and save an XLSX workbook.
    /// </summary>
    /// <example>
    /// <code>
    /// var workbook = new Workbook();
    /// var sheet = workbook.Worksheets[0];
    ///
    /// sheet.Cells["A1"].PutValue("Item");
    /// sheet.Cells["B1"].PutValue("Price");
    /// sheet.Cells["A2"].PutValue("Apple");
    /// sheet.Cells["B2"].PutValue(2.5);
    ///
    /// workbook.Settings.Culture = CultureInfo.GetCultureInfo("en-US");
    /// workbook.Save("report.xlsx");
    /// </code>
    /// </example>
    public class Workbook : IDisposable
    {
        private readonly WorkbookModel _model;
        private readonly WorksheetCollection _worksheets;
        private readonly WorkbookSettings _settings;
        private readonly WorkbookProperties _properties;
        private readonly DocumentProperties _documentProperties;
        private readonly DefinedNameCollection _definedNames;

        /// <summary>
        /// Initializes a new workbook with one default worksheet.
        /// </summary>
        public Workbook()
        {
            _model = new WorkbookModel();
            _worksheets = new WorksheetCollection(this);
            _settings = new WorkbookSettings(_model.Settings);
            _properties = new WorkbookProperties(_model);
            _documentProperties = new DocumentProperties(_model.DocumentProperties);
            _definedNames = new DefinedNameCollection(this);
            LoadDiagnostics = new LoadDiagnostics();
        }

        /// <summary>
        /// Opens an existing workbook from a file path using default load options.
        /// </summary>
        public Workbook(string fileName) : this(fileName, new LoadOptions()) { }

        /// <summary>
        /// Opens an existing workbook from a stream using default load options.
        /// </summary>
        public Workbook(Stream stream) : this(stream, new LoadOptions()) { }

        /// <summary>
        /// Opens an existing workbook from a file path using explicit load options.
        /// </summary>
        public Workbook(string fileName, LoadOptions options) : this()
        {
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            if (options == null) throw new ArgumentNullException(nameof(options));

            using (var stream = File.OpenRead(fileName))
            {
                LoadFromStream(stream, options);
            }
        }

        /// <summary>
        /// Opens an existing workbook from a stream using explicit load options.
        /// </summary>
        public Workbook(Stream stream, LoadOptions options) : this()
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) throw new ArgumentNullException(nameof(options));

            LoadFromStream(stream, options);
        }

        /// <summary>
        /// Gets the worksheets in workbook order.
        /// </summary>
        public WorksheetCollection Worksheets
        {
            get
            {
                return _worksheets;
            }
        }

        /// <summary>
        /// Gets workbook-level settings such as the date system and display culture.
        /// </summary>
        public WorkbookSettings Settings
        {
            get
            {
                return _settings;
            }
        }

        /// <summary>
        /// Gets workbook metadata and view settings exposed by the supported public API.
        /// </summary>
        public WorkbookProperties Properties
        {
            get
            {
                return _properties;
            }
        }

        /// <summary>
        /// Gets the document properties facade for core and extended metadata.
        /// </summary>
        public DocumentProperties DocumentProperties
        {
            get
            {
                return _documentProperties;
            }
        }

        /// <summary>
        /// Gets the workbook-defined names collection.
        /// </summary>
        public DefinedNameCollection DefinedNames
        {
            get
            {
                return _definedNames;
            }
        }

        /// <summary>
        /// Gets diagnostics collected while loading the current workbook.
        /// </summary>
        public LoadDiagnostics LoadDiagnostics { get; }

        internal WorkbookModel Model
        {
            get
            {
                return _model;
            }
        }

        internal void EnsureUniqueSheetName(string sheetName, WorksheetModel currentSheet = null)
        {
            for (var index = 0; index < _model.Worksheets.Count; index++)
            {
                var existing = _model.Worksheets[index];
                if (ReferenceEquals(existing, currentSheet))
                {
                    continue;
                }

                if (string.Equals(existing.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new CellsException("Worksheet name '" + sheetName + "' already exists.");
                }
            }
        }

        internal void EnsureValidDefinedNameScope(int? localSheetIndex)
        {
            if (!localSheetIndex.HasValue)
            {
                return;
            }

            if (localSheetIndex.Value < 0 || localSheetIndex.Value >= _model.Worksheets.Count)
            {
                throw new CellsException("Defined name scope must refer to an existing worksheet.");
            }
        }

        internal void EnsureUniqueDefinedName(DefinedNameModel currentDefinedName, string name, int? localSheetIndex)
        {
            for (var index = 0; index < _model.DefinedNames.Count; index++)
            {
                var existing = _model.DefinedNames[index];
                if (ReferenceEquals(existing, currentDefinedName))
                {
                    continue;
                }

                if (!string.Equals(existing.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (DefinedNameUtility.SameScope(existing.LocalSheetIndex, localSheetIndex))
                {
                    throw new CellsException("Defined name '" + name + "' already exists in the same scope.");
                }
            }
        }

        /// <summary>
        /// Saves the workbook to an XLSX file using default save options.
        /// </summary>
        public void Save(string fileName)
        {
            Save(fileName, new SaveOptions());
        }

        /// <summary>
        /// Saves the workbook to a file using the specified save format.
        /// </summary>
        public void Save(string fileName, SaveFormat format)
        {
            Save(fileName, new SaveOptions { SaveFormat = format });
        }

        /// <summary>
        /// Saves the workbook to a file using explicit save options.
        /// </summary>
        public void Save(string fileName, SaveOptions options)
        {
            if (fileName == null) throw new ArgumentNullException(nameof(fileName));
            if (options == null) throw new ArgumentNullException(nameof(options));

            using (var stream = File.Create(fileName))
            {
                Save(stream, options);
            }
        }

        /// <summary>
        /// Saves the workbook to a stream using the specified save format.
        /// </summary>
        public void Save(Stream stream, SaveFormat format)
        {
            Save(stream, new SaveOptions { SaveFormat = format });
        }

        /// <summary>
        /// Saves the workbook to a stream using explicit save options.
        /// </summary>
        public void Save(Stream stream, SaveOptions options)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (options == null) throw new ArgumentNullException(nameof(options));

            try
            {
                XlsxWorkbookSerializer.Save(_model, stream, options);
            }
            catch (CellsException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new WorkbookSaveException("Failed to save XLSX workbook.", exception);
            }
        }

        /// <summary>
        /// Releases resources associated with the workbook instance.
        /// </summary>
        public void Dispose()
        {
        }

        private void LoadFromStream(Stream stream, LoadOptions options)
        {
            try
            {
                var loadedModel = XlsxWorkbookSerializer.Load(stream, options, LoadDiagnostics);
                // Keep the Workbook facade and child facade instances stable after load by
                // copying the loaded state into the existing model graph.
                _model.Worksheets.Clear();
                _model.Worksheets.AddRange(loadedModel.Worksheets);
                _model.Settings.DateSystem = loadedModel.Settings.DateSystem;
                _model.Settings.DisplayCulture = (CultureInfo)loadedModel.Settings.DisplayCulture.Clone();
                _model.Properties.CopyFrom(loadedModel.Properties);
                _model.DocumentProperties.CopyFrom(loadedModel.DocumentProperties);
                _model.DefaultStyle = loadedModel.DefaultStyle.Clone();
                _model.ActiveSheetIndex = loadedModel.ActiveSheetIndex;
                _model.DefinedNames.Clear();
                _model.DefinedNames.AddRange(loadedModel.DefinedNames);
                _model.RawThemeXml = loadedModel.RawThemeXml;
                _model.ExternalLinks.Clear();
                _model.ExternalLinks.AddRange(loadedModel.ExternalLinks);
            }
            catch (CellsException)
            {
                throw;
            }
            catch (Exception exception)
            {
                throw new WorkbookLoadException("Failed to load XLSX workbook.", exception);
            }
        }
    }
}
