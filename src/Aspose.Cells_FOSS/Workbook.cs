using System.Globalization;
using Aspose.Cells_FOSS.Core;

namespace Aspose.Cells_FOSS;

public class Workbook : IDisposable
{
    private readonly WorkbookModel _model;
    private readonly WorksheetCollection _worksheets;
    private readonly WorkbookSettings _settings;
    private readonly WorkbookProperties _properties;
    private readonly DocumentProperties _documentProperties;
    private readonly DefinedNameCollection _definedNames;

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

    public Workbook(string fileName) : this(fileName, new LoadOptions()) { }
    public Workbook(Stream stream) : this(stream, new LoadOptions()) { }

    public Workbook(string fileName, LoadOptions options) : this()
    {
        if (fileName is null) throw new ArgumentNullException(nameof(fileName));
        if (options is null) throw new ArgumentNullException(nameof(options));

        using var stream = File.OpenRead(fileName);
        LoadFromStream(stream, options);
    }

    public Workbook(Stream stream, LoadOptions options) : this()
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (options is null) throw new ArgumentNullException(nameof(options));

        LoadFromStream(stream, options);
    }

    public WorksheetCollection Worksheets
    {
        get
        {
            return _worksheets;
        }
    }

    public WorkbookSettings Settings
    {
        get
        {
            return _settings;
        }
    }

    public WorkbookProperties Properties
    {
        get
        {
            return _properties;
        }
    }

    public DocumentProperties DocumentProperties
    {
        get
        {
            return _documentProperties;
        }
    }

    public DefinedNameCollection DefinedNames
    {
        get
        {
            return _definedNames;
        }
    }

    public LoadDiagnostics LoadDiagnostics { get; }

    internal WorkbookModel Model
    {
        get
        {
            return _model;
        }
    }

    internal void EnsureUniqueSheetName(string sheetName, WorksheetModel? currentSheet = null)
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

    internal void EnsureUniqueDefinedName(DefinedNameModel? currentDefinedName, string name, int? localSheetIndex)
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

    public void Save(string fileName)
    {
        Save(fileName, new SaveOptions());
    }

    public void Save(string fileName, SaveFormat format)
    {
        Save(fileName, new SaveOptions { SaveFormat = format });
    }

    public void Save(string fileName, SaveOptions options)
    {
        if (fileName is null) throw new ArgumentNullException(nameof(fileName));
        if (options is null) throw new ArgumentNullException(nameof(options));

        using var stream = File.Create(fileName);
        Save(stream, options);
    }

    public void Save(Stream stream, SaveFormat format)
    {
        Save(stream, new SaveOptions { SaveFormat = format });
    }

    public void Save(Stream stream, SaveOptions options)
    {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (options is null) throw new ArgumentNullException(nameof(options));

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

    public void Dispose()
    {
    }

    private void LoadFromStream(Stream stream, LoadOptions options)
    {
        try
        {
            var loadedModel = XlsxWorkbookSerializer.Load(stream, options, LoadDiagnostics);
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
