using System.Reflection;
using Aspose.Cells_FOSS.Core;
using static Aspose.Cells_FOSS.CompareOpenXml.ComparisonValueHelpers;
using static Aspose.Cells_FOSS.CompareOpenXml.OpenXmlComparisonSupport;

namespace Aspose.Cells_FOSS.CompareOpenXml;

internal static class Program
{
    private const string CompareRoot = @"E:\Compare\xlsx";
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".xlsx",
        ".xlsm",
        ".xltx",
        ".xltm",
    };

    private enum CompareMode
    {
        Data,
        Style,
    }

    private static int Main(string[] args)
    {
        var mode = ResolveCompareMode(args);
        var caseDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
        var logPath = Path.Combine(caseDirectory, "compareopenxml.log");

        if (!Directory.Exists(CompareRoot))
        {
            Console.Error.WriteLine($"Compare folder '{CompareRoot}' was not found.");
            return 1;
        }

        var workbookFiles = Directory.EnumerateFiles(CompareRoot, "*", SearchOption.AllDirectories)
            .Where(delegate(string path) { return SupportedExtensions.Contains(Path.GetExtension(path)); })
            .Where(delegate(string path) { return !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal); })
            .OrderBy(delegate(string path) { return path; }, StringComparer.OrdinalIgnoreCase)
            .ToList();

        Console.WriteLine($"Mode: {mode}");
        Console.WriteLine($"Comparing {workbookFiles.Count} workbook(s) from {CompareRoot}");

        using var logWriter = new StreamWriter(logPath, false);
        logWriter.WriteLine($"StartedAtUtc: {DateTime.UtcNow:O}");
        logWriter.WriteLine($"Mode: {mode}");
        logWriter.WriteLine($"CompareRoot: {CompareRoot}");
        logWriter.WriteLine($"WorkbookCount: {workbookFiles.Count}");
        logWriter.WriteLine();
        logWriter.Flush();

        var mismatchedFiles = 0;

        for (var index = 0; index < workbookFiles.Count; index++)
        {
            var workbookPath = workbookFiles[index];
            Console.WriteLine($"[{index + 1}/{workbookFiles.Count}] Comparing file: {Path.GetFileName(workbookPath)}");
            Console.WriteLine($"Path: {workbookPath}");

            logWriter.WriteLine($"[{index + 1}/{workbookFiles.Count}] Workbook: {Path.GetFileName(workbookPath)}");
            logWriter.WriteLine($"Path: {workbookPath}");

            try
            {
                var mismatches = mode == CompareMode.Data
                    ? CompareData(workbookPath)
                    : CompareStyles(workbookPath);

                if (mismatches.Count == 0)
                {
                    logWriter.WriteLine("Result: matched");
                    logWriter.WriteLine();
                    logWriter.Flush();
                    continue;
                }

                mismatchedFiles++;
                var mismatchLabel = mode == CompareMode.Data ? "cell mismatch(es)" : "style mismatch(es)";
                logWriter.WriteLine($"Result: mismatched ({mismatches.Count} {mismatchLabel})");
                foreach (var mismatch in mismatches)
                {
                    logWriter.WriteLine($"Cell: {mismatch.CellKey}");
                    logWriter.WriteLine($"  Aspose.Cells_FOSS: {mismatch.LibrarySnapshot}");
                    logWriter.WriteLine($"  Open XML SDK:     {mismatch.OpenXmlSnapshot}");
                }

                logWriter.WriteLine();
                logWriter.Flush();
            }
            catch (Exception exception)
            {
                mismatchedFiles++;
                logWriter.WriteLine("Result: error");
                logWriter.WriteLine($"Error: {exception.GetType().Name}: {exception.Message}");
                logWriter.WriteLine(exception.ToString());
                logWriter.WriteLine();
                logWriter.Flush();
            }
        }

        logWriter.WriteLine($"FinishedAtUtc: {DateTime.UtcNow:O}");
        logWriter.WriteLine($"MismatchedFiles: {mismatchedFiles}");
        logWriter.Flush();

        if (mismatchedFiles > 0)
        {
            Console.WriteLine($"Comparison completed with {mismatchedFiles} mismatched file(s). Log: {logPath}");
            return 1;
        }

        Console.WriteLine($"Comparison completed without mismatches. Log: {logPath}");
        return 0;
    }

    private static CompareMode ResolveCompareMode(string[] args)
    {
        if (args.Length > 0 && TryParseCompareMode(args[0], out var mode))
        {
            return mode;
        }

        while (true)
        {
            Console.WriteLine("Select compare mode:");
            Console.WriteLine("1. Compare data");
            Console.WriteLine("2. Compare style settings");
            Console.Write("Option: ");
            var input = Console.ReadLine();
            if (TryParseCompareMode(input, out mode))
            {
                return mode;
            }

            Console.WriteLine("Invalid option. Enter 1, 2, data, or style.");
        }
    }

    private static bool TryParseCompareMode(string? value, out CompareMode mode)
    {
        var normalized = value?.Trim().ToLowerInvariant();
        switch (normalized)
        {
            case "1":
            case "data":
            case "cell":
            case "cells":
                mode = CompareMode.Data;
                return true;

            case "2":
            case "style":
            case "styles":
                mode = CompareMode.Style;
                return true;

            default:
                mode = CompareMode.Data;
                return false;
        }
    }

    private static List<SnapshotMismatch> CompareData(string workbookPath)
    {
        var libraryCells = ReadCellDataWithLibrary(workbookPath);
        var openXmlCells = ReadCellDataWithOpenXmlSdk(workbookPath);
        return CompareCellMaps(libraryCells, openXmlCells);
    }

    private static List<SnapshotMismatch> CompareStyles(string workbookPath)
    {
        var libraryStyles = ReadCellStylesWithLibrary(workbookPath);
        var openXmlStyles = ReadCellStylesWithOpenXmlSdk(workbookPath);
        return CompareStyleMaps(libraryStyles, openXmlStyles);
    }

    private static Dictionary<string, CellSnapshot> ReadCellDataWithLibrary(string workbookPath)
    {
        var workbook = new Workbook(workbookPath);
        return EnumerateLibraryCells(workbook, delegate(string sheetName, string cellReference, Cell cell)
        {
            return new CellSnapshot(
                sheetName,
                cellReference,
                NormalizeCellType(cell.Type),
                NormalizeValue(cell.Value),
                NormalizeFormulaText(cell.Formula));
        });
    }

    private static Dictionary<string, StyleSnapshot> ReadCellStylesWithLibrary(string workbookPath)
    {
        var workbook = new Workbook(workbookPath);
        return EnumerateLibraryCells(workbook, delegate(string sheetName, string cellReference, Cell cell)
        {
            return CreateLibraryStyleSnapshot(sheetName, cellReference, cell.GetStyle());
        });
    }

    private static Dictionary<string, TSnapshot> EnumerateLibraryCells<TSnapshot>(Workbook workbook, Func<string, string, Cell, TSnapshot> snapshotFactory)
    {
        var modelProperty = typeof(Workbook).GetProperty("Model", BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("Workbook.Model reflection lookup failed.");
        var workbookModel = modelProperty.GetValue(workbook)
            ?? throw new InvalidOperationException("Workbook.Model returned null.");

        var worksheetsProperty = workbookModel.GetType().GetProperty("Worksheets")
            ?? throw new InvalidOperationException("WorkbookModel.Worksheets reflection lookup failed.");
        var worksheets = worksheetsProperty.GetValue(workbookModel) as System.Collections.IEnumerable
            ?? throw new InvalidOperationException("WorkbookModel.Worksheets is not enumerable.");

        var map = new Dictionary<string, TSnapshot>(StringComparer.OrdinalIgnoreCase);

        foreach (var worksheet in worksheets)
        {
            if (worksheet is null)
            {
                continue;
            }

            var worksheetType = worksheet.GetType();
            var sheetName = worksheetType.GetProperty("Name")?.GetValue(worksheet) as string ?? "<unknown>";
            var sheet = workbook.Worksheets[sheetName];
            var cells = worksheetType.GetProperty("Cells")?.GetValue(worksheet) as System.Collections.IEnumerable
                ?? throw new InvalidOperationException("WorksheetModel.Cells is not enumerable.");

            foreach (var entry in cells)
            {
                if (entry is null)
                {
                    continue;
                }

                var entryType = entry.GetType();
                var addressText = entryType.GetProperty("Key")?.GetValue(entry)?.ToString();
                if (!TryNormalizeCellReference(addressText, out var cellReference))
                {
                    continue;
                }

                var cell = sheet.Cells[cellReference];
                map[BuildCellKey(sheetName, cellReference)] = snapshotFactory(sheetName, cellReference, cell);
            }
        }

        return map;
    }

    private static List<SnapshotMismatch> CompareCellMaps(IReadOnlyDictionary<string, CellSnapshot> libraryCells, IReadOnlyDictionary<string, CellSnapshot> openXmlCells)
    {
        var allKeys = new SortedSet<string>(libraryCells.Keys, StringComparer.OrdinalIgnoreCase);
        allKeys.UnionWith(openXmlCells.Keys);

        var mismatches = new List<SnapshotMismatch>();
        foreach (var key in allKeys)
        {
            libraryCells.TryGetValue(key, out var librarySnapshot);
            openXmlCells.TryGetValue(key, out var openXmlSnapshot);

            if (librarySnapshot is null || openXmlSnapshot is null)
            {
                if (IsEffectivelyBlank(librarySnapshot) && IsEffectivelyBlank(openXmlSnapshot))
                {
                    continue;
                }

                mismatches.Add(new SnapshotMismatch(key, librarySnapshot?.ToString() ?? "<missing>", openXmlSnapshot?.ToString() ?? "<missing>"));
                continue;
            }

            if (!string.Equals(librarySnapshot.CellType, openXmlSnapshot.CellType, StringComparison.Ordinal)
                || !string.Equals(librarySnapshot.Value, openXmlSnapshot.Value, StringComparison.Ordinal)
                || !string.Equals(librarySnapshot.Formula, openXmlSnapshot.Formula, StringComparison.Ordinal))
            {
                mismatches.Add(new SnapshotMismatch(key, librarySnapshot.ToString(), openXmlSnapshot.ToString()));
            }
        }

        return mismatches;
    }

    private static List<SnapshotMismatch> CompareStyleMaps(IReadOnlyDictionary<string, StyleSnapshot> libraryStyles, IReadOnlyDictionary<string, StyleSnapshot> openXmlStyles)
    {
        var allKeys = new SortedSet<string>(libraryStyles.Keys, StringComparer.OrdinalIgnoreCase);
        allKeys.UnionWith(openXmlStyles.Keys);

        var mismatches = new List<SnapshotMismatch>();
        foreach (var key in allKeys)
        {
            libraryStyles.TryGetValue(key, out var librarySnapshot);
            openXmlStyles.TryGetValue(key, out var openXmlSnapshot);

            if (librarySnapshot is null || openXmlSnapshot is null)
            {
                if (IsEffectivelyDefaultStyle(librarySnapshot) && IsEffectivelyDefaultStyle(openXmlSnapshot))
                {
                    continue;
                }

                mismatches.Add(new SnapshotMismatch(key, librarySnapshot?.ToString() ?? "<missing>", openXmlSnapshot?.ToString() ?? "<missing>"));
                continue;
            }

            if (!librarySnapshot.Equals(openXmlSnapshot))
            {
                mismatches.Add(new SnapshotMismatch(key, librarySnapshot.ToString(), openXmlSnapshot.ToString()));
            }
        }

        return mismatches;
    }

    private static bool IsEffectivelyBlank(CellSnapshot? snapshot)
    {
        return snapshot is null
            || (string.Equals(snapshot.CellType, "Blank", StringComparison.Ordinal)
                && string.IsNullOrEmpty(snapshot.Formula)
                && string.Equals(snapshot.Value, "<null>", StringComparison.Ordinal));
    }

    private static bool IsEffectivelyDefaultStyle(StyleSnapshot? snapshot)
    {
        return snapshot is null || snapshot.Equals(StyleSnapshot.Default);
    }

    private static string NormalizeCellType(CellValueType type)
    {
        switch (type)
        {
            case CellValueType.String:
                return "String";
            case CellValueType.Number:
                return "Number";
            case CellValueType.Boolean:
                return "Boolean";
            case CellValueType.DateTime:
                return "DateTime";
            case CellValueType.Formula:
                return "Formula";
            default:
                return "Blank";
        }
    }
}
