namespace Aspose.Cells_FOSS;

internal static class WorkbookPropertySupport
{
    internal static string NormalizeShowObjects(string? value)
    {
        return NormalizeChoice(value, "showObjects", "all", "placeholders", "none");
    }

    internal static string NormalizeUpdateLinks(string? value)
    {
        return NormalizeChoice(value, "updateLinks", "userSet", "never", "always");
    }

    internal static string NormalizeVisibility(string? value)
    {
        return NormalizeChoice(value, "visibility", "visible", "hidden", "veryHidden");
    }

    internal static string NormalizeCalculationMode(string? value)
    {
        return NormalizeChoice(value, "calcMode", "auto", "manual", "autoNoTable");
    }

    internal static string NormalizeReferenceMode(string? value)
    {
        return NormalizeChoice(value, "refMode", "A1", "R1C1");
    }

    private static string NormalizeChoice(string? value, string propertyName, params string[] allowed)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var trimmed = value!.Trim();
        for (var index = 0; index < allowed.Length; index++)
        {
            if (string.Equals(allowed[index], trimmed, System.StringComparison.OrdinalIgnoreCase))
            {
                return allowed[index];
            }
        }

        throw new CellsException("Unsupported " + propertyName + " value '" + value + "'.");
    }
}
