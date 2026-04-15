using System.Linq;
using System.IO;
using System.Collections.Generic;
using System;
namespace Aspose.Cells_FOSS
{
    internal static class DefinedNameUtility
    {
        internal const string PrintAreaDefinedName = "_xlnm.Print_Area";
        internal const string PrintTitlesDefinedName = "_xlnm.Print_Titles";
        internal const string FilterDatabaseDefinedName = "_xlnm._FilterDatabase";

        internal static bool IsReservedName(string name)
        {
            return string.Equals(name, PrintAreaDefinedName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, PrintTitlesDefinedName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, FilterDatabaseDefinedName, StringComparison.OrdinalIgnoreCase);
        }

        internal static string NormalizeName(string name)
        {
            var normalized = (name ?? string.Empty).Trim();
            if (normalized.Length == 0)
            {
                throw new CellsException("Defined name must be non-empty.");
            }

            if (IsReservedName(normalized))
            {
                throw new CellsException("Built-in print defined names must be managed through PageSetup.");
            }

            return normalized;
        }

        internal static string NormalizeFormula(string formula)
        {
            var normalized = (formula ?? string.Empty).Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1).Trim();
            }

            if (normalized.Length == 0)
            {
                throw new CellsException("Defined name formula must be non-empty.");
            }

            return normalized;
        }

        internal static string NormalizeComment(string comment)
        {
            return (comment ?? string.Empty).Trim();
        }

        internal static bool SameScope(int? left, int? right)
        {
            if (!left.HasValue && !right.HasValue)
            {
                return true;
            }

            return left.HasValue && right.HasValue && left.Value == right.Value;
        }
    }
}
