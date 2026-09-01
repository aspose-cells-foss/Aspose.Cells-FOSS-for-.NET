using System.IO;
using System;
using System.Collections.Generic;

namespace Aspose.Cells_FOSS.Core
{
    /// <summary>
    /// Represents shared string repository.
    /// </summary>
    internal sealed class SharedStringRepository
    {
        private readonly Dictionary<string, int> _indices = new Dictionary<string, int>(StringComparer.Ordinal);
        private readonly List<SharedStringEntry> _entries = new List<SharedStringEntry>();

        /// <summary>
        /// Gets the values.
        /// </summary>
        public IReadOnlyList<string> Values
        {
            get
            {
                var values = new List<string>(_entries.Count);
                for (var index = 0; index < _entries.Count; index++)
                {
                    values.Add(_entries[index].Text ?? string.Empty);
                }

                return values;
            }
        }

        /// <summary>
        /// Gets the logical shared-string entries.
        /// </summary>
        public IReadOnlyList<SharedStringEntry> Entries
        {
            get
            {
                return _entries;
            }
        }

        /// <summary>
        /// Clears the current state.
        /// </summary>
        public void Clear()
        {
            _indices.Clear();
            _entries.Clear();
        }

        /// <summary>
        /// Attempts to get value.
        /// </summary>
        /// <param name="index">The zero-based index.</param>
        /// <param name="value">The value.</param>
        /// <returns><see langword="true"/> if the operation succeeds; otherwise, <see langword="false"/>.</returns>
        public bool TryGetValue(int index, out string value)
        {
            if (index >= 0 && index < _entries.Count)
            {
                value = _entries[index].Text ?? string.Empty;
                return true;
            }

            value = string.Empty;
            return false;
        }

        /// <summary>
        /// Attempts to get a full shared-string entry.
        /// </summary>
        public bool TryGetEntry(int index, out SharedStringEntry entry)
        {
            if (index >= 0 && index < _entries.Count)
            {
                entry = _entries[index].Clone();
                return true;
            }

            entry = null;
            return false;
        }

        /// <summary>
        /// Performs intern.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns>The int.</returns>
        public int Intern(string value)
        {
            return Intern(new SharedStringEntry
            {
                Text = value ?? string.Empty,
            });
        }

        /// <summary>
        /// Interns a rich-text-capable shared-string entry.
        /// </summary>
        public int Intern(SharedStringEntry entry)
        {
            var safeEntry = entry == null ? new SharedStringEntry() : entry.Clone();
            safeEntry.Text = safeEntry.Text ?? string.Empty;

            var key = BuildKey(safeEntry);
            int index;
            if (_indices.TryGetValue(key, out index))
            {
                return index;
            }

            index = _entries.Count;
            _entries.Add(safeEntry);
            _indices[key] = index;
            return index;
        }

        private static string BuildKey(SharedStringEntry entry)
        {
            var runs = entry.Runs;
            if (runs == null || runs.Count == 0)
            {
                return "T|" + (entry.Text ?? string.Empty);
            }

            var builder = new System.Text.StringBuilder();
            builder.Append("R|");
            builder.Append(entry.Text ?? string.Empty);
            for (var index = 0; index < runs.Count; index++)
            {
                var run = runs[index];
                builder.Append('|');
                builder.Append(run.StartIndex);
                builder.Append(':');
                builder.Append(run.Length);
                builder.Append(':');
                builder.Append(run.Font.Name ?? string.Empty);
                builder.Append(':');
                builder.Append(run.Font.Size.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(':');
                builder.Append(run.Font.Bold ? '1' : '0');
                builder.Append(':');
                builder.Append(run.Font.Italic ? '1' : '0');
                builder.Append(':');
                builder.Append((int)run.Font.Underline);
                builder.Append(':');
                builder.Append(run.Font.StrikeThrough ? '1' : '0');
                builder.Append(':');
                builder.Append(run.Font.Color.A.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(run.Font.Color.R.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(run.Font.Color.G.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(run.Font.Color.B.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
                builder.Append(':');
                builder.Append(run.Font.Color.ThemeIndex.HasValue ? run.Font.Color.ThemeIndex.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty);
                builder.Append(':');
                builder.Append(run.Font.Color.Tint.HasValue ? run.Font.Color.Tint.Value.ToString("R", System.Globalization.CultureInfo.InvariantCulture) : string.Empty);
                builder.Append(':');
                builder.Append(run.Font.Color.Indexed.HasValue ? run.Font.Color.Indexed.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty);
                builder.Append(':');
                builder.Append(run.Font.Family.HasValue ? run.Font.Family.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : string.Empty);
                builder.Append(':');
                builder.Append((int)run.Font.Scheme);
            }

            return builder.ToString();
        }
    }
}

