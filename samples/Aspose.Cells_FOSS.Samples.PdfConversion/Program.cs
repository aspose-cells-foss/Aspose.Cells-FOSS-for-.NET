using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.PdfConversion
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            var repositoryRoot = ResolveRepositoryRoot();
            DateTime? evaluationDate = null;
            var printOverrides = new PrintOverrides();
            var inputArguments = new List<string>();
            for (var index = 0; index < args.Length; index++)
            {
                if (printOverrides.Read(args, ref index)) continue;
                if (args[index] == "--conditional-date")
                {
                    DateTime parsedDate;
                    if (index + 1 >= args.Length || !DateTime.TryParseExact(args[index + 1],
                        "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        throw new ArgumentException("--conditional-date requires a yyyy-MM-dd date.");
                    }

                    evaluationDate = parsedDate;
                    index++;
                    continue;
                }

                inputArguments.Add(args[index]);
            }

            var inputPaths = ResolveInputPaths(repositoryRoot, inputArguments.ToArray());

            for (var index = 0; index < inputPaths.Count; index++)
            {
                var inputPath = inputPaths[index];
                var outputPath = Path.ChangeExtension(inputPath, ".pdf");

                var workbook = new Workbook(inputPath);
                printOverrides.Apply(workbook);
                workbook.Save(outputPath, new PdfSaveOptions
                {
                    ConditionalFormattingEvaluationDate = evaluationDate,
                });

                Console.WriteLine("Input: " + inputPath);
                Console.WriteLine("Output: " + outputPath);
            }
        }

        private static string ResolveRepositoryRoot()
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var directory = new DirectoryInfo(currentDirectory);

            while (directory != null)
            {
                var inputDirectory = Path.Combine(directory.FullName, "samples", "files");
                if (Directory.Exists(inputDirectory))
                {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            throw new DirectoryNotFoundException("Could not locate the repository root containing the samples/files directory.");
        }

        private static List<string> ResolveInputPaths(string repositoryRoot, string[] args)
        {
            if (args != null && args.Length > 0)
            {
                return ResolveExplicitPaths(repositoryRoot, args);
            }

            return ResolveAllInputPaths(repositoryRoot);
        }

        private static List<string> ResolveExplicitPaths(string repositoryRoot, string[] args)
        {
            var results = new List<string>();
            for (var index = 0; index < args.Length; index++)
            {
                var inputArgument = args[index];
                if (string.IsNullOrWhiteSpace(inputArgument))
                {
                    continue;
                }

                results.Add(ResolveExplicitPath(repositoryRoot, inputArgument));
            }

            if (results.Count == 0)
            {
                throw new ArgumentException("At least one non-empty .xlsx path must be provided.");
            }

            return results;
        }

        private static List<string> ResolveAllInputPaths(string repositoryRoot)
        {
            var inputDirectory = Path.Combine(repositoryRoot, "samples", "files");
            var files = Directory.GetFiles(inputDirectory, "*.xlsx", SearchOption.AllDirectories);
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);

            var results = new List<string>(files);
            if (results.Count == 0)
            {
                throw new FileNotFoundException("Could not find any .xlsx files under the samples/files directory.");
            }

            return results;
        }

        private static string ResolveExplicitPath(string repositoryRoot, string inputArgument)
        {
            var candidatePath = inputArgument;
            if (!Path.IsPathRooted(candidatePath))
            {
                candidatePath = Path.Combine(repositoryRoot, candidatePath);
                if (!File.Exists(candidatePath))
                {
                    candidatePath = Path.Combine(repositoryRoot, "samples", "files", inputArgument);
                }
            }

            if (!File.Exists(candidatePath))
            {
                throw new FileNotFoundException("The specified XLSX file does not exist.", candidatePath);
            }

            if (!string.Equals(Path.GetExtension(candidatePath), ".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException("The input file must have an .xlsx extension.", nameof(inputArgument));
            }

            return Path.GetFullPath(candidatePath);
        }
    }
}
