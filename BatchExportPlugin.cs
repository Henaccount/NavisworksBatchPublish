using System;
using System.Collections.Generic;
using System.IO;
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;

namespace NavisworksBatchBoxPublish
{
    [Plugin(
        "NavisworksBatchBoxPublish.BatchExport",
        "OAI1",
        DisplayName = "Navisworks Batch Search Set Publish",
        ToolTip = "Headless export of all saved search sets to NWD with hidden items excluded")]
    public sealed class BatchExportPlugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            RunArguments args = null;
            Logger logger = null;
            Document doc = null;

            TryWriteStartupTrace(parameters);

            try
            {
                args = RunArguments.Parse(parameters ?? Array.Empty<string>());

                doc = Autodesk.Navisworks.Api.Application.ActiveDocument;
                if (doc == null)
                {
                    throw new InvalidOperationException(
                        "No active document. Use -OpenFile before -ExecuteAddInPlugin.");
                }

                Directory.CreateDirectory(args.OutputDirectory);

                string logPath = string.IsNullOrWhiteSpace(args.LogFile)
                    ? Path.Combine(args.OutputDirectory, "NavisworksBatchSearchSetPublish.log")
                    : args.LogFile;

                logger = new Logger(logPath);
                logger.Info("Started.");
                logger.Info("Scanning saved sets tree for search sets.");

                ResetState(doc);

                List<SearchSetSpec> searchSets = CollectSearchSets(doc.SelectionSets.Value, null, logger);
                logger.Info($"Search sets discovered: {searchSets.Count}");

                if (searchSets.Count == 0)
                {
                    logger.Warn("No search sets were found in the active document.");
                    return 0;
                }

                int exportedCount = 0;
                int emptyCount = 0;
                var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (SearchSetSpec searchSet in searchSets)
                {
                    logger.Info($"Processing search set: {searchSet.Path}");

                    ResetState(doc);

                    ModelItemCollection selected = ExecuteSearchSet(searchSet, doc);
                    if (selected == null || selected.Count == 0)
                    {
                        emptyCount++;
                        logger.Warn($"Search set returned no items: {searchSet.Path}. Skipping export.");
                        continue;
                    }

                    doc.CurrentSelection.CopyFrom(selected);
                    IsolateSelection(doc, selected);

                    string outputFile = BuildUniqueOutputPath(
                        args.OutputDirectory,
                        args.FilePrefix,
                        searchSet.FileStem,
                        usedFileNames);

                    if (File.Exists(outputFile))
                    {
                        File.Delete(outputFile);
                    }

                    var options = new NwdExportOptions
                    {
                        ExcludeHiddenItems = true,
                        FileVersion = (int)DocumentFileVersion.Navisworks2026
                    };

                    bool ok = doc.TryExportToNwd(outputFile, options);
                    if (!ok)
                    {
                        throw new InvalidOperationException(
                            "TryExportToNwd returned false for '" + outputFile + "'.");
                    }

                    exportedCount++;
                    logger.Info($"Exported: {outputFile} ({selected.Count} selected items)");
                }

                logger.Info($"Finished. Exported {exportedCount} file(s). Empty search sets skipped: {emptyCount}.");
                return 0;
            }
            catch (Exception ex)
            {
                TryLogException(args, ex, parameters);
                return 1;
            }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        ResetState(doc);
                        doc.CurrentSelection.Clear();
                    }
                }
                catch
                {
                    // Keep cleanup failures from masking the original result.
                }

                logger?.Dispose();
            }
        }

        private static List<SearchSetSpec> CollectSearchSets(
            SavedItemCollection items,
            string parentPath,
            Logger logger)
        {
            var result = new List<SearchSetSpec>();

            if (items == null)
            {
                return result;
            }

            foreach (SavedItem item in items)
            {
                if (item == null)
                {
                    continue;
                }

                string displayName = string.IsNullOrWhiteSpace(item.DisplayName)
                    ? "unnamed"
                    : item.DisplayName.Trim();

                string currentPath = string.IsNullOrWhiteSpace(parentPath)
                    ? displayName
                    : parentPath + "/" + displayName;

                try
                {
                    if (item.IsGroup && item is GroupItem group)
                    {
                        result.AddRange(CollectSearchSets(group.Children, currentPath, logger));
                        continue;
                    }

                    if (item is SelectionSet selectionSet)
                    {
                        if (selectionSet.HasSearch)
                        {
                            result.Add(new SearchSetSpec(currentPath, selectionSet));
                        }
                        else
                        {
                            logger?.Info($"Ignoring explicit selection set: {currentPath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    logger?.Warn($"Skipping saved item '{currentPath}' because it could not be inspected: {ex.Message}");
                }
            }

            return result;
        }

        private static ModelItemCollection ExecuteSearchSet(SearchSetSpec searchSet, Document doc)
        {
            if (searchSet == null)
            {
                throw new ArgumentNullException(nameof(searchSet));
            }

            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (!searchSet.Set.HasSearch)
            {
                return new ModelItemCollection();
            }

            Search search = searchSet.Set.Search;
            if (search == null)
            {
                return new ModelItemCollection();
            }

            return search.FindAll(doc, false);
        }

        private static void ResetState(Document doc)
        {
            doc.Models.SetHidden(doc.Models.RootItemDescendantsAndSelf, false);
            doc.CurrentSelection.Clear();
        }

        private static void IsolateSelection(Document doc, ModelItemCollection selected)
        {
            if (selected == null || selected.Count == 0)
            {
                return;
            }

            ModelItemCollection toShow = ExpandSelectionForVisibility(selected);

            doc.Models.SetHidden(doc.Models.RootItemDescendantsAndSelf, true);
            doc.Models.SetHidden(toShow, false);
        }

        private static ModelItemCollection ExpandSelectionForVisibility(ModelItemCollection selected)
        {
            var result = new ModelItemCollection();

            foreach (ModelItem item in selected)
            {
                if (item == null)
                {
                    continue;
                }

                result.Add(item);

                if (item.AncestorsAndSelf != null)
                {
                    result.AddRange(item.AncestorsAndSelf);
                }

                if (item.Descendants != null)
                {
                    result.AddRange(item.Descendants);
                }
            }

            return result;
        }

        private static string BuildUniqueOutputPath(
            string outputDirectory,
            string filePrefix,
            string searchSetStem,
            ISet<string> usedFileNames)
        {
            string baseName = SanitizeFileName((filePrefix ?? string.Empty) + (searchSetStem ?? string.Empty));
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = "export";
            }

            string candidate = baseName;
            int suffix = 2;
            while (usedFileNames.Contains(candidate))
            {
                candidate = baseName + "_" + suffix.ToString("000");
                suffix++;
            }

            usedFileNames.Add(candidate);
            return Path.Combine(outputDirectory, candidate + ".nwd");
        }

        private static string SanitizeFileName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            value = value.Trim().Replace('/', '_').Replace('\\', '_');

            char[] invalid = Path.GetInvalidFileNameChars();
            char[] chars = value.ToCharArray();
            for (int i = 0; i < chars.Length; i++)
            {
                for (int j = 0; j < invalid.Length; j++)
                {
                    if (chars[i] == invalid[j])
                    {
                        chars[i] = '_';
                        break;
                    }
                }
            }

            return new string(chars);
        }

        private static void TryWriteStartupTrace(string[] rawParameters)
        {
            try
            {
                string path = Path.Combine(Path.GetTempPath(), "NavisworksBatchSearchSetPublish.start.log");
                using (var writer = new StreamWriter(path, true))
                {
                    writer.WriteLine("=== " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " ===");
                    if (rawParameters == null || rawParameters.Length == 0)
                    {
                        writer.WriteLine("Raw parameters: <none>");
                    }
                    else
                    {
                        writer.WriteLine("Raw parameters: " + string.Join(" | ", rawParameters));
                    }
                }
            }
            catch
            {
                // Intentionally ignored.
            }
        }

        private static void TryLogException(RunArguments args, Exception ex, string[] rawParameters)
        {
            try
            {
                string directory = (args != null && !string.IsNullOrWhiteSpace(args.OutputDirectory))
                    ? args.OutputDirectory
                    : Path.GetTempPath();

                string logPath = (args != null && !string.IsNullOrWhiteSpace(args.LogFile))
                    ? args.LogFile
                    : Path.Combine(directory, "NavisworksBatchSearchSetPublish.error.log");

                using (var logger = new Logger(logPath))
                {
                    logger.Error(ex.ToString());
                    if (rawParameters != null && rawParameters.Length > 0)
                    {
                        logger.Error("Raw parameters: " + string.Join(" | ", rawParameters));
                    }
                }
            }
            catch
            {
                // Intentionally ignored: we do not want logging failures to hide the real problem.
            }
        }

        private sealed class SearchSetSpec
        {
            public SearchSetSpec(string path, SelectionSet set)
            {
                Path = string.IsNullOrWhiteSpace(path) ? "unnamed" : path.Trim();
                FileStem = Path.Replace("/", "__");
                Set = set ?? throw new ArgumentNullException(nameof(set));
            }

            public string Path { get; }
            public string FileStem { get; }
            public SelectionSet Set { get; }
        }

        private sealed class RunArguments
        {
            public string OutputDirectory { get; private set; }
            public string FilePrefix { get; private set; } = string.Empty;
            public string LogFile { get; private set; }

            public static RunArguments Parse(string[] parameters)
            {
                var args = new RunArguments();

                foreach (string raw in parameters)
                {
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        continue;
                    }

                    string token = Unquote(raw.Trim());

                    if (StartsWithKey(token, "outdir"))
                    {
                        args.OutputDirectory = ValueAfterEquals(token, "outdir");
                    }
                    else if (StartsWithKey(token, "prefix"))
                    {
                        args.FilePrefix = ValueAfterEquals(token, "prefix");
                    }
                    else if (StartsWithKey(token, "log"))
                    {
                        args.LogFile = ValueAfterEquals(token, "log");
                    }
                    else if (IsHelp(token))
                    {
                        throw new ArgumentException(HelpText);
                    }
                    else
                    {
                        throw new ArgumentException("Unknown argument: " + raw + Environment.NewLine + HelpText);
                    }
                }

                if (string.IsNullOrWhiteSpace(args.OutputDirectory))
                {
                    throw new ArgumentException("outdir is required." + Environment.NewLine + HelpText);
                }

                return args;
            }

            private static bool StartsWithKey(string token, string key)
            {
                return token.StartsWith(key + "=", StringComparison.OrdinalIgnoreCase)
                    || token.StartsWith("--" + key + "=", StringComparison.OrdinalIgnoreCase);
            }

            private static string ValueAfterEquals(string token, string key)
            {
                int index = token.IndexOf('=');
                if (index < 0 || index == token.Length - 1)
                {
                    return string.Empty;
                }

                return Unquote(token.Substring(index + 1));
            }

            private static bool IsHelp(string token)
            {
                return token.Equals("help", StringComparison.OrdinalIgnoreCase)
                    || token.Equals("--help", StringComparison.OrdinalIgnoreCase)
                    || token.Equals("/?", StringComparison.OrdinalIgnoreCase)
                    || token.Equals("-?", StringComparison.OrdinalIgnoreCase);
            }

            private static string Unquote(string value)
            {
                if (string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }

                value = value.Trim();
                if (value.Length >= 2 && value[0] == '"' && value[value.Length - 1] == '"')
                {
                    return value.Substring(1, value.Length - 2);
                }

                return value;
            }

            private const string HelpText =
                "Arguments:\n" +
                "  outdir=C:\\Exports\n" +
                "  prefix=Area_\n" +
                "  log=C:\\Exports\\NavisworksBatchSearchSetPublish.log\n" +
                "\n" +
                "Notes:\n" +
                "  - All saved search sets are exported recursively.\n" +
                "  - Explicit selection sets are ignored.\n" +
                "  - Folder names are included in output names using '__'.\n" +
                "  - Tokens may also be passed as --outdir=..., --prefix=..., --log=....";
        }

        private sealed class Logger : IDisposable
        {
            private readonly StreamWriter _writer;

            public Logger(string path)
            {
                string dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrWhiteSpace(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                _writer = new StreamWriter(path, true)
                {
                    AutoFlush = true
                };
            }

            public void Info(string message)
            {
                Write("INFO", message);
            }

            public void Warn(string message)
            {
                Write("WARN", message);
            }

            public void Error(string message)
            {
                Write("ERROR", message);
            }

            private void Write(string level, string message)
            {
                _writer.WriteLine($"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\t{level}\t{message}");
            }

            public void Dispose()
            {
                _writer.Dispose();
            }
        }
    }
}
