using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ConstructionControl
{
    internal static class ExternalToolPaths
    {
        private const string SoftMakerBundleFolder = "SoftMaker.Office.Professional.v2024.1230.1206";

        public static string ResolvePdfEditorPath(string configuredPath)
            => ResolveExecutablePath(
                configuredPath,
                GetPdfEditorCandidates(),
                new[] { "PDFXEdit64.exe", "PDFXEdit.exe" },
                GetPdfSearchRoots());

        public static string ResolveSpreadsheetEditorPath(string configuredPath)
            => ResolveExecutablePath(
                configuredPath,
                GetSpreadsheetEditorCandidates(),
                new[] { "PlanMaker.exe" },
                GetSpreadsheetSearchRoots());

        public static string NormalizeConfiguredExecutablePath(string rawPath)
        {
            if (string.IsNullOrWhiteSpace(rawPath))
                return string.Empty;

            var expanded = Environment.ExpandEnvironmentVariables(rawPath.Trim());
            try
            {
                return Path.GetFullPath(expanded);
            }
            catch
            {
                return expanded;
            }
        }

        private static string ResolveExecutablePath(
            string configuredPath,
            IEnumerable<string> directCandidates,
            IEnumerable<string> executableNames,
            IEnumerable<string> searchRoots)
        {
            var normalizedConfiguredPath = NormalizeConfiguredExecutablePath(configuredPath);
            if (File.Exists(normalizedConfiguredPath))
                return normalizedConfiguredPath;

            foreach (var candidate in directCandidates)
            {
                var normalized = NormalizeConfiguredExecutablePath(candidate);
                if (File.Exists(normalized))
                    return normalized;
            }

            var names = executableNames?
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray() ?? Array.Empty<string>();

            foreach (var root in searchRoots ?? Enumerable.Empty<string>())
            {
                var found = FindExecutable(root, names);
                if (!string.IsNullOrWhiteSpace(found))
                    return found;
            }

            return string.Empty;
        }

        private static IEnumerable<string> GetPdfEditorCandidates()
        {
            yield return ExpandPath(@"%ProgramFiles%\Tracker Software\PDF Editor\PDFXEdit.exe");
            yield return ExpandPath(@"%ProgramFiles%\Tracker Software\PDF Editor\PDFXEdit64.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\Tracker Software\PDF Editor\PDFXEdit.exe");
            yield return ExpandPath(@"%ProgramFiles%\Tracker Software\PDF-XChange Editor\PDFXEdit.exe");
            yield return ExpandPath(@"%ProgramFiles%\Tracker Software\PDF-XChange Editor\PDFXEdit64.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\Tracker Software\PDF-XChange Editor\PDFXEdit.exe");
        }

        private static IEnumerable<string> GetSpreadsheetEditorCandidates()
        {
            yield return Path.Combine(AppContext.BaseDirectory, "Dependencies", SoftMakerBundleFolder, "PlanMaker.exe");
            yield return Path.Combine(AppContext.BaseDirectory, "Dependencies", SoftMakerBundleFolder, "program", "PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker Office Professional 2024\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker Office Professional 2024\program\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker Office Professional 2024\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker Office Professional 2024\program\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker FreeOffice 2024\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker FreeOffice 2024\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker FreeOffice 2021\PlanMaker.exe");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker FreeOffice 2021\PlanMaker.exe");
        }

        private static IEnumerable<string> GetPdfSearchRoots()
        {
            yield return ExpandPath(@"%ProgramFiles%\Tracker Software");
            yield return ExpandPath(@"%ProgramFiles(x86)%\Tracker Software");
        }

        private static IEnumerable<string> GetSpreadsheetSearchRoots()
        {
            yield return Path.Combine(AppContext.BaseDirectory, "Dependencies");
            yield return Path.Combine(AppContext.BaseDirectory, SoftMakerBundleFolder);
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker Office Professional 2024");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker Office Professional 2024");
            yield return ExpandPath(@"%ProgramFiles%\SoftMaker FreeOffice 2024");
            yield return ExpandPath(@"%ProgramFiles(x86)%\SoftMaker FreeOffice 2024");
        }

        private static string FindExecutable(string rootPath, IReadOnlyCollection<string> executableNames)
        {
            if (string.IsNullOrWhiteSpace(rootPath) || executableNames == null || executableNames.Count == 0)
                return string.Empty;

            var normalizedRoot = NormalizeConfiguredExecutablePath(rootPath);
            if (!Directory.Exists(normalizedRoot))
                return string.Empty;

            try
            {
                foreach (var executableName in executableNames)
                {
                    var directPath = Path.Combine(normalizedRoot, executableName);
                    if (File.Exists(directPath))
                        return Path.GetFullPath(directPath);
                }

                foreach (var executableName in executableNames)
                {
                    var found = Directory
                        .EnumerateFiles(normalizedRoot, executableName, SearchOption.AllDirectories)
                        .FirstOrDefault(path => File.Exists(path));

                    if (!string.IsNullOrWhiteSpace(found))
                        return Path.GetFullPath(found);
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private static string ExpandPath(string rawPath)
            => string.IsNullOrWhiteSpace(rawPath)
                ? string.Empty
                : Environment.ExpandEnvironmentVariables(rawPath.Trim());
    }
}
