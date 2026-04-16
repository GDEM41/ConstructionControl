using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace ConstructionControl
{
    internal sealed class HiddenWorksActPreviewArtifact : IDisposable
    {
        public HiddenWorksActPreviewArtifact(string docxPath, string xpsPath)
        {
            DocxPath = docxPath ?? string.Empty;
            XpsPath = xpsPath ?? string.Empty;
        }

        public string DocxPath { get; }
        public string XpsPath { get; }

        public void Dispose()
        {
            TryDelete(DocxPath);
            TryDelete(XpsPath);
        }

        private static void TryDelete(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                File.Delete(path);
            }
            catch
            {
                // Ignore best-effort preview cleanup failures.
            }
        }
    }

    internal static class HiddenWorksActWordPreviewBuilder
    {
        private readonly record struct ParagraphFieldParts(string Primary, string Secondary);

        private static readonly CultureInfo RuCulture = new("ru-RU");
        private static readonly Regex MultiWhitespaceRegex = new(@"\s+", RegexOptions.Compiled);
        private static readonly string PreviewDirectoryPath = Path.Combine(
            Path.GetTempPath(),
            "ConstructionControl",
            "HiddenWorksActPreview");

        private const int ParagraphHeaderTitle = 2;
        private const int ParagraphObject = 4;
        private const int ParagraphActDate = 6;
        private const int ParagraphGeneralContractor = 8;
        private const int ParagraphSubcontractor = 10;
        private const int ParagraphTechnicalSupervisor = 12;
        private const int ParagraphProjectOrganization = 14;
        private const int ParagraphWorkExecutor = 16;
        private const int ParagraphPointOneTitle = 20;
        private const int ParagraphProjectDocumentation = 22;
        private const int ParagraphMaterials = 24;
        private const int ParagraphDeviations = 27;
        private const int ParagraphStartDate = 30;
        private const int ParagraphEndDate = 33;
        private const int ParagraphContractorSigner = 44;
        private const int ParagraphSubcontractorSigner = 52;
        private const int ParagraphTechnicalSigner = 60;
        private const int ParagraphAuthorSigner = 68;

        private const int WdFormatXps = 18;
        private const int WdDoNotSaveChanges = 0;

        private const string PrefixObject = "выполненных на объекте: ";
        private const string PrefixGeneralContractor = "Представителя генподрядной строительно-монтажной организации ";
        private const string PrefixSubcontractor = "представителя субподрядной строительно-монтажной организации (в случаях выполнения работ субподрядной организацией) ";
        private const string PrefixTechnicalSupervisor = "представителя технического надзора заказчика ";
        private const string PrefixProjectOrganization = "представителя проектной организации (в случаях осуществления авторского надзора проектной организацией) ";
        private const string PrefixWorkExecutor = "произвела осмотр работ, выполненных ";
        private const string PrefixProjectDocumentation = "2. Работы выполнены по проектной документации ";
        private const string PrefixMaterials = "3. При выполнении работ применены ";

        public static Task<HiddenWorksActPreviewArtifact> BuildPreviewAsync(HiddenWorkActRecord act)
        {
            if (act == null)
                throw new ArgumentNullException(nameof(act));

            return RunStaAsync(() => BuildPreview(act));
        }

        private static HiddenWorksActPreviewArtifact BuildPreview(HiddenWorkActRecord act)
        {
            EnsureWordAvailable();
            CleanupPreviewDirectory();
            Directory.CreateDirectory(PreviewDirectoryPath);

            var stamp = $"{DateTime.Now:yyyyMMddHHmmssfff}_{Guid.NewGuid():N}";
            var previewDocxPath = Path.Combine(PreviewDirectoryPath, $"{stamp}.docx");
            var previewXpsPath = Path.Combine(PreviewDirectoryPath, $"{stamp}.xps");

            File.Copy(ResolveTemplatePath(), previewDocxPath, true);

            dynamic word = null;
            dynamic document = null;

            try
            {
                word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")
                    ?? throw new InvalidOperationException("Microsoft Word не найден."));
                word.Visible = false;
                word.DisplayAlerts = 0;

                document = word.Documents.Open(
                    previewDocxPath,
                    ConfirmConversions: false,
                    ReadOnly: false,
                    AddToRecentFiles: false,
                    Visible: false);

                ApplyActValues(document, act);
                document.Save();
                document.SaveAs2(previewXpsPath, WdFormatXps);
            }
            catch
            {
                SafeDelete(previewDocxPath);
                SafeDelete(previewXpsPath);
                throw;
            }
            finally
            {
                try
                {
                    document?.Close(SaveChanges: WdDoNotSaveChanges);
                }
                catch
                {
                    // Ignore Word shutdown errors for temp preview files.
                }

                try
                {
                    word?.Quit(SaveChanges: WdDoNotSaveChanges);
                }
                catch
                {
                    // Ignore Word shutdown errors for temp preview files.
                }

                ReleaseComObject(document);
                ReleaseComObject(word);
            }

            return new HiddenWorksActPreviewArtifact(previewDocxPath, previewXpsPath);
        }

        private static void ApplyActValues(dynamic document, HiddenWorkActRecord act)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var workTitle = NormalizeInlineText(act.WorkTitle);
            var generalContractorParts = SplitFieldAroundSigner(act.GeneralContractorInfo, act.ContractorSignerName);
            var technicalSupervisorParts = SplitFieldAroundSigner(act.TechnicalSupervisorInfo, act.TechnicalSupervisorSignerName);
            SetWholeParagraphText(document, ParagraphHeaderTitle, workTitle, italic: true, underline: true, bold: true);
            SetParagraphFieldLine(document, ParagraphObject, PrefixObject, WrapInQuotes(act.FullObjectName), italic: true, underline: true);

            SetWholeParagraphText(document, ParagraphActDate, BuildActDateLine(act.EndDate));
            ApplyFormattingToParagraphToken(document, ParagraphActDate, act.EndDate.Day.ToString("00", CultureInfo.InvariantCulture), italic: true, underline: true);
            ApplyFormattingToParagraphToken(document, ParagraphActDate, GetGenitiveMonthName(act.EndDate), italic: true, underline: true);
            ApplyFormattingToParagraphToken(document, ParagraphActDate, act.EndDate.Year.ToString(CultureInfo.InvariantCulture), italic: true, underline: true);

            SetParagraphSplitFieldLine(document, ParagraphGeneralContractor, PrefixGeneralContractor, generalContractorParts.Primary, generalContractorParts.Secondary, italic: true, underline: true);
            SetParagraphFieldLine(document, ParagraphSubcontractor, PrefixSubcontractor, NormalizeInlineText(act.SubcontractorInfo), italic: true, underline: true, keepUnderlineWhenEmpty: true);
            SetParagraphSplitFieldLine(document, ParagraphTechnicalSupervisor, PrefixTechnicalSupervisor, technicalSupervisorParts.Primary, technicalSupervisorParts.Secondary, italic: true, underline: true);
            SetParagraphFieldLine(document, ParagraphProjectOrganization, PrefixProjectOrganization, NormalizeInlineText(act.ProjectOrganizationInfo), italic: true, underline: true);
            SetParagraphFieldLine(document, ParagraphWorkExecutor, PrefixWorkExecutor, NormalizeInlineText(act.WorkExecutorInfo), italic: true, underline: true);

            SetWholeParagraphText(document, ParagraphPointOneTitle, workTitle, italic: true, underline: true, bold: true);
            SetParagraphFieldLine(document, ParagraphProjectDocumentation, PrefixProjectDocumentation, NormalizeInlineText(act.ProjectDocumentation), italic: true, underline: true);
            SetParagraphFieldLine(document, ParagraphMaterials, PrefixMaterials, BuildMaterialsText(act), italic: true, underline: true);
            SetParagraphFieldLine(document, ParagraphDeviations, string.Empty, NormalizeInlineText(act.Deviations), italic: true, underline: true, keepUnderlineWhenEmpty: true);

            SetWholeParagraphText(document, ParagraphStartDate, FormatWorkDate(act.StartDate), italic: true);
            SetWholeParagraphText(document, ParagraphEndDate, FormatWorkDate(act.EndDate), italic: true);

            SetWholeParagraphText(document, ParagraphContractorSigner, NormalizeInlineText(act.ContractorSignerName), italic: true, preserveBlankPlaceholder: true);
            SetWholeParagraphText(document, ParagraphSubcontractorSigner, NormalizeInlineText(act.SubcontractorInfo), italic: true, preserveBlankPlaceholder: true);
            SetWholeParagraphText(document, ParagraphTechnicalSigner, NormalizeInlineText(act.TechnicalSupervisorSignerName), italic: true, preserveBlankPlaceholder: true);
            SetWholeParagraphText(document, ParagraphAuthorSigner, NormalizeInlineText(act.ProjectOrganizationSignerName), italic: true, preserveBlankPlaceholder: true);
        }

        private static void SetParagraphFieldLine(
            dynamic document,
            int paragraphIndex,
            string prefix,
            string value,
            bool italic,
            bool underline,
            bool keepUnderlineWhenEmpty = false)
        {
            SetParagraphLineCore(document, paragraphIndex, prefix, value, string.Empty, italic, underline, keepUnderlineWhenEmpty);
        }

        private static void SetParagraphSplitFieldLine(
            dynamic document,
            int paragraphIndex,
            string prefix,
            string primaryValue,
            string secondaryValue,
            bool italic,
            bool underline,
            bool keepUnderlineWhenEmpty = false)
        {
            SetParagraphLineCore(document, paragraphIndex, prefix, primaryValue, secondaryValue, italic, underline, keepUnderlineWhenEmpty);
        }

        private static void SetParagraphLineCore(
            dynamic document,
            int paragraphIndex,
            string prefix,
            string primaryValue,
            string secondaryValue,
            bool italic,
            bool underline,
            bool keepUnderlineWhenEmpty)
        {
            dynamic paragraphRange = GetEditableParagraphRange(document, paragraphIndex);
            var originalText = Convert.ToString(paragraphRange.Text) ?? string.Empty;
            var maxTrailingTabCount = CountTrailingTabs(originalText);
            var normalizedPrefix = prefix ?? string.Empty;
            var normalizedPrimaryValue = NormalizeInlineText(primaryValue);
            var normalizedSecondaryValue = NormalizeInlineText(secondaryValue);
            var hasPrimaryValue = !string.IsNullOrWhiteSpace(normalizedPrimaryValue);
            var hasSecondaryValue = !string.IsNullOrWhiteSpace(normalizedSecondaryValue);
            var shouldUnderlineTail = hasPrimaryValue || hasSecondaryValue || keepUnderlineWhenEmpty;
            var start = (int)paragraphRange.Start;

            var replacementText = normalizedPrefix;
            var formattedStart = start + replacementText.Length;

            if (hasPrimaryValue)
                replacementText += normalizedPrimaryValue;

            if (hasSecondaryValue)
            {
                if (hasPrimaryValue)
                    replacementText += '\t';

                replacementText += normalizedSecondaryValue;
            }

            var finalText = replacementText;

            if (shouldUnderlineTail && maxTrailingTabCount > 0)
            {
                paragraphRange.Text = replacementText;
                var baseLineCount = GetLineCount(paragraphRange);

                for (var tabCount = 1; tabCount <= maxTrailingTabCount; tabCount++)
                {
                    var candidateText = replacementText + new string('\t', tabCount);
                    paragraphRange.Text = candidateText;
                    if (GetLineCount(paragraphRange) > baseLineCount)
                        break;

                    finalText = candidateText;
                }
            }

            paragraphRange.Text = finalText;

            if (shouldUnderlineTail && finalText.Length > replacementText.Length)
            {
                dynamic formattedRange = document.Range(formattedStart, start + finalText.Length);
                formattedRange.Italic = italic;
                formattedRange.Underline = underline ? 1 : 0;
                ReleaseComObject(formattedRange);
            }
            else if (hasPrimaryValue || hasSecondaryValue)
            {
                dynamic formattedRange = document.Range(formattedStart, start + replacementText.Length);
                formattedRange.Italic = italic;
                formattedRange.Underline = underline ? 1 : 0;
                ReleaseComObject(formattedRange);
            }

            ReleaseComObject(paragraphRange);
        }


        private static void SetWholeParagraphText(
            dynamic document,
            int paragraphIndex,
            string text,
            bool italic = false,
            bool underline = false,
            bool bold = false,
            bool preserveBlankPlaceholder = false)
        {
            dynamic paragraphRange = GetEditableParagraphRange(document, paragraphIndex);
            var replacement = NormalizeWholeParagraphText(text, preserveBlankPlaceholder);
            var start = (int)paragraphRange.Start;
            paragraphRange.Text = replacement;

            if (!string.IsNullOrWhiteSpace(replacement))
            {
                dynamic formattedRange = document.Range(start, start + replacement.Length);
                formattedRange.Italic = italic;
                formattedRange.Underline = underline ? 1 : 0;
                formattedRange.Bold = bold;
                ReleaseComObject(formattedRange);
            }

            ReleaseComObject(paragraphRange);
        }

        private static void ApplyFormattingToParagraphToken(dynamic document, int paragraphIndex, string token, bool italic, bool underline)
        {
            if (string.IsNullOrWhiteSpace(token))
                return;

            dynamic paragraphRange = GetEditableParagraphRange(document, paragraphIndex);
            var text = Convert.ToString(paragraphRange.Text) ?? string.Empty;
            var tokenIndex = text.IndexOf(token, StringComparison.Ordinal);
            if (tokenIndex < 0)
            {
                ReleaseComObject(paragraphRange);
                return;
            }

            dynamic tokenRange = document.Range(paragraphRange.Start + tokenIndex, paragraphRange.Start + tokenIndex + token.Length);
            tokenRange.Italic = italic;
            tokenRange.Underline = underline ? 1 : 0;

            ReleaseComObject(tokenRange);
            ReleaseComObject(paragraphRange);
        }

        private static dynamic GetEditableParagraphRange(dynamic document, int paragraphIndex)
        {
            dynamic paragraph = document.Paragraphs.Item(paragraphIndex);
            dynamic paragraphRange = paragraph.Range.Duplicate;
            paragraphRange.End = paragraphRange.End - 1;
            ReleaseComObject(paragraph);
            return paragraphRange;
        }

        private static string BuildMaterialsText(HiddenWorkActRecord act)
        {
            var materials = (act?.Materials ?? new())
                .Where(x => x != null && x.IsSelected && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x =>
                {
                    var parts = new[]
                    {
                        NormalizeInlineText(x.MaterialName),
                        string.IsNullOrWhiteSpace(x.Passport)
                            ? string.Empty
                            : $"сертификат качества №{NormalizeInlineText(x.Passport)}",
                        x.ArrivalDate.HasValue
                            ? $"от {x.ArrivalDate.Value:dd.MM.yyyy}"
                            : string.Empty
                    };

                    return string.Join(", ", parts.Where(part => !string.IsNullOrWhiteSpace(part)));
                })
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .ToList();

            return materials.Count == 0 ? string.Empty : string.Join("; ", materials);
        }

        private static ParagraphFieldParts SplitFieldAroundSigner(string fullValue, string signerName)
        {
            var normalizedValue = NormalizeInlineText(fullValue);
            var normalizedSignerName = NormalizeInlineText(signerName);
            if (string.IsNullOrWhiteSpace(normalizedValue))
                return new ParagraphFieldParts(string.Empty, string.Empty);

            if (string.IsNullOrWhiteSpace(normalizedSignerName))
                return new ParagraphFieldParts(normalizedValue, string.Empty);

            var signerIndex = normalizedValue.LastIndexOf(normalizedSignerName, StringComparison.CurrentCultureIgnoreCase);
            if (signerIndex <= 0)
                return new ParagraphFieldParts(normalizedValue, string.Empty);

            var separatorIndex = normalizedValue.LastIndexOf(',', signerIndex);
            if (separatorIndex >= 0)
            {
                return new ParagraphFieldParts(
                    normalizedValue[..(separatorIndex + 1)].TrimEnd(),
                    normalizedValue[(separatorIndex + 1)..].Trim());
            }

            return new ParagraphFieldParts(
                normalizedValue[..signerIndex].TrimEnd(),
                normalizedValue[signerIndex..].Trim());
        }

        private static int CountTrailingTabs(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;

            var count = 0;
            for (var index = text.Length - 1; index >= 0 && text[index] == '\t'; index--)
                count++;

            return count;
        }

        private static int GetLineCount(dynamic range)
        {
            try
            {
                var lineCount = Convert.ToInt32(range.ComputeStatistics(1), CultureInfo.InvariantCulture);
                return Math.Max(1, lineCount);
            }
            catch
            {
                return 1;
            }
        }

        private static string BuildActDateLine(DateTime date)
            => $"«{date:dd}» {GetGenitiveMonthName(date)} {date:yyyy} года.";

        private static string FormatWorkDate(DateTime date)
            => $"{date:dd.MM.yyyy}г";

        private static string WrapInQuotes(string value)
        {
            var normalized = NormalizeInlineText(value);
            return string.IsNullOrWhiteSpace(normalized)
                ? string.Empty
                : $"«{normalized}»";
        }

        private static string NormalizeInlineText(string text)
        {
            var normalized = MultiWhitespaceRegex.Replace(text ?? string.Empty, " ").Trim();
            return normalized;
        }

        private static string NormalizeWholeParagraphText(string text, bool preserveBlankPlaceholder)
        {
            var normalized = NormalizeInlineText(text);
            if (!preserveBlankPlaceholder || !string.IsNullOrWhiteSpace(normalized))
                return normalized;

            return " ";
        }

        private static string GetGenitiveMonthName(DateTime date)
        {
            var monthIndex = date.Month - 1;
            if (monthIndex < 0 || monthIndex >= RuCulture.DateTimeFormat.MonthGenitiveNames.Length)
                return date.ToString("MMMM", RuCulture);

            var month = RuCulture.DateTimeFormat.MonthGenitiveNames[monthIndex];
            return string.IsNullOrWhiteSpace(month)
                ? date.ToString("MMMM", RuCulture)
                : month;
        }

        private static string ResolveTemplatePath()
        {
            var templateCandidates = new[]
            {
                Path.Combine(AppContext.BaseDirectory, "templates", "HiddenWorksActTemplate.reference.docx"),
                Path.Combine(AppContext.BaseDirectory, "templates", "HiddenWorksActTemplate.docx")
            };

            foreach (var candidate in templateCandidates)
            {
                if (File.Exists(candidate))
                    return candidate;
            }

            throw new FileNotFoundException("Не найден шаблон акта скрытых работ в папке templates.");
        }

        private static void CleanupPreviewDirectory()
        {
            if (!Directory.Exists(PreviewDirectoryPath))
                return;

            try
            {
                foreach (var filePath in Directory.EnumerateFiles(PreviewDirectoryPath, "*", SearchOption.TopDirectoryOnly))
                {
                    try
                    {
                        var fileInfo = new FileInfo(filePath);
                        if (fileInfo.LastWriteTimeUtc < DateTime.UtcNow.AddDays(-2))
                            fileInfo.Delete();
                    }
                    catch
                    {
                        // Ignore best-effort cleanup failures.
                    }
                }
            }
            catch
            {
                // Ignore preview cleanup failures.
            }
        }

        private static void EnsureWordAvailable()
        {
            if (Type.GetTypeFromProgID("Word.Application") == null)
                throw new InvalidOperationException("Для точного предпросмотра нужен установленный Microsoft Word.");
        }

        private static Task<TResult> RunStaAsync<TResult>(Func<TResult> action)
        {
            var completion = new TaskCompletionSource<TResult>(TaskCreationOptions.RunContinuationsAsynchronously);
            var thread = new Thread(() =>
            {
                try
                {
                    completion.SetResult(action());
                }
                catch (Exception ex)
                {
                    completion.SetException(ex);
                }
            });

            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            return completion.Task;
        }

        private static void SafeDelete(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return;

            try
            {
                File.Delete(path);
            }
            catch
            {
                // Ignore best-effort cleanup failures.
            }
        }

        private static void ReleaseComObject(object instance)
        {
            if (instance == null || !Marshal.IsComObject(instance))
                return;

            try
            {
                Marshal.FinalReleaseComObject(instance);
            }
            catch
            {
                // Ignore COM release failures during preview cleanup.
            }
        }
    }
}
