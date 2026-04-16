using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace ConstructionControl
{
    internal static class HiddenWorksActDocumentBuilder
    {
        private static readonly CultureInfo RuCulture = new("ru-RU");
        private const double A4Width = 793.7;
        private const double A4Height = 1122.5;

        public static FlowDocument BuildSingle(HiddenWorkActRecord act)
            => Build(act == null ? Array.Empty<HiddenWorkActRecord>() : new[] { act });

        public static FlowDocument Build(IEnumerable<HiddenWorkActRecord> acts)
        {
            var document = CreateDocument();
            var orderedActs = (acts ?? Enumerable.Empty<HiddenWorkActRecord>())
                .Where(x => x != null)
                .ToList();

            if (orderedActs.Count == 0)
            {
                document.Blocks.Add(new Paragraph(new Run("Нет актов для предпросмотра."))
                {
                    FontSize = 16,
                    FontWeight = FontWeights.SemiBold,
                    Margin = new Thickness(0, 12, 0, 0)
                });

                return document;
            }

            for (var i = 0; i < orderedActs.Count; i++)
                AppendAct(document, orderedActs[i], i > 0);

            return document;
        }

        private static FlowDocument CreateDocument()
        {
            return new FlowDocument
            {
                FontFamily = new FontFamily("Times New Roman"),
                FontSize = 10.5,
                PageWidth = A4Width,
                PageHeight = A4Height,
                ColumnWidth = A4Width,
                PagePadding = new Thickness(28, 22, 28, 22),
                TextAlignment = TextAlignment.Left
            };
        }

        private static void AppendAct(FlowDocument document, HiddenWorkActRecord act, bool breakPageBefore)
        {
            var workTitle = NormalizeText(act.WorkTitle);
            var materialsText = BuildMaterialsText(act);
            var deviationsText = act.Deviations?.Trim() ?? string.Empty;

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Center,
                14,
                FontWeights.Bold,
                null,
                new Thickness(0, 0, 0, 3),
                null,
                breakPageBefore,
                new Run("Акт освидетельствования скрытых работ № "),
                CreateUnderline("            ", true)));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Center,
                12,
                FontWeights.SemiBold,
                FontStyles.Italic,
                new Thickness(0, 0, 0, 0),
                null,
                false,
                CreateUnderline(workTitle, true, true)));
            document.Blocks.Add(CreateCaption("наименование работ"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Center,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 0),
                null,
                false,
                new Run("выполненных на объекте: "),
                CreateUnderline($"«{NormalizeText(act.FullObjectName)}»", italic: true)));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 6),
                null,
                false,
                new Run("«"),
                CreateUnderline(act.EndDate.Day.ToString("00", CultureInfo.InvariantCulture), italic: true),
                new Run("» "),
                CreateUnderline(GetGenitiveMonthName(act.EndDate), italic: true),
                new Run(" "),
                CreateUnderline(act.EndDate.Year.ToString(CultureInfo.InvariantCulture), italic: true),
                new Run(" года.")));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 0, 0, 2),
                null,
                false,
                new Run("Комиссия в составе:")));

            document.Blocks.Add(CreateCommissionParagraph(
                "Представителя генподрядной строительно-монтажной организации ",
                NormalizeText(act.GeneralContractorInfo)));
            document.Blocks.Add(CreateCaption("должность, фамилия, инициалы"));
            document.Blocks.Add(CreateCommissionParagraph(
                "представителя субподрядной строительно-монтажной организации (в случаях выполнения работ субподрядной организацией) ",
                NormalizeText(act.SubcontractorInfo)));
            document.Blocks.Add(CreateCaption("должность, фамилия, инициалы"));
            document.Blocks.Add(CreateCommissionParagraph(
                "представителя технического надзора заказчика ",
                NormalizeText(act.TechnicalSupervisorInfo)));
            document.Blocks.Add(CreateCaption("должность, фамилия, инициалы"));
            document.Blocks.Add(CreateCommissionParagraph(
                "представителя проектной организации (в случаях осуществления авторского надзора проектной организацией) ",
                NormalizeText(act.ProjectOrganizationInfo)));
            document.Blocks.Add(CreateCaption("должность, фамилия, инициалы"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 0, 0, 0),
                null,
                false,
                new Run("произвела осмотр работ, выполненных "),
                CreateUnderline(NormalizeText(act.WorkExecutorInfo), italic: true)));
            document.Blocks.Add(CreateCaption("наименование строительно-монтажной организации"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 5, 0, 3),
                null,
                false,
                new Run("и составила настоящий акт о нижеследующем:")));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 0, 0, 0),
                null,
                false,
                new Run("1. К освидетельствованию предъявлены следующие работы")));
            document.Blocks.Add(CreateParagraph(
                TextAlignment.Center,
                12,
                FontWeights.SemiBold,
                FontStyles.Italic,
                new Thickness(0, 0, 0, 0),
                null,
                false,
                CreateUnderline(workTitle, italic: true)));
            document.Blocks.Add(CreateCaption("наименование работ"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 0),
                null,
                false,
                new Run("2. Работы выполнены по проектной документации "),
                CreateUnderline(NormalizeText(act.ProjectDocumentation), italic: true)));
            document.Blocks.Add(CreateCaption("наименование проектной организации, номер чертежей и дата их составления"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 0),
                null,
                false,
                new Run("3. При выполнении работ применены: "),
                CreateUnderline(materialsText, italic: true)));
            document.Blocks.Add(CreateCaption("наименование материалов, конструкций, изделий со ссылкой на сертификаты или иные документы, подтверждающие качество"));

            document.Blocks.Add(CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 0),
                null,
                false,
                new Run("4. При выполнении работ отсутствуют/допущены (нужное подчеркнуть) нарушения требований ТНПА и (или) проектной документации ")));
            document.Blocks.Add(CreateParagraph(
                TextAlignment.Center,
                11,
                null,
                FontStyles.Italic,
                new Thickness(0, 0, 0, 0),
                null,
                false,
                CreateUnderline(NormalizeText(deviationsText), italic: true)));
            document.Blocks.Add(CreateCaption("при наличии отклонений указывается, кем согласованы, номер чертежей и дата согласования"));

            document.Blocks.Add(BuildWorkDatesTable(act));

            document.Blocks.Add(BuildSignatureTable(act));
        }

        private static Paragraph CreateCommissionParagraph(string title, string value)
        {
            return CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0, 1, 0, 0),
                null,
                false,
                new Run(title),
                CreateUnderline(value, italic: true));
        }

        private static Table BuildSignatureTable(HiddenWorkActRecord act)
        {
            var table = new Table
            {
                CellSpacing = 0,
                Margin = new Thickness(0, 4, 0, 0)
            };

            table.Columns.Add(new TableColumn { Width = new GridLength(250) });
            table.Columns.Add(new TableColumn { Width = new GridLength(110) });
            table.Columns.Add(new TableColumn { Width = new GridLength(96) });
            table.Columns.Add(new TableColumn { Width = new GridLength(170) });

            var group = new TableRowGroup();
            table.RowGroups.Add(group);

            AppendSignatureRows(group, "Представитель подрядчика", act.ContractorSignerName);
            AppendSignatureRows(group, "Представитель субподрядной организации\n(в случаях выполнения работ субподрядной организацией)", act.SubcontractorInfo);
            AppendSignatureRows(group, "Представитель технического надзора", act.TechnicalSupervisorSignerName);
            AppendSignatureRows(group, "Представитель авторского надзора", act.ProjectOrganizationSignerName);

            return table;
        }

        private static Table BuildWorkDatesTable(HiddenWorkActRecord act)
        {
            var table = new Table
            {
                CellSpacing = 0,
                Margin = new Thickness(0, 2, 0, 6)
            };

            table.Columns.Add(new TableColumn { Width = new GridLength(180) });
            table.Columns.Add(new TableColumn { Width = new GridLength(120) });

            var group = new TableRowGroup();
            table.RowGroups.Add(group);

            var startRow = new TableRow();
            startRow.Cells.Add(CreateDateCell("5. Даты: начало работ ", underline: false));
            startRow.Cells.Add(CreateDateCell(act.StartDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture), underline: true));
            group.Rows.Add(startRow);

            var endRow = new TableRow();
            endRow.Cells.Add(CreateDateCell("окончание работ ", underline: false));
            endRow.Cells.Add(CreateDateCell(act.EndDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture), underline: true));
            group.Rows.Add(endRow);

            return table;
        }

        private static TableCell CreateDateCell(string text, bool underline)
        {
            var paragraph = CreateParagraph(
                TextAlignment.Left,
                11,
                null,
                null,
                new Thickness(0),
                null,
                false,
                underline
                    ? CreateUnderline(text, italic: true)
                    : new Run(text ?? string.Empty));

            return new TableCell(paragraph)
            {
                BorderThickness = new Thickness(0),
                Padding = new Thickness(0)
            };
        }

        private static void AppendSignatureRows(TableRowGroup group, string role, string personName)
        {
            var valueRow = new TableRow();
            valueRow.Cells.Add(CreateSignatureCell(role, TextAlignment.Left, false));
            valueRow.Cells.Add(CreateSignatureCell("______________", TextAlignment.Center, true));
            valueRow.Cells.Add(CreateSignatureCell("______________", TextAlignment.Center, true));
            valueRow.Cells.Add(CreateSignatureCell(NormalizeText(personName), TextAlignment.Center, true));
            group.Rows.Add(valueRow);

            var captionRow = new TableRow();
            captionRow.Cells.Add(CreateSignatureCell(string.Empty, TextAlignment.Left, false, 8));
            captionRow.Cells.Add(CreateSignatureCell("(дата и время)", TextAlignment.Center, false, 8));
            captionRow.Cells.Add(CreateSignatureCell("(подпись)", TextAlignment.Center, false, 8));
            captionRow.Cells.Add(CreateSignatureCell("(инициалы, фамилия)", TextAlignment.Center, false, 8));
            group.Rows.Add(captionRow);
        }

        private static TableCell CreateSignatureCell(string text, TextAlignment alignment, bool underline, double fontSize = 10.5)
        {
            var paragraph = CreateParagraph(
                alignment,
                fontSize,
                null,
                null,
                new Thickness(0),
                null,
                false,
                underline
                    ? CreateUnderline(text, italic: underline)
                    : new Run(text ?? string.Empty));

            return new TableCell(paragraph)
            {
                BorderThickness = new Thickness(0),
                Padding = new Thickness(2, 1, 2, 1)
            };
        }

        private static Paragraph CreateCaption(string text)
        {
            return CreateParagraph(
                TextAlignment.Center,
                7.5,
                null,
                null,
                new Thickness(0, 0, 0, 1),
                new SolidColorBrush(Color.FromRgb(90, 90, 90)),
                false,
                new Run(text ?? string.Empty));
        }

        private static Paragraph CreateParagraph(
            TextAlignment alignment = TextAlignment.Left,
            double fontSize = 12,
            FontWeight? weight = null,
            FontStyle? style = null,
            Thickness? margin = null,
            Brush foreground = null,
            bool breakPageBefore = false,
            params Inline[] inlines)
        {
            var paragraph = new Paragraph
            {
                Margin = margin ?? new Thickness(0, 0, 0, 2),
                TextAlignment = alignment,
                FontSize = fontSize,
                LineHeight = Math.Max(fontSize + 1.4, 9.5),
                BreakPageBefore = breakPageBefore
            };

            if (weight.HasValue)
                paragraph.FontWeight = weight.Value;
            if (style.HasValue)
                paragraph.FontStyle = style.Value;
            if (foreground != null)
                paragraph.Foreground = foreground;

            foreach (var inline in inlines ?? Array.Empty<Inline>())
            {
                if (inline != null)
                    paragraph.Inlines.Add(inline);
            }

            return paragraph;
        }

        private static Inline CreateUnderline(string text, bool bold = false, bool italic = false)
        {
            var run = new Run(string.IsNullOrWhiteSpace(text) ? "____________________" : text.Trim())
            {
                TextDecorations = TextDecorations.Underline
            };

            if (bold)
                run.FontWeight = FontWeights.SemiBold;
            if (italic)
                run.FontStyle = FontStyles.Italic;

            return run;
        }

        private static string BuildMaterialsText(HiddenWorkActRecord act)
        {
            var materials = (act?.Materials ?? new())
                .Where(x => x != null && x.IsSelected && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x =>
                {
                    var parts = new List<string>
                    {
                        x.MaterialName.Trim()
                    };

                    if (!string.IsNullOrWhiteSpace(x.Passport))
                        parts.Add($"сертификат качества №{x.Passport.Trim()}");
                    if (x.ArrivalDate.HasValue)
                        parts.Add($"от {x.ArrivalDate.Value:dd.MM.yyyy}");

                    return string.Join(", ", parts);
                })
                .ToList();

            return materials.Count == 0 ? "____________________" : string.Join("; ", materials);
        }

        private static string NormalizeText(string text)
        {
            return string.IsNullOrWhiteSpace(text) ? "____________________" : text.Trim();
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
    }
}
