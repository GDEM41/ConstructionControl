param(
    [string]$OutputPath
)

$projectRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $projectRoot 'templates\HiddenWorksActTemplate.docx'
}

$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
$outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
}

$openXmlPathCandidates = @(
    (Join-Path $env:USERPROFILE '.nuget\packages\documentformat.openxml\3.1.1\lib\net46\DocumentFormat.OpenXml.dll'),
    (Join-Path $env:USERPROFILE '.nuget\packages\documentformat.openxml\3.1.1\lib\net8.0\DocumentFormat.OpenXml.dll')
)
$openXmlFrameworkPathCandidates = @(
    (Join-Path $env:USERPROFILE '.nuget\packages\documentformat.openxml.framework\3.1.1\lib\net46\DocumentFormat.OpenXml.Framework.dll'),
    (Join-Path $env:USERPROFILE '.nuget\packages\documentformat.openxml.framework\3.1.1\lib\net8.0\DocumentFormat.OpenXml.Framework.dll')
)

$openXmlPath = $openXmlPathCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
$openXmlFrameworkPath = $openXmlFrameworkPathCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $openXmlPath -or -not $openXmlFrameworkPath) {
    throw 'Не найдены DocumentFormat.OpenXml.dll и/или DocumentFormat.OpenXml.Framework.dll. Выполните restore NuGet-пакетов и повторите попытку.'
}

[System.Reflection.Assembly]::LoadFrom($openXmlFrameworkPath) | Out-Null
[System.Reflection.Assembly]::LoadFrom($openXmlPath) | Out-Null

if (-not ('MasterPro.Tools.HiddenWorksActTemplateBuilder' -as [type])) {
    Add-Type -IgnoreWarnings -ReferencedAssemblies @('WindowsBase', $openXmlFrameworkPath, $openXmlPath) -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace MasterPro.Tools
{
    public static class HiddenWorksActTemplateBuilder
    {
        public static void Create(string filePath)
        {
            using (var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                var mainPart = document.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = new Body();
                mainPart.Document.Append(body);

                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 280,
                    CreateRun("Акт освидетельствования скрытых работ № ", true, false, false, 28),
                    CreateRun("{{НОМЕР_АКТА}}", true, true, true, 28)));

                body.Append(CreateParagraph(JustificationValues.Center, 0, 0, 260,
                    CreateRun("{{НАИМЕНОВАНИЕ_РАБОТ}}", false, true, true, 26)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("наименование работ", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Center, 0, 140, 260,
                    CreateRun("выполненных на объекте: ", false, false, false, 24),
                    CreateRun("«{{ПОЛНОЕ_НАЗВАНИЕ_ОБЪЕКТА}}»", false, true, true, 24)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 140, 260,
                    CreateRun("«", false, false, false, 24),
                    CreateRun("{{ДЕНЬ_АКТА}}", false, true, true, 24),
                    CreateRun("» ", false, false, false, 24),
                    CreateRun("{{МЕСЯЦ_АКТА}}", false, true, true, 24),
                    CreateRun(" ", false, false, false, 24),
                    CreateRun("{{ГОД_АКТА}}", false, true, true, 24),
                    CreateRun(" г.", false, false, false, 24)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 100, 260,
                    CreateRun("Комиссия в составе:", false, false, false, 24)));

                AppendCommissionBlock(
                    body,
                    "представителя генподрядной строительно-монтажной организации ",
                    "{{ГЕНПОДРЯДЧИК_ОРГАНИЗАЦИЯ}}",
                    "{{ГЕНПОДРЯДЧИК_ДОЛЖНОСТЬ_ФИО}}");

                AppendCommissionBlock(
                    body,
                    "представителя субподрядной строительно-монтажной организации (в случаях выполнения работ субподрядной организацией) ",
                    "{{СУБПОДРЯДЧИК_ОРГАНИЗАЦИЯ}}",
                    "{{СУБПОДРЯДЧИК_ДОЛЖНОСТЬ_ФИО}}");

                AppendCommissionBlock(
                    body,
                    "представителя технического надзора заказчика: ",
                    null,
                    "{{ТЕХНАДЗОР_ДОЛЖНОСТЬ_ФИО}}");

                AppendCommissionBlock(
                    body,
                    "представителя проектной организации (в случаях осуществления авторского надзора проектной организацией): ",
                    "{{ПРОЕКТНАЯ_ОРГАНИЗАЦИЯ}}",
                    "{{АВТОРСКИЙ_НАДЗОР_ДОЛЖНОСТЬ_ФИО}}");

                body.Append(CreateParagraph(JustificationValues.Left, 0, 0, 260,
                    CreateRun("произвела осмотр работ, выполненных ", false, false, false, 24),
                    CreateRun("{{ИСПОЛНИТЕЛЬ_РАБОТ_ОРГАНИЗАЦИЯ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("наименование строительно-монтажной организации", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 120, 260,
                    CreateRun("и составила настоящий акт о нижеследующем:", false, false, false, 24)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 40, 260,
                    CreateRun("1. К освидетельствованию предъявлены следующие работы", false, false, false, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 0, 260,
                    CreateRun("{{ПЕРЕЧЕНЬ_СКРЫТЫХ_РАБОТ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("наименование работ", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 0, 260,
                    CreateRun("2. Работы выполнены по проектной документации ", false, false, false, 24),
                    CreateRun("{{ПРОЕКТНАЯ_ДОКУМЕНТАЦИЯ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("наименование проектной организации, номер чертежей и дата их составления", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 0, 260,
                    CreateRun("3. При выполнении работ применена ", false, false, false, 24),
                    CreateRun("{{МАТЕРИАЛЫ_И_СЕРТИФИКАТЫ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("наименование материалов, конструкций, изделий со ссылкой на сертификаты или иные документы, подтверждающие качество", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 40, 260,
                    CreateRun("4. При выполнении работ отсутствуют/допущены (нужное подчеркнуть) нарушения требований ТНПА и (или) проектной документации", false, false, false, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 0, 260,
                    CreateRun("{{ОТКЛОНЕНИЯ_ИЛИ_СОГЛАСОВАНИЯ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Center, 0, 120, 220,
                    CreateRun("при наличии отклонений указывается, кем согласованы, номер чертежей и дата согласования", false, false, false, 18)));

                body.Append(CreateParagraph(JustificationValues.Left, 0, 40, 260,
                    CreateRun("5. Даты: начало работ ", false, false, false, 24),
                    CreateRun("{{ДАТА_НАЧАЛА_РАБОТ}}", false, true, true, 24)));
                body.Append(CreateParagraph(JustificationValues.Left, 720, 180, 260,
                    CreateRun("окончание работ ", false, false, false, 24),
                    CreateRun("{{ДАТА_ОКОНЧАНИЯ_РАБОТ}}", false, true, true, 24)));

                body.Append(CreateSignatureTable());
                body.Append(CreateSectionProperties());

                mainPart.Document.Save();
            }
        }

        private static void AppendCommissionBlock(Body body, string introText, string organizationPlaceholder, string personPlaceholder)
        {
            var introRuns = new List<OpenXmlElement>();
            introRuns.Add(CreateRun(introText, false, false, false, 24));
            if (!string.IsNullOrWhiteSpace(organizationPlaceholder))
            {
                introRuns.Add(CreateRun(organizationPlaceholder, false, true, true, 24));
            }

            body.Append(CreateParagraph(JustificationValues.Left, 0, 0, 260, introRuns.ToArray()));
            body.Append(CreateParagraph(JustificationValues.Center, 0, 0, 260,
                CreateRun(personPlaceholder, false, true, true, 24)));
            body.Append(CreateParagraph(JustificationValues.Center, 0, 100, 220,
                CreateRun("должность, фамилия, инициалы", false, false, false, 18)));
        }

        private static Table CreateSignatureTable()
        {
            var table = new Table();
            table.Append(new TableProperties(
                new TableStyle { Val = "TableGrid" },
                new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                new TableLayout { Type = TableLayoutValues.Fixed },
                new TableBorders(
                    new TopBorder { Val = BorderValues.Nil },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Nil },
                    new InsideVerticalBorder { Val = BorderValues.Nil })));

            table.Append(new TableGrid(
                new GridColumn() { Width = "4300" },
                new GridColumn() { Width = "1900" },
                new GridColumn() { Width = "1700" },
                new GridColumn() { Width = "2400" }));

            AppendSignatureBlock(table, "Представитель подрядчика", "{{ДАТА_ВРЕМЯ_ПОДРЯДЧИК}}", "{{ПОДПИСЬ_ПОДРЯДЧИК}}", "{{ИНИЦИАЛЫ_ФАМИЛИЯ_ПОДРЯДЧИК}}");
            AppendSignatureBlock(table, "Представитель субподрядной организации (в случаях выполнения работ субподрядной организацией)", "{{ДАТА_ВРЕМЯ_СУБПОДРЯДЧИК}}", "{{ПОДПИСЬ_СУБПОДРЯДЧИК}}", "{{ИНИЦИАЛЫ_ФАМИЛИЯ_СУБПОДРЯДЧИК}}");
            AppendSignatureBlock(table, "Представитель технического надзора", "{{ДАТА_ВРЕМЯ_ТЕХНАДЗОР}}", "{{ПОДПИСЬ_ТЕХНАДЗОР}}", "{{ИНИЦИАЛЫ_ФАМИЛИЯ_ТЕХНАДЗОР}}");
            AppendSignatureBlock(table, "Представитель авторского надзора", "{{ДАТА_ВРЕМЯ_АВТОРСКИЙ_НАДЗОР}}", "{{ПОДПИСЬ_АВТОРСКИЙ_НАДЗОР}}", "{{ИНИЦИАЛЫ_ФАМИЛИЯ_АВТОРСКИЙ_НАДЗОР}}");

            return table;
        }

        private static void AppendSignatureBlock(Table table, string title, string datePlaceholder, string signPlaceholder, string personPlaceholder)
        {
            var valueRow = new TableRow(
                CreateCell(CreateParagraph(JustificationValues.Left, 0, 20, 240, CreateRun(title, false, false, false, 22)), 4300),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 20, 240, CreateRun(datePlaceholder, false, true, true, 20)), 1900),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 20, 240, CreateRun(signPlaceholder, false, true, true, 20)), 1700),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 20, 240, CreateRun(personPlaceholder, false, true, true, 20)), 2400));

            var captionRow = new TableRow(
                CreateCell(CreateParagraph(JustificationValues.Left, 0, 80, 200, CreateRun(string.Empty, false, false, false, 18)), 4300),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 80, 200, CreateRun("(дата и время)", false, false, false, 18)), 1900),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 80, 200, CreateRun("(подпись)", false, false, false, 18)), 1700),
                CreateCell(CreateParagraph(JustificationValues.Center, 0, 80, 200, CreateRun("(инициалы, фамилия)", false, false, false, 18)), 2400));

            table.Append(valueRow);
            table.Append(captionRow);
        }

        private static TableCell CreateCell(Paragraph paragraph, int width)
        {
            return new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width.ToString(CultureInfo.InvariantCulture) },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
                paragraph);
        }

        private static Paragraph CreateParagraph(JustificationValues justification, int before, int after, int line, params OpenXmlElement[] elements)
        {
            var paragraph = new Paragraph();
            paragraph.Append(new ParagraphProperties(
                new Justification { Val = justification },
                new SpacingBetweenLines
                {
                    Before = before.ToString(CultureInfo.InvariantCulture),
                    After = after.ToString(CultureInfo.InvariantCulture),
                    Line = line.ToString(CultureInfo.InvariantCulture),
                    LineRule = LineSpacingRuleValues.Auto
                }));

            foreach (var element in elements)
            {
                paragraph.Append(element);
            }

            return paragraph;
        }

        private static Run CreateRun(string text, bool bold, bool italic, bool underline, int fontSize)
        {
            var runProperties = new RunProperties(
                new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    ComplexScript = "Times New Roman"
                },
                new FontSize { Val = fontSize.ToString(CultureInfo.InvariantCulture) },
                new FontSizeComplexScript { Val = fontSize.ToString(CultureInfo.InvariantCulture) });

            if (bold)
            {
                runProperties.Append(new Bold());
            }

            if (italic)
            {
                runProperties.Append(new Italic());
            }

            if (underline)
            {
                runProperties.Append(new Underline { Val = UnderlineValues.Single });
            }

            return new Run(
                runProperties,
                new Text(text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve });
        }

        private static SectionProperties CreateSectionProperties()
        {
            return new SectionProperties(
                new PageSize { Width = 11906U, Height = 16838U },
                new PageMargin
                {
                    Top = 850,
                    Right = 850U,
                    Bottom = 850,
                    Left = 1134U,
                    Header = 709U,
                    Footer = 709U,
                    Gutter = 0U
                });
        }
    }
}
'@
}

if (Test-Path $OutputPath) {
    Remove-Item -LiteralPath $OutputPath -Force
}

[MasterPro.Tools.HiddenWorksActTemplateBuilder]::Create($OutputPath)
Write-Host "Создан шаблон Word: $OutputPath"
