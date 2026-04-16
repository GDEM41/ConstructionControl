using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Threading;
using System.Windows.Xps.Packaging;

namespace ConstructionControl
{
    public partial class MainWindow
    {
        private readonly ObservableCollection<HiddenWorkActRecord> hiddenWorkActRows = new();
        private readonly List<HiddenWorkActRecord> subscribedHiddenWorkActRecords = new();
        private readonly List<ObservableCollection<HiddenWorkActMaterialEntry>> subscribedHiddenWorkMaterialCollections = new();
        private readonly List<HiddenWorkActMaterialEntry> subscribedHiddenWorkMaterialItems = new();
        private HiddenWorkActRecord selectedHiddenWorkAct;
        private bool hiddenWorkActsInitialized;
        private bool hiddenWorkActStateDirty = true;
        private bool isRefreshingHiddenWorkActs;
        private bool isUpdatingHiddenWorkActEditor;
        private bool hiddenWorkActPreviewRefreshQueued;
        private string lastHiddenWorkActPreviewKey = string.Empty;
        private int hiddenWorkActPreviewRenderVersion;
        private XpsDocument hiddenWorkActPreviewXpsDocument;
        private HiddenWorksActPreviewArtifact hiddenWorkActPreviewArtifact;

        private sealed class HiddenWorkActSourceSegment
        {
            public string GroupKey { get; init; } = string.Empty;
            public string WorkTemplateKey { get; init; } = string.Empty;
            public string ActionName { get; init; } = string.Empty;
            public string WorkName { get; init; } = string.Empty;
            public string BlocksText { get; init; } = string.Empty;
            public string MarksText { get; init; } = string.Empty;
            public DateTime StartDate { get; init; }
            public DateTime EndDate { get; init; }
            public List<ProductionJournalEntry> Rows { get; init; } = new();
        }

        private sealed class HiddenWorkArrivalPickerItem
        {
            public JournalRecord Record { get; init; }
            public string Display { get; init; } = string.Empty;
        }

        private void HiddenWorkActsTabButton_Click(object sender, RoutedEventArgs e) => SelectMainTab(HiddenWorkActsTab);

        private void InitializeHiddenWorkActs()
        {
            EnsureHiddenWorkActStorage();
            RefreshHiddenWorkActState();
            hiddenWorkActsInitialized = true;
        }

        private void MarkHiddenWorkActStateDirty()
        {
            hiddenWorkActStateDirty = true;
        }

        private void EnsureHiddenWorkActStorage()
        {
            if (currentObject == null)
                return;

            currentObject.HiddenWorkActs ??= new List<HiddenWorkActRecord>();
            currentObject.HiddenWorkMaterialPresets ??= new List<HiddenWorkMaterialPreset>();
            currentObject.HiddenWorkTitlePrefixReplacements ??= new Dictionary<string, string>(StringComparer.CurrentCultureIgnoreCase);
            currentObject.HiddenWorkDefaults ??= new HiddenWorkActDefaults();

            foreach (var act in currentObject.HiddenWorkActs)
            {
                if (act == null)
                    continue;

                act.Materials ??= new ObservableCollection<HiddenWorkActMaterialEntry>();
                var isNewAct = act.Id == Guid.Empty;
                if (act.Id == Guid.Empty)
                    act.Id = Guid.NewGuid();

                if (string.IsNullOrWhiteSpace(act.GroupKey))
                    act.GroupKey = BuildHiddenWorkGroupKey(act.ActionName, act.WorkName, act.BlocksText, act.MarksText);
                if (string.IsNullOrWhiteSpace(act.WorkTemplateKey))
                    act.WorkTemplateKey = BuildHiddenWorkTemplateKey(act.ActionName, act.WorkName);
                if (string.IsNullOrWhiteSpace(act.WorkTitle))
                    act.WorkTitle = BuildHiddenWorkTitle(act.ActionName, act.WorkName, act.BlocksText, act.MarksText);

                if (!isNewAct)
                    continue;

                if (string.IsNullOrWhiteSpace(act.FullObjectName))
                    act.FullObjectName = GetDefaultHiddenWorkObjectName();
                if (string.IsNullOrWhiteSpace(act.GeneralContractorInfo))
                    act.GeneralContractorInfo = GetDefaultHiddenWorkGeneralContractorInfo();
                if (string.IsNullOrWhiteSpace(act.SubcontractorInfo))
                    act.SubcontractorInfo = GetDefaultHiddenWorkSubcontractorInfo();
                if (string.IsNullOrWhiteSpace(act.TechnicalSupervisorInfo))
                    act.TechnicalSupervisorInfo = GetDefaultHiddenWorkTechnicalSupervisorInfo();
                if (string.IsNullOrWhiteSpace(act.ProjectOrganizationInfo))
                    act.ProjectOrganizationInfo = GetDefaultHiddenWorkProjectOrganizationInfo();
                if (string.IsNullOrWhiteSpace(act.WorkExecutorInfo))
                    act.WorkExecutorInfo = GetDefaultHiddenWorkExecutorInfo();
                if (string.IsNullOrWhiteSpace(act.ProjectDocumentation))
                    act.ProjectDocumentation = GetDefaultHiddenWorkProjectDocumentation();
                if (string.IsNullOrWhiteSpace(act.Deviations))
                    act.Deviations = GetDefaultHiddenWorkDeviations();
                if (string.IsNullOrWhiteSpace(act.ContractorSignerName))
                    act.ContractorSignerName = GetDefaultHiddenWorkContractorSigner();
                if (string.IsNullOrWhiteSpace(act.TechnicalSupervisorSignerName))
                    act.TechnicalSupervisorSignerName = GetDefaultHiddenWorkTechnicalSigner();
                if (string.IsNullOrWhiteSpace(act.ProjectOrganizationSignerName))
                    act.ProjectOrganizationSignerName = GetDefaultHiddenWorkProjectSigner();
            }
        }

        private void RefreshHiddenWorkActState(bool saveAfterRefresh = false)
        {
            EnsureHiddenWorkActStorage();

            var selectedId = selectedHiddenWorkAct?.Id;

            isRefreshingHiddenWorkActs = true;
            try
            {
                SyncHiddenWorkActRecords();
                RewireHiddenWorkActSubscriptions();

                hiddenWorkActRows.Clear();
                foreach (var act in (currentObject?.HiddenWorkActs ?? new List<HiddenWorkActRecord>())
                    .OrderByDescending(x => x.EndDate.Date)
                    .ThenByDescending(x => x.StartDate.Date)
                    .ThenBy(x => x.WorkTitle, StringComparer.CurrentCultureIgnoreCase))
                {
                    hiddenWorkActRows.Add(act);
                }

                if (HiddenWorkActsGrid != null)
                    HiddenWorkActsGrid.ItemsSource = hiddenWorkActRows;

                selectedHiddenWorkAct = selectedId.HasValue
                    ? hiddenWorkActRows.FirstOrDefault(x => x.Id == selectedId.Value)
                    : hiddenWorkActRows.FirstOrDefault();

                if (HiddenWorkActsGrid != null)
                    HiddenWorkActsGrid.SelectedItem = selectedHiddenWorkAct;

                hiddenWorkActStateDirty = false;
                lastHiddenWorkActPreviewKey = string.Empty;
            }
            finally
            {
                isRefreshingHiddenWorkActs = false;
            }

            UpdateHiddenWorkActSummary();
            UpdateHiddenWorkActEditorState();
            UpdateHiddenWorkActPreview();

            if (saveAfterRefresh && currentObject != null)
                SaveState(SaveTrigger.System);
        }

        private void SyncHiddenWorkActRecords()
        {
            if (currentObject == null)
                return;

            var fixedActs = currentObject.HiddenWorkActs
                .Where(x => x != null && x.IsFixed)
                .OrderBy(x => x.EndDate.Date)
                .ThenBy(x => x.StartDate.Date)
                .ToList();

            var draftActsByGroup = currentObject.HiddenWorkActs
                .Where(x => x != null && !x.IsFixed)
                .GroupBy(x => x.GroupKey ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ToDictionary(
                    x => x.Key,
                    x => x.OrderBy(y => y.StartDate.Date).ThenBy(y => y.EndDate.Date).ToList(),
                    StringComparer.CurrentCultureIgnoreCase);

            var nextActs = new List<HiddenWorkActRecord>();

            foreach (var segmentGroup in BuildHiddenWorkActSourceSegments()
                .GroupBy(x => x.GroupKey, StringComparer.CurrentCultureIgnoreCase))
            {
                draftActsByGroup.TryGetValue(segmentGroup.Key, out var existingDrafts);
                existingDrafts ??= new List<HiddenWorkActRecord>();

                var orderedSegments = segmentGroup
                    .OrderBy(x => x.StartDate.Date)
                    .ThenBy(x => x.EndDate.Date)
                    .ToList();

                for (var i = 0; i < orderedSegments.Count; i++)
                {
                    var act = i < existingDrafts.Count
                        ? existingDrafts[i]
                        : new HiddenWorkActRecord();

                    ApplyDraftSegmentToAct(act, orderedSegments[i]);
                    nextActs.Add(act);
                }
            }

            nextActs.AddRange(fixedActs);
            currentObject.HiddenWorkActs = nextActs
                .OrderBy(x => x.EndDate.Date)
                .ThenBy(x => x.StartDate.Date)
                .ThenBy(x => x.WorkTitle, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private List<HiddenWorkActSourceSegment> BuildHiddenWorkActSourceSegments()
        {
            var result = new List<HiddenWorkActSourceSegment>();
            if (currentObject?.ProductionJournal == null)
                return result;

            var rowsByGroup = currentObject.ProductionJournal
                .Where(x => x != null && x.RequiresHiddenWorkAct)
                .GroupBy(BuildHiddenWorkGroupKey, StringComparer.CurrentCultureIgnoreCase);

            foreach (var group in rowsByGroup)
            {
                var orderedRows = group
                    .OrderBy(x => x.Date.Date)
                    .ThenBy(x => x.WorkName, StringComparer.CurrentCultureIgnoreCase)
                    .ToList();

                if (orderedRows.Count == 0)
                    continue;

                var remainingRows = orderedRows.ToList();
                var fixedActs = (currentObject.HiddenWorkActs ?? new List<HiddenWorkActRecord>())
                    .Where(x => x != null
                        && x.IsFixed
                        && string.Equals(x.GroupKey ?? string.Empty, group.Key, StringComparison.CurrentCultureIgnoreCase))
                    .OrderBy(x => x.StartDate.Date)
                    .ThenBy(x => x.EndDate.Date)
                    .ToList();

                foreach (var fixedAct in fixedActs)
                {
                    var beforeFixed = remainingRows
                        .Where(x => x.Date.Date < fixedAct.StartDate.Date)
                        .ToList();

                    if (beforeFixed.Count > 0)
                    {
                        result.Add(CreateHiddenWorkActSourceSegment(beforeFixed));
                        remainingRows.RemoveAll(beforeFixed.Contains);
                    }

                    remainingRows.RemoveAll(x => x.Date.Date >= fixedAct.StartDate.Date && x.Date.Date <= fixedAct.EndDate.Date);
                }

                if (remainingRows.Count > 0)
                    result.Add(CreateHiddenWorkActSourceSegment(remainingRows));
            }

            return result
                .OrderBy(x => x.EndDate.Date)
                .ThenBy(x => x.StartDate.Date)
                .ToList();
        }

        private HiddenWorkActSourceSegment CreateHiddenWorkActSourceSegment(List<ProductionJournalEntry> rows)
        {
            var firstRow = rows
                .OrderBy(x => x.Date.Date)
                .First();

            return new HiddenWorkActSourceSegment
            {
                GroupKey = BuildHiddenWorkGroupKey(firstRow),
                WorkTemplateKey = BuildHiddenWorkTemplateKey(firstRow.ActionName, firstRow.WorkName),
                ActionName = (firstRow.ActionName ?? string.Empty).Trim(),
                WorkName = (firstRow.WorkName ?? string.Empty).Trim(),
                BlocksText = (firstRow.BlocksText ?? string.Empty).Trim(),
                MarksText = (firstRow.MarksText ?? string.Empty).Trim(),
                StartDate = rows.Min(x => x.Date.Date),
                EndDate = rows.Max(x => x.Date.Date),
                Rows = rows
            };
        }

        private void ApplyDraftSegmentToAct(HiddenWorkActRecord act, HiddenWorkActSourceSegment segment)
        {
            if (act == null || segment == null)
                return;

            act.GroupKey = segment.GroupKey;
            act.WorkTemplateKey = segment.WorkTemplateKey;
            act.ActionName = segment.ActionName;
            act.WorkName = segment.WorkName;
            act.BlocksText = segment.BlocksText;
            act.MarksText = segment.MarksText;
            act.StartDate = segment.StartDate;
            act.EndDate = segment.EndDate;
            act.WorkTitle = BuildHiddenWorkTitle(segment.ActionName, segment.WorkName, segment.BlocksText, segment.MarksText);
            act.FullObjectName = string.IsNullOrWhiteSpace(act.FullObjectName)
                ? GetDefaultHiddenWorkObjectName()
                : act.FullObjectName;
            act.GeneralContractorInfo = string.IsNullOrWhiteSpace(act.GeneralContractorInfo)
                ? GetDefaultHiddenWorkGeneralContractorInfo()
                : act.GeneralContractorInfo;
            act.SubcontractorInfo = string.IsNullOrWhiteSpace(act.SubcontractorInfo)
                ? GetDefaultHiddenWorkSubcontractorInfo()
                : act.SubcontractorInfo;
            act.TechnicalSupervisorInfo = string.IsNullOrWhiteSpace(act.TechnicalSupervisorInfo)
                ? GetDefaultHiddenWorkTechnicalSupervisorInfo()
                : act.TechnicalSupervisorInfo;
            act.ProjectOrganizationInfo = string.IsNullOrWhiteSpace(act.ProjectOrganizationInfo)
                ? GetDefaultHiddenWorkProjectOrganizationInfo()
                : act.ProjectOrganizationInfo;
            act.WorkExecutorInfo = string.IsNullOrWhiteSpace(act.WorkExecutorInfo)
                ? GetDefaultHiddenWorkExecutorInfo()
                : act.WorkExecutorInfo;
            act.ProjectDocumentation = string.IsNullOrWhiteSpace(act.ProjectDocumentation)
                ? GetDefaultHiddenWorkProjectDocumentation()
                : act.ProjectDocumentation;
            act.Deviations = string.IsNullOrWhiteSpace(act.Deviations)
                ? GetDefaultHiddenWorkDeviations()
                : act.Deviations;
            act.ContractorSignerName = string.IsNullOrWhiteSpace(act.ContractorSignerName)
                ? GetDefaultHiddenWorkContractorSigner()
                : act.ContractorSignerName;
            act.TechnicalSupervisorSignerName = string.IsNullOrWhiteSpace(act.TechnicalSupervisorSignerName)
                ? GetDefaultHiddenWorkTechnicalSigner()
                : act.TechnicalSupervisorSignerName;
            act.ProjectOrganizationSignerName = string.IsNullOrWhiteSpace(act.ProjectOrganizationSignerName)
                ? GetDefaultHiddenWorkProjectSigner()
                : act.ProjectOrganizationSignerName;

            if (act.Materials == null || act.Materials.Count == 0)
                act.Materials = new ObservableCollection<HiddenWorkActMaterialEntry>(BuildSuggestedHiddenWorkMaterials(act, segment.Rows, segment.StartDate));
        }

        private void RewireHiddenWorkActSubscriptions()
        {
            foreach (var record in subscribedHiddenWorkActRecords)
                record.PropertyChanged -= HiddenWorkActRecord_PropertyChanged;
            foreach (var collection in subscribedHiddenWorkMaterialCollections)
                collection.CollectionChanged -= HiddenWorkActMaterials_CollectionChanged;
            foreach (var material in subscribedHiddenWorkMaterialItems)
                material.PropertyChanged -= HiddenWorkActMaterial_PropertyChanged;

            subscribedHiddenWorkActRecords.Clear();
            subscribedHiddenWorkMaterialCollections.Clear();
            subscribedHiddenWorkMaterialItems.Clear();

            if (currentObject?.HiddenWorkActs == null)
                return;

            foreach (var act in currentObject.HiddenWorkActs)
            {
                if (act == null)
                    continue;

                act.PropertyChanged += HiddenWorkActRecord_PropertyChanged;
                subscribedHiddenWorkActRecords.Add(act);

                act.Materials ??= new ObservableCollection<HiddenWorkActMaterialEntry>();
                act.Materials.CollectionChanged += HiddenWorkActMaterials_CollectionChanged;
                subscribedHiddenWorkMaterialCollections.Add(act.Materials);

                foreach (var material in act.Materials)
                {
                    if (material == null)
                        continue;

                    material.PropertyChanged += HiddenWorkActMaterial_PropertyChanged;
                    subscribedHiddenWorkMaterialItems.Add(material);
                }
            }
        }

        private void HiddenWorkActRecord_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (isRefreshingHiddenWorkActs)
                return;

            if (sender is HiddenWorkActRecord act && ReferenceEquals(act, selectedHiddenWorkAct))
            {
                UpdateHiddenWorkActEditorState();
                UpdateHiddenWorkActPreview();
            }

            UpdateHiddenWorkActSummary();
        }

        private void HiddenWorkActMaterials_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e?.OldItems != null)
            {
                foreach (var item in e.OldItems.OfType<HiddenWorkActMaterialEntry>())
                {
                    item.PropertyChanged -= HiddenWorkActMaterial_PropertyChanged;
                    subscribedHiddenWorkMaterialItems.Remove(item);
                }
            }

            if (e?.NewItems != null)
            {
                foreach (var item in e.NewItems.OfType<HiddenWorkActMaterialEntry>())
                {
                    item.PropertyChanged += HiddenWorkActMaterial_PropertyChanged;
                    if (!subscribedHiddenWorkMaterialItems.Contains(item))
                        subscribedHiddenWorkMaterialItems.Add(item);
                }
            }

            if (isRefreshingHiddenWorkActs)
                return;

            NotifyHiddenWorkActMaterialsChanged(sender);
            UpdateHiddenWorkActPreview();
            PersistHiddenWorkActChanges();
        }

        private void HiddenWorkActMaterial_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (isRefreshingHiddenWorkActs)
                return;

            NotifyHiddenWorkActMaterialsChanged(sender);
            UpdateHiddenWorkActPreview();
            PersistHiddenWorkActChanges();
        }

        private void UpdateHiddenWorkActSummary()
        {
            var acts = currentObject?.HiddenWorkActs ?? new List<HiddenWorkActRecord>();
            var fixedCount = acts.Count(x => x.IsFixed);
            var printedCount = acts.Count(x => x.IsPrinted);
            HiddenWorkActsSummaryText.Text = $"Всего: {acts.Count}  |  Черновики: {acts.Count - fixedCount}  |  Зафиксировано: {fixedCount}  |  Распечатано: {printedCount}";
        }

        private void NotifyHiddenWorkActMaterialsChanged(object source)
        {
            if (currentObject?.HiddenWorkActs == null || source == null)
                return;

            HiddenWorkActRecord owner = source switch
            {
                ObservableCollection<HiddenWorkActMaterialEntry> collection
                    => currentObject.HiddenWorkActs.FirstOrDefault(x => ReferenceEquals(x?.Materials, collection)),
                HiddenWorkActMaterialEntry material
                    => currentObject.HiddenWorkActs.FirstOrDefault(x => x?.Materials?.Contains(material) == true),
                _ => null
            };

            owner?.NotifyMaterialsChanged();
        }

        private void UpdateHiddenWorkActEditorState()
        {
            isUpdatingHiddenWorkActEditor = true;
            try
            {
                if (HiddenWorkActEditorPanel != null)
                    HiddenWorkActEditorPanel.DataContext = selectedHiddenWorkAct;

                if (HiddenWorkActEditorBorder != null)
                    HiddenWorkActEditorBorder.IsEnabled = selectedHiddenWorkAct != null;

                if (FixHiddenWorkActButton != null)
                    FixHiddenWorkActButton.IsEnabled = selectedHiddenWorkAct != null && !selectedHiddenWorkAct.IsFixed;

                if (UnfixHiddenWorkActButton != null)
                    UnfixHiddenWorkActButton.IsEnabled = selectedHiddenWorkAct != null && selectedHiddenWorkAct.IsFixed;

                if (PrintSelectedHiddenWorkActButton != null)
                    PrintSelectedHiddenWorkActButton.IsEnabled = selectedHiddenWorkAct != null;

                if (HiddenWorkActStatusText != null)
                    HiddenWorkActStatusText.Text = selectedHiddenWorkAct?.StateDisplay ?? string.Empty;

                var datesEditable = selectedHiddenWorkAct?.IsFixed == true;
                if (HiddenWorkActStartDatePicker != null)
                    HiddenWorkActStartDatePicker.IsEnabled = datesEditable;
                if (HiddenWorkActEndDatePicker != null)
                    HiddenWorkActEndDatePicker.IsEnabled = datesEditable;
            }
            finally
            {
                isUpdatingHiddenWorkActEditor = false;
            }
        }

        private void UpdateHiddenWorkActPreview()
        {
            if (hiddenWorkActPreviewRefreshQueued)
                return;

            hiddenWorkActPreviewRefreshQueued = true;
            Dispatcher.BeginInvoke(new Action(RenderHiddenWorkActPreview), DispatcherPriority.Background);
        }

        private async void RenderHiddenWorkActPreview()
        {
            hiddenWorkActPreviewRefreshQueued = false;

            if (HiddenWorkActPreviewViewer == null)
                return;

            var previewKey = BuildHiddenWorkActPreviewKey(selectedHiddenWorkAct);
            if (string.Equals(previewKey, lastHiddenWorkActPreviewKey, StringComparison.Ordinal))
                return;

            var renderVersion = ++hiddenWorkActPreviewRenderVersion;

            if (selectedHiddenWorkAct == null)
            {
                lastHiddenWorkActPreviewKey = previewKey;
                ClearHiddenWorkActFixedPreview();
                HiddenWorkActPreviewViewer.Document = HiddenWorksActDocumentBuilder.Build(Array.Empty<HiddenWorkActRecord>());
                return;
            }

            try
            {
                var snapshot = CloneHiddenWorkAct(selectedHiddenWorkAct);
                var previewArtifact = await HiddenWorksActWordPreviewBuilder.BuildPreviewAsync(snapshot);

                if (renderVersion != hiddenWorkActPreviewRenderVersion
                    || !string.Equals(previewKey, BuildHiddenWorkActPreviewKey(selectedHiddenWorkAct), StringComparison.Ordinal))
                {
                    previewArtifact.Dispose();
                    return;
                }

                var newXpsDocument = new XpsDocument(previewArtifact.XpsPath, FileAccess.Read);
                var newSequence = newXpsDocument.GetFixedDocumentSequence();

                ClearHiddenWorkActFixedPreview();

                hiddenWorkActPreviewXpsDocument = newXpsDocument;
                hiddenWorkActPreviewArtifact = previewArtifact;
                lastHiddenWorkActPreviewKey = previewKey;
                HiddenWorkActPreviewViewer.Document = newSequence;
            }
            catch
            {
                lastHiddenWorkActPreviewKey = string.Empty;
                ClearHiddenWorkActFixedPreview();
                HiddenWorkActPreviewViewer.Document = HiddenWorksActDocumentBuilder.BuildSingle(CloneHiddenWorkAct(selectedHiddenWorkAct));
            }
        }

        private void ClearHiddenWorkActFixedPreview()
        {
            try
            {
                hiddenWorkActPreviewXpsDocument?.Close();
            }
            catch
            {
                // Ignore preview document cleanup errors.
            }
            finally
            {
                hiddenWorkActPreviewXpsDocument = null;
            }

            hiddenWorkActPreviewArtifact?.Dispose();
            hiddenWorkActPreviewArtifact = null;
        }

        private static string BuildHiddenWorkActPreviewKey(HiddenWorkActRecord act)
        {
            if (act == null)
                return "empty";

            var materials = string.Join(";", (act.Materials ?? new ObservableCollection<HiddenWorkActMaterialEntry>())
                .Where(x => x != null)
                .Select(x => $"{x.IsSelected}|{x.MaterialName}|{x.Passport}|{x.ArrivalDate:yyyyMMdd}"));

            return string.Join("||",
                act.Id,
                act.WorkTitle,
                act.FullObjectName,
                act.GeneralContractorInfo,
                act.SubcontractorInfo,
                act.TechnicalSupervisorInfo,
                act.ProjectOrganizationInfo,
                act.WorkExecutorInfo,
                act.ProjectDocumentation,
                act.Deviations,
                act.ContractorSignerName,
                act.TechnicalSupervisorSignerName,
                act.ProjectOrganizationSignerName,
                act.StartDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture),
                act.EndDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture),
                act.IsFixed,
                act.IsPrinted,
                materials);
        }

        private void PersistHiddenWorkActChanges(bool systemSave = true)
        {
            if (isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor || currentObject == null)
                return;

            EnsureHiddenWorkActStorage();

            foreach (var act in currentObject.HiddenWorkActs ?? new List<HiddenWorkActRecord>())
                SaveHiddenWorkMaterialPreset(act);

            UpdateHiddenWorkActSummary();
            UpdateHiddenWorkActEditorState();
            UpdateHiddenWorkActPreview();

            SaveState(systemSave ? SaveTrigger.System : SaveTrigger.Manual);
        }

        private void SaveHiddenWorkMaterialPreset(HiddenWorkActRecord act)
        {
            if (currentObject == null || act == null || string.IsNullOrWhiteSpace(act.WorkTemplateKey))
                return;

            var selectedNames = (act.Materials ?? new ObservableCollection<HiddenWorkActMaterialEntry>())
                .Where(x => x != null && x.IsSelected && !string.IsNullOrWhiteSpace(x.MaterialName))
                .Select(x => x.MaterialName.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (selectedNames.Count == 0)
                return;

            currentObject.HiddenWorkMaterialPresets ??= new List<HiddenWorkMaterialPreset>();
            var existing = currentObject.HiddenWorkMaterialPresets.FirstOrDefault(x =>
                string.Equals(x.WorkTemplateKey ?? string.Empty, act.WorkTemplateKey, StringComparison.CurrentCultureIgnoreCase));

            if (existing == null)
            {
                currentObject.HiddenWorkMaterialPresets.Add(new HiddenWorkMaterialPreset
                {
                    WorkTemplateKey = act.WorkTemplateKey,
                    MaterialNames = selectedNames,
                    UpdatedAtUtc = DateTime.UtcNow
                });
                return;
            }

            existing.MaterialNames = selectedNames;
            existing.UpdatedAtUtc = DateTime.UtcNow;
        }

        private HiddenWorkActDefaults GetHiddenWorkDefaults()
            => currentObject?.HiddenWorkDefaults ?? new HiddenWorkActDefaults();

        private string GetDefaultHiddenWorkObjectName()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.FullObjectName))
                return defaults.FullObjectName.Trim();

            return string.IsNullOrWhiteSpace(currentObject?.FullObjectName) ? currentObject?.Name ?? string.Empty : currentObject.FullObjectName;
        }

        private string GetDefaultHiddenWorkGeneralContractorInfo()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.GeneralContractorInfo))
                return defaults.GeneralContractorInfo.Trim();

            var organization = currentObject?.GeneralContractorRepresentative?.Trim() ?? string.Empty;
            var foreman = currentObject?.ResponsibleForeman?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(organization))
                return foreman;
            if (string.IsNullOrWhiteSpace(foreman) || organization.Contains(foreman, StringComparison.CurrentCultureIgnoreCase))
                return organization;
            return $"{organization}, {foreman}";
        }

        private string GetDefaultHiddenWorkSubcontractorInfo()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.SubcontractorInfo))
                return defaults.SubcontractorInfo.Trim();
            return string.Empty;
        }

        private string GetDefaultHiddenWorkTechnicalSupervisorInfo()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.TechnicalSupervisorInfo))
                return defaults.TechnicalSupervisorInfo.Trim();
            return currentObject?.TechnicalSupervisorRepresentative?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkProjectOrganizationInfo()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.ProjectOrganizationInfo))
                return defaults.ProjectOrganizationInfo.Trim();
            return currentObject?.ProjectOrganizationRepresentative?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkExecutorInfo()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.WorkExecutorInfo))
                return defaults.WorkExecutorInfo.Trim();
            return currentObject?.GeneralContractorRepresentative?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkProjectDocumentation()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.ProjectDocumentation))
                return defaults.ProjectDocumentation.Trim();
            return currentObject?.ProjectDocumentationName?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkDeviations()
        {
            return string.Empty;
        }

        private string GetDefaultHiddenWorkContractorSigner()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.ContractorSignerName))
                return defaults.ContractorSignerName.Trim();

            if (!string.IsNullOrWhiteSpace(currentObject?.ResponsibleForeman))
                return currentObject.ResponsibleForeman.Trim();
            if (!string.IsNullOrWhiteSpace(currentObject?.SiteManagerName))
                return currentObject.SiteManagerName.Trim();
            return currentObject?.GeneralContractorRepresentative?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkTechnicalSigner()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.TechnicalSupervisorSignerName))
                return defaults.TechnicalSupervisorSignerName.Trim();
            return currentObject?.TechnicalSupervisorRepresentative?.Trim() ?? string.Empty;
        }

        private string GetDefaultHiddenWorkProjectSigner()
        {
            var defaults = GetHiddenWorkDefaults();
            if (!string.IsNullOrWhiteSpace(defaults.ProjectOrganizationSignerName))
                return defaults.ProjectOrganizationSignerName.Trim();
            return currentObject?.ProjectOrganizationRepresentative?.Trim() ?? string.Empty;
        }

        private string BuildHiddenWorkGroupKey(ProductionJournalEntry row)
            => row == null
                ? string.Empty
                : BuildHiddenWorkGroupKey(row.ActionName, row.WorkName, row.BlocksText, row.MarksText);

        private string BuildHiddenWorkGroupKey(string actionName, string workName, string blocksText, string marksText)
        {
            return string.Join("||",
                NormalizeHiddenWorkKeyPart(actionName),
                NormalizeHiddenWorkKeyPart(BuildHiddenWorkBaseText(actionName, workName)),
                NormalizeHiddenWorkKeyPart(blocksText),
                NormalizeHiddenWorkKeyPart(marksText));
        }

        private string BuildHiddenWorkTemplateKey(string actionName, string workName)
        {
            return string.Join("||",
                NormalizeHiddenWorkKeyPart(actionName),
                NormalizeHiddenWorkKeyPart(BuildHiddenWorkBaseText(actionName, workName)));
        }

        private static string NormalizeHiddenWorkKeyPart(string value)
        {
            return Regex.Replace((value ?? string.Empty).Trim(), @"\s+", " ").ToUpperInvariant();
        }

        private string BuildHiddenWorkBaseText(string actionName, string workName)
        {
            var action = (actionName ?? string.Empty).Trim();
            var work = ApplyHiddenWorkTitlePrefixReplacement((workName ?? string.Empty).Trim());
            var baseText = string.IsNullOrWhiteSpace(action)
                ? work
                : work.StartsWith(action, StringComparison.CurrentCultureIgnoreCase)
                    ? work
                    : $"{action} {work}";

            baseText = ApplyHiddenWorkTitlePrefixReplacement(baseText);

            return NormalizeHiddenWorkSentenceCase(baseText);
        }

        private string NormalizeHiddenWorkSentenceCase(string value)
        {
            var normalized = Regex.Replace((value ?? string.Empty).Trim(), @"\s+", " ");
            if (string.IsNullOrWhiteSpace(normalized))
                return string.Empty;

            var culture = new CultureInfo("ru-RU");
            var lower = normalized.ToLower(culture);
            return char.ToUpper(lower[0], culture) + lower[1..];
        }

        private string ApplyHiddenWorkTitlePrefixReplacement(string sourceText)
        {
            var normalizedSource = Regex.Replace((sourceText ?? string.Empty).Trim(), @"\s+", " ");
            if (string.IsNullOrWhiteSpace(normalizedSource)
                || currentObject?.HiddenWorkTitlePrefixReplacements == null
                || currentObject.HiddenWorkTitlePrefixReplacements.Count == 0)
            {
                return normalizedSource;
            }

            foreach (var pair in currentObject.HiddenWorkTitlePrefixReplacements
                .Where(x => !string.IsNullOrWhiteSpace(x.Key) && !string.IsNullOrWhiteSpace(x.Value))
                .OrderByDescending(x => x.Key.Trim().Length))
            {
                if (!TryApplyHiddenWorkTitlePrefixReplacement(normalizedSource, pair.Key.Trim(), pair.Value.Trim(), out var replaced))
                    continue;

                return replaced;
            }

            return normalizedSource;
        }

        private bool TryApplyHiddenWorkTitlePrefixReplacement(string sourceText, string sourcePrefix, string replacement, out string replacedText)
        {
            replacedText = sourceText;
            if (string.IsNullOrWhiteSpace(sourceText) || string.IsNullOrWhiteSpace(sourcePrefix) || string.IsNullOrWhiteSpace(replacement))
                return false;

            if (sourceText.StartsWith(sourcePrefix, StringComparison.CurrentCultureIgnoreCase))
            {
                var suffix = sourceText[sourcePrefix.Length..].Trim();
                replacedText = string.IsNullOrWhiteSpace(suffix)
                    ? replacement
                    : $"{replacement} {suffix}";
                return true;
            }

            var sourceMatches = Regex.Matches(sourceText, @"[^\s,;:/\\\-]+").Cast<Match>().ToList();
            var prefixMatches = Regex.Matches(sourcePrefix, @"[^\s,;:/\\\-]+").Cast<Match>().ToList();
            if (prefixMatches.Count == 0 || sourceMatches.Count < prefixMatches.Count)
                return false;

            for (var i = 0; i < prefixMatches.Count; i++)
            {
                var left = NormalizeProductionLookupTokenForMatching(sourceMatches[i].Value);
                var right = NormalizeProductionLookupTokenForMatching(prefixMatches[i].Value);
                if (!string.Equals(left, right, StringComparison.Ordinal))
                    return false;
            }

            var endIndex = sourceMatches[prefixMatches.Count - 1].Index + sourceMatches[prefixMatches.Count - 1].Length;
            var suffixBySemantic = sourceText[endIndex..].TrimStart();
            replacedText = string.IsNullOrWhiteSpace(suffixBySemantic)
                ? replacement
                : $"{replacement} {suffixBySemantic}";
            return true;
        }

        private string BuildHiddenWorkTitle(string actionName, string workName, string blocksText, string marksText)
        {
            var title = BuildHiddenWorkBaseText(actionName, workName);
            var axes = BuildHiddenWorkAxesText(blocksText);
            var marks = BuildHiddenWorkMarksText(marksText);

            if (!string.IsNullOrWhiteSpace(axes))
                title += $" в осях {axes}";
            if (!string.IsNullOrWhiteSpace(marks))
                title += $" на отм. {marks}";

            return LevelMarkHelper.PreventSingleLetterWrap(title);
        }

        private string BuildHiddenWorkAxesText(string blocksText)
        {
            var blocks = LevelMarkHelper.ParseBlocks(blocksText);
            if (blocks.Count == 0)
                return (blocksText ?? string.Empty).Trim();

            var axes = new List<string>();
            foreach (var block in blocks)
            {
                if (currentObject?.BlockAxesByNumber != null
                    && currentObject.BlockAxesByNumber.TryGetValue(block, out var axis)
                    && !string.IsNullOrWhiteSpace(axis))
                {
                    axes.Add(axis.Trim());
                }
                else
                {
                    axes.Add(block.ToString(CultureInfo.InvariantCulture));
                }
            }

            return string.Join(", ", axes
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase));
        }

        private string BuildHiddenWorkMarksText(string marksText)
        {
            var marks = LevelMarkHelper.ParseMarks(marksText);
            if (marks.Count == 0)
                return (marksText ?? string.Empty).Trim();

            return string.Join(", ", marks);
        }

        private List<HiddenWorkActMaterialEntry> BuildSuggestedHiddenWorkMaterials(
            HiddenWorkActRecord act,
            IEnumerable<ProductionJournalEntry> rows,
            DateTime startDate)
        {
            var candidateNames = (rows ?? Enumerable.Empty<ProductionJournalEntry>())
                .SelectMany(x => ParseProductionItems(x?.ElementsText)
                    .Where(item => x != null && ShouldTrackProductionItemInDemand(x, item.MaterialName)))
                .Select(x => x.MaterialName?.Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var mappedNames = GetMappedProductionMaterialsForWork(act?.WorkName)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var arrivalTypeNames = journal
                .Where(x => string.Equals((x.Category ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && !string.IsNullOrWhiteSpace(x.MaterialName)
                    && AreProductionLookupValuesEquivalent((x.MaterialGroup ?? string.Empty).Trim(), act?.WorkName ?? string.Empty))
                .Select(x => x.MaterialName.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var presetNames = currentObject?.HiddenWorkMaterialPresets?
                .FirstOrDefault(x => string.Equals(
                    x.WorkTemplateKey ?? string.Empty,
                    act?.WorkTemplateKey ?? string.Empty,
                    StringComparison.CurrentCultureIgnoreCase))
                ?.MaterialNames?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList()
                ?? new List<string>();

            var allNames = presetNames
                .Concat(candidateNames)
                .Concat(mappedNames)
                .Concat(arrivalTypeNames)
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .OrderBy(x => x, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var usePresetSelection = presetNames.Count > 0;
            var defaultSelectedNames = candidateNames
                .Concat(mappedNames)
                .Concat(arrivalTypeNames)
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            return allNames
                .Select(name =>
                {
                    var arrival = FindBestHiddenWorkMaterialArrival(name, startDate, act?.WorkName);
                    return new HiddenWorkActMaterialEntry
                    {
                        IsSelected = usePresetSelection
                            ? presetNames.Contains(name, StringComparer.CurrentCultureIgnoreCase)
                            : defaultSelectedNames.Contains(name, StringComparer.CurrentCultureIgnoreCase),
                        MaterialName = name,
                        Passport = arrival?.Passport?.Trim() ?? string.Empty,
                        ArrivalDate = arrival?.Date.Date
                    };
                })
                .ToList();
        }

        private JournalRecord FindBestHiddenWorkMaterialArrival(string materialName, DateTime startDate, string workName)
        {
            if (string.IsNullOrWhiteSpace(materialName))
                return null;

            var candidates = journal
                .Where(x => string.Equals((x.Category ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && !string.IsNullOrWhiteSpace(x.MaterialName)
                    && AreProductionLookupValuesEquivalent(x.MaterialName.Trim(), materialName.Trim()))
                .ToList();

            if (candidates.Count == 0)
                return null;

            var preferredGroup = (workName ?? string.Empty).Trim();
            var sameGroup = string.IsNullOrWhiteSpace(preferredGroup)
                ? candidates
                : candidates.Where(x => AreProductionLookupValuesEquivalent(
                    (x.MaterialGroup ?? string.Empty).Trim(),
                    preferredGroup)).ToList();

            var pool = sameGroup.Count > 0 ? sameGroup : candidates;
            var onOrBefore = pool
                .Where(x => x.Date.Date <= startDate.Date)
                .OrderByDescending(x => x.Date.Date)
                .ThenByDescending(x => string.IsNullOrWhiteSpace(x.Passport) ? 0 : 1)
                .ToList();

            if (onOrBefore.Count > 0)
                return onOrBefore[0];

            return pool
                .OrderByDescending(x => x.Date.Date)
                .ThenByDescending(x => string.IsNullOrWhiteSpace(x.Passport) ? 0 : 1)
                .FirstOrDefault();
        }

        private List<ProductionJournalEntry> GetProductionRowsForHiddenWorkAct(HiddenWorkActRecord act)
        {
            if (currentObject?.ProductionJournal == null || act == null)
                return new List<ProductionJournalEntry>();

            return currentObject.ProductionJournal
                .Where(x => x != null
                    && x.RequiresHiddenWorkAct
                    && string.Equals(BuildHiddenWorkGroupKey(x), act.GroupKey, StringComparison.CurrentCultureIgnoreCase)
                    && x.Date.Date >= act.StartDate.Date
                    && x.Date.Date <= act.EndDate.Date)
                .OrderBy(x => x.Date.Date)
                .ToList();
        }

        private HiddenWorkActRecord FindHiddenWorkActForProductionRow(ProductionJournalEntry row)
        {
            if (row == null || currentObject?.HiddenWorkActs == null)
                return null;

            var key = BuildHiddenWorkGroupKey(row);
            return currentObject.HiddenWorkActs
                .Where(x => string.Equals(x.GroupKey ?? string.Empty, key, StringComparison.CurrentCultureIgnoreCase)
                    && row.Date.Date >= x.StartDate.Date
                    && row.Date.Date <= x.EndDate.Date)
                .OrderByDescending(x => x.IsFixed)
                .ThenBy(x => x.EndDate.Date)
                .FirstOrDefault()
                ?? currentObject.HiddenWorkActs.FirstOrDefault(x =>
                    string.Equals(x.GroupKey ?? string.Empty, key, StringComparison.CurrentCultureIgnoreCase));
        }

        private void OpenHiddenWorkActForProductionRow(ProductionJournalEntry row)
        {
            if (row == null)
                return;

            if (!row.RequiresHiddenWorkAct)
            {
                MessageBox.Show("Для этой записи акт скрытых работ не требуется.");
                return;
            }

            EnsureHiddenWorkActStorage();
            if (!hiddenWorkActsInitialized || hiddenWorkActStateDirty)
                RefreshHiddenWorkActState();

            var act = FindHiddenWorkActForProductionRow(row);
            if (act == null)
            {
                MessageBox.Show("Для выбранной записи акт пока не сформирован.");
                return;
            }

            SelectMainTab(HiddenWorkActsTab);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                selectedHiddenWorkAct = act;
                if (HiddenWorkActsGrid != null)
                    SelectGridItem(HiddenWorkActsGrid, act);
                UpdateHiddenWorkActEditorState();
                UpdateHiddenWorkActPreview();
            }), DispatcherPriority.Background);
        }

        private void RefreshHiddenWorkActs_Click(object sender, RoutedEventArgs e)
        {
            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState();
        }

        private void FixSelectedHiddenWorkAct_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
            {
                MessageBox.Show("Выберите акт.");
                return;
            }

            selectedHiddenWorkAct.IsFixed = true;
            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState(saveAfterRefresh: true);
        }

        private void UnfixSelectedHiddenWorkAct_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
            {
                MessageBox.Show("Выберите акт.");
                return;
            }

            if (!selectedHiddenWorkAct.IsFixed)
            {
                MessageBox.Show("Выбранный акт уже не зафиксирован.");
                return;
            }

            if (selectedHiddenWorkAct.IsPrinted)
            {
                var resetPrinted = MessageBox.Show(
                    "Акт помечен как распечатанный. Снять фиксацию можно только вместе с отметкой печати. Продолжить?",
                    "Снятие фиксации",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Question);
                if (resetPrinted != MessageBoxResult.Yes)
                    return;

                selectedHiddenWorkAct.IsPrinted = false;
            }

            selectedHiddenWorkAct.IsFixed = false;
            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState(saveAfterRefresh: true);
        }

        private void PrintSelectedHiddenWorkAct_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
            {
                MessageBox.Show("Выберите акт.");
                return;
            }

            if (!PrintHiddenWorkActs(new[] { selectedHiddenWorkAct }, "Акт скрытых работ"))
                return;

            selectedHiddenWorkAct.IsPrinted = true;
            selectedHiddenWorkAct.IsFixed = true;
            MarkHiddenWorkActStateDirty();
            PersistHiddenWorkActChanges();
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    RefreshHiddenWorkActState(saveAfterRefresh: true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось обновить вкладку актов после печати.{Environment.NewLine}{ex.Message}", "Акты скрытых работ", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }), DispatcherPriority.Background);
        }

        private void PrintPendingHiddenWorkActs_Click(object sender, RoutedEventArgs e)
        {
            EnsureHiddenWorkActStorage();
            if (!hiddenWorkActsInitialized || hiddenWorkActStateDirty)
                RefreshHiddenWorkActState();

            var acts = (currentObject?.HiddenWorkActs ?? new List<HiddenWorkActRecord>())
                .Where(x => x != null && !x.IsPrinted)
                .OrderBy(x => x.EndDate.Date)
                .ThenBy(x => x.StartDate.Date)
                .ToList();

            if (acts.Count == 0)
            {
                MessageBox.Show("Все акты уже отмечены как распечатанные.");
                return;
            }

            if (!PrintHiddenWorkActs(acts, "Акты скрытых работ"))
                return;

            foreach (var act in acts)
            {
                act.IsPrinted = true;
                act.IsFixed = true;
            }

            MarkHiddenWorkActStateDirty();
            PersistHiddenWorkActChanges();
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    RefreshHiddenWorkActState(saveAfterRefresh: true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Не удалось обновить вкладку актов после печати.{Environment.NewLine}{ex.Message}", "Акты скрытых работ", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }), DispatcherPriority.Background);
        }

        private void DeleteSelectedHiddenWorkAct_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null || currentObject?.HiddenWorkActs == null)
            {
                MessageBox.Show("Выберите акт.");
                return;
            }

            var firstConfirm = MessageBox.Show(
                "Удалить выбранный акт? Это действие нельзя отменить автоматически.",
                "Удаление акта",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (firstConfirm != MessageBoxResult.Yes)
                return;

            var secondConfirm = MessageBox.Show(
                "Подтвердите удаление еще раз. Точно удалить акт?",
                "Удаление акта",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (secondConfirm != MessageBoxResult.Yes)
                return;

            var removed = currentObject.HiddenWorkActs.Remove(selectedHiddenWorkAct);
            if (!removed)
                return;

            selectedHiddenWorkAct = null;
            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState(saveAfterRefresh: true);
        }

        private bool PrintHiddenWorkActs(IEnumerable<HiddenWorkActRecord> acts, string title)
        {
            var items = (acts ?? Enumerable.Empty<HiddenWorkActRecord>()).Where(x => x != null).ToList();
            if (items.Count == 0)
                return false;

            try
            {
                var snapshot = items.Select(CloneHiddenWorkAct).ToList();
                var document = HiddenWorksActDocumentBuilder.Build(snapshot);
                var dialog = new PrintDialog();
                if (dialog.ShowDialog() != true)
                    return false;

                dialog.PrintDocument(((IDocumentPaginatorSource)document).DocumentPaginator, title);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось распечатать акт.{Environment.NewLine}{ex.Message}", "Печать акта", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }
        }

        private static HiddenWorkActRecord CloneHiddenWorkAct(HiddenWorkActRecord act)
        {
            return new HiddenWorkActRecord
            {
                Id = act?.Id ?? Guid.NewGuid(),
                GroupKey = act?.GroupKey ?? string.Empty,
                WorkTemplateKey = act?.WorkTemplateKey ?? string.Empty,
                ActionName = act?.ActionName ?? string.Empty,
                WorkName = act?.WorkName ?? string.Empty,
                BlocksText = act?.BlocksText ?? string.Empty,
                MarksText = act?.MarksText ?? string.Empty,
                WorkTitle = act?.WorkTitle ?? string.Empty,
                FullObjectName = act?.FullObjectName ?? string.Empty,
                GeneralContractorInfo = act?.GeneralContractorInfo ?? string.Empty,
                SubcontractorInfo = act?.SubcontractorInfo ?? string.Empty,
                TechnicalSupervisorInfo = act?.TechnicalSupervisorInfo ?? string.Empty,
                ProjectOrganizationInfo = act?.ProjectOrganizationInfo ?? string.Empty,
                WorkExecutorInfo = act?.WorkExecutorInfo ?? string.Empty,
                ProjectDocumentation = act?.ProjectDocumentation ?? string.Empty,
                Deviations = act?.Deviations ?? string.Empty,
                ContractorSignerName = act?.ContractorSignerName ?? string.Empty,
                TechnicalSupervisorSignerName = act?.TechnicalSupervisorSignerName ?? string.Empty,
                ProjectOrganizationSignerName = act?.ProjectOrganizationSignerName ?? string.Empty,
                StartDate = act?.StartDate ?? DateTime.Today,
                EndDate = act?.EndDate ?? DateTime.Today,
                IsFixed = act?.IsFixed == true,
                IsPrinted = act?.IsPrinted == true,
                Materials = new ObservableCollection<HiddenWorkActMaterialEntry>((act?.Materials ?? new ObservableCollection<HiddenWorkActMaterialEntry>())
                    .Where(x => x != null)
                    .Select(x => new HiddenWorkActMaterialEntry
                    {
                        IsSelected = x.IsSelected,
                        MaterialName = x.MaterialName,
                        Passport = x.Passport,
                        ArrivalDate = x.ArrivalDate
                    }))
            };
        }

        private void HiddenWorkActsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HiddenWorkActsGrid?.SelectedItem is HiddenWorkActRecord act)
                selectedHiddenWorkAct = act;
            else if (hiddenWorkActRows.Count == 0)
                selectedHiddenWorkAct = null;

            UpdateHiddenWorkActEditorState();
            UpdateHiddenWorkActPreview();
        }

        private void HiddenWorkActFixedCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            if ((sender as FrameworkElement)?.DataContext is not HiddenWorkActRecord act)
                return;

            if (act.IsPrinted)
                act.IsFixed = true;

            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState(saveAfterRefresh: true);
        }

        private void HiddenWorkActPrintedCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            if ((sender as FrameworkElement)?.DataContext is not HiddenWorkActRecord act)
                return;

            if (act.IsPrinted)
                act.IsFixed = true;

            MarkHiddenWorkActStateDirty();
            RefreshHiddenWorkActState(saveAfterRefresh: true);
        }

        private void HiddenWorkActEditorField_LostFocus(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null || isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            SaveHiddenWorkDefaultsFromSelectedAct();
            PersistHiddenWorkActChanges();
        }

        private void SaveHiddenWorkDefaultsFromSelectedAct()
        {
            if (selectedHiddenWorkAct == null || currentObject == null)
                return;

            currentObject.HiddenWorkDefaults ??= new HiddenWorkActDefaults();
            currentObject.HiddenWorkDefaults.FullObjectName = selectedHiddenWorkAct.FullObjectName?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.GeneralContractorInfo = selectedHiddenWorkAct.GeneralContractorInfo?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.SubcontractorInfo = selectedHiddenWorkAct.SubcontractorInfo?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.TechnicalSupervisorInfo = selectedHiddenWorkAct.TechnicalSupervisorInfo?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.ProjectOrganizationInfo = selectedHiddenWorkAct.ProjectOrganizationInfo?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.WorkExecutorInfo = selectedHiddenWorkAct.WorkExecutorInfo?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.ProjectDocumentation = selectedHiddenWorkAct.ProjectDocumentation?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.Deviations = string.Empty;
            currentObject.HiddenWorkDefaults.ContractorSignerName = selectedHiddenWorkAct.ContractorSignerName?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.TechnicalSupervisorSignerName = selectedHiddenWorkAct.TechnicalSupervisorSignerName?.Trim() ?? string.Empty;
            currentObject.HiddenWorkDefaults.ProjectOrganizationSignerName = selectedHiddenWorkAct.ProjectOrganizationSignerName?.Trim() ?? string.Empty;
        }

        private void HiddenWorkActDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (selectedHiddenWorkAct == null || isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            PersistHiddenWorkActChanges();
        }

        private void HiddenWorkActMaterialsGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (selectedHiddenWorkAct == null || isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            Dispatcher.BeginInvoke(new Action(() => PersistHiddenWorkActChanges()), DispatcherPriority.Background);
        }

        private void HiddenWorkActMaterialDate_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (selectedHiddenWorkAct == null || isRefreshingHiddenWorkActs || isUpdatingHiddenWorkActEditor)
                return;

            PersistHiddenWorkActChanges();
        }

        private void ResetHiddenWorkActMaterials_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
                return;

            var rows = GetProductionRowsForHiddenWorkAct(selectedHiddenWorkAct);
            selectedHiddenWorkAct.Materials = new ObservableCollection<HiddenWorkActMaterialEntry>(
                BuildSuggestedHiddenWorkMaterials(selectedHiddenWorkAct, rows, selectedHiddenWorkAct.StartDate));
            RewireHiddenWorkActSubscriptions();
            PersistHiddenWorkActChanges();
        }

        private void AddHiddenWorkActMaterial_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
                return;

            selectedHiddenWorkAct.Materials ??= new ObservableCollection<HiddenWorkActMaterialEntry>();
            var material = new HiddenWorkActMaterialEntry
            {
                IsSelected = true
            };

            selectedHiddenWorkAct.Materials.Add(material);
            if (HiddenWorkActMaterialsGrid != null)
                HiddenWorkActMaterialsGrid.SelectedItem = material;
            PersistHiddenWorkActChanges();
        }

        private void RemoveHiddenWorkActMaterial_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct?.Materials == null)
                return;

            if (HiddenWorkActMaterialsGrid?.SelectedItem is not HiddenWorkActMaterialEntry material)
            {
                MessageBox.Show("Выберите материал в таблице.");
                return;
            }

            selectedHiddenWorkAct.Materials.Remove(material);
            PersistHiddenWorkActChanges();
        }

        private void OpenHiddenWorkActMaterialArrivalPicker_Click(object sender, RoutedEventArgs e)
        {
            if (selectedHiddenWorkAct == null)
                return;

            var material = (sender as FrameworkElement)?.DataContext as HiddenWorkActMaterialEntry
                ?? HiddenWorkActMaterialsGrid?.SelectedItem as HiddenWorkActMaterialEntry;
            if (material == null)
                return;

            var selectedArrival = PromptSelectHiddenWorkArrival(material, selectedHiddenWorkAct);
            if (selectedArrival == null)
                return;

            if (string.IsNullOrWhiteSpace(material.MaterialName))
                material.MaterialName = selectedArrival.MaterialName?.Trim() ?? string.Empty;
            material.Passport = selectedArrival.Passport?.Trim() ?? material.Passport;
            material.ArrivalDate = selectedArrival.Date.Date;
            PersistHiddenWorkActChanges();
        }

        private JournalRecord PromptSelectHiddenWorkArrival(HiddenWorkActMaterialEntry material, HiddenWorkActRecord act)
        {
            var typedName = material?.MaterialName?.Trim() ?? string.Empty;
            var preferredGroup = act?.WorkName?.Trim() ?? string.Empty;

            var pool = journal
                .Where(x => x != null
                    && string.Equals((x.Category ?? string.Empty).Trim(), "Основные", StringComparison.CurrentCultureIgnoreCase)
                    && !string.IsNullOrWhiteSpace(x.MaterialName))
                .ToList();
            if (pool.Count == 0)
            {
                MessageBox.Show("В приходе нет записей по основным материалам.");
                return null;
            }

            var nameMatches = string.IsNullOrWhiteSpace(typedName)
                ? pool
                : pool.Where(x => AreProductionLookupValuesEquivalent(x.MaterialName?.Trim() ?? string.Empty, typedName)).ToList();

            var groupMatches = string.IsNullOrWhiteSpace(preferredGroup)
                ? pool
                : pool.Where(x => AreProductionLookupValuesEquivalent(x.MaterialGroup?.Trim() ?? string.Empty, preferredGroup)).ToList();

            var candidates = nameMatches.Count > 0
                ? nameMatches
                : groupMatches.Count > 0
                    ? groupMatches
                    : pool;

            var items = candidates
                .OrderByDescending(x => x.Date.Date)
                .ThenBy(x => x.MaterialGroup ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .ThenBy(x => x.MaterialName ?? string.Empty, StringComparer.CurrentCultureIgnoreCase)
                .Select(x => new HiddenWorkArrivalPickerItem
                {
                    Record = x,
                    Display = $"{x.Date:dd.MM.yyyy} | {x.MaterialGroup} | {x.MaterialName} | Паспорт: {x.Passport}"
                })
                .ToList();
            if (items.Count == 0)
            {
                MessageBox.Show("Подходящих записей в приходе не найдено.");
                return null;
            }

            var dialog = new Window
            {
                Title = "Выбор прихода материала",
                Owner = this,
                Width = 860,
                Height = 560,
                WindowStartupLocation = WindowStartupLocation.CenterOwner
            };

            var root = new Grid { Margin = new Thickness(14) };
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            dialog.Content = root;

            var searchBox = new TextBox
            {
                Margin = new Thickness(0, 0, 0, 8)
            };
            root.Children.Add(searchBox);

            var list = new ListBox
            {
                DisplayMemberPath = nameof(HiddenWorkArrivalPickerItem.Display)
            };
            Grid.SetRow(list, 1);
            root.Children.Add(list);

            var visibleItems = new ObservableCollection<HiddenWorkArrivalPickerItem>(items);
            list.ItemsSource = visibleItems;

            void ApplyFilter()
            {
                var filter = searchBox.Text?.Trim() ?? string.Empty;
                var filtered = string.IsNullOrWhiteSpace(filter)
                    ? items
                    : items.Where(x => (x.Display ?? string.Empty).Contains(filter, StringComparison.CurrentCultureIgnoreCase)).ToList();

                visibleItems.Clear();
                foreach (var row in filtered)
                    visibleItems.Add(row);
            }

            searchBox.TextChanged += (_, _) => ApplyFilter();

            var footer = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 10, 0, 0)
            };
            var okButton = new Button
            {
                Content = "Выбрать",
                MinWidth = 118,
                IsDefault = true
            };
            var cancelButton = new Button
            {
                Content = "Отмена",
                MinWidth = 110,
                IsCancel = true,
                Margin = new Thickness(8, 0, 0, 0),
                Style = FindResource("SecondaryButton") as Style
            };
            footer.Children.Add(okButton);
            footer.Children.Add(cancelButton);
            Grid.SetRow(footer, 2);
            root.Children.Add(footer);

            HiddenWorkArrivalPickerItem selected = null;
            void ConfirmSelection()
            {
                selected = list.SelectedItem as HiddenWorkArrivalPickerItem;
                if (selected == null)
                    return;

                dialog.DialogResult = true;
            }

            okButton.Click += (_, _) => ConfirmSelection();
            list.MouseDoubleClick += (_, _) => ConfirmSelection();

            return dialog.ShowDialog() == true ? selected?.Record : null;
        }

        private void OpenProductionHiddenWorkActButton_Click(object sender, RoutedEventArgs e)
        {
            if (selectedProductionRow == null || currentObject?.ProductionJournal?.Contains(selectedProductionRow) != true)
            {
                MessageBox.Show("Сначала сохраните запись в ПР или выберите уже сохраненную строку.");
                return;
            }

            OpenHiddenWorkActForProductionRow(selectedProductionRow);
        }

        private void ProductionOpenHiddenWorkAct_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.Tag is ProductionJournalEntry row)
                OpenHiddenWorkActForProductionRow(row);
        }

        private void EnsureArmoringCompanionDecision(ProductionJournalEntry row)
        {
            if (row == null || row.IsGeneratedCompanion || !RequiresArmoringCompanionPrompt(row))
                return;

            if (HasExistingArmoringCompanion(row))
            {
                row.ArmoringPromptHandled = true;
                row.ArmoringCompanionRequested = false;
                return;
            }

            if (row.ArmoringPromptHandled)
                return;

            var result = MessageBox.Show(
                "Добавить вчерашним днем армирование этого участка?",
                "Акт скрытых работ",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            row.ArmoringPromptHandled = true;
            row.ArmoringCompanionRequested = result == MessageBoxResult.Yes;
        }

        private bool RequiresArmoringCompanionPrompt(ProductionJournalEntry row)
        {
            if (row == null)
                return false;

            var actionName = row.ActionName?.Trim() ?? string.Empty;
            var workName = row.WorkName?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(actionName) && string.IsNullOrWhiteSpace(workName))
                return false;

            return actionName.IndexOf("бетонир", StringComparison.CurrentCultureIgnoreCase) >= 0
                || workName.IndexOf("бетонир", StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        private string BuildArmoringCompanionActionName(string actionName)
        {
            return ReplaceBetoningWithArmoring(actionName);
        }

        private string BuildArmoringCompanionWorkName(string workName)
        {
            return ReplaceBetoningWithArmoring(workName);
        }

        private string ReplaceBetoningWithArmoring(string value)
        {
            var normalized = value?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(normalized))
                return string.Empty;

            var replaced = Regex.Replace(normalized, "бетонир\\w*", "армирование", RegexOptions.IgnoreCase).Trim();
            return string.IsNullOrWhiteSpace(replaced) ? normalized : replaced;
        }

        private bool HasExistingArmoringCompanion(ProductionJournalEntry row)
        {
            if (row == null || currentObject?.ProductionJournal == null)
                return false;

            var previousDay = row.Date.AddDays(-1).Date;
            var armoringAction = BuildArmoringCompanionActionName(row.ActionName);
            var armoringWork = BuildArmoringCompanionWorkName(row.WorkName);

            return currentObject.ProductionJournal.Any(x =>
                !ReferenceEquals(x, row)
                && x.Date.Date == previousDay
                && string.Equals((x.ActionName ?? string.Empty).Trim(), armoringAction, StringComparison.CurrentCultureIgnoreCase)
                && string.Equals((x.WorkName ?? string.Empty).Trim(), armoringWork, StringComparison.CurrentCultureIgnoreCase)
                && string.Equals((x.BlocksText ?? string.Empty).Trim(), (row.BlocksText ?? string.Empty).Trim(), StringComparison.CurrentCultureIgnoreCase)
                && string.Equals((x.MarksText ?? string.Empty).Trim(), (row.MarksText ?? string.Empty).Trim(), StringComparison.CurrentCultureIgnoreCase));
        }
    }
}
