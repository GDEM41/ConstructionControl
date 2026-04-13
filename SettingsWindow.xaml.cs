using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using WinForms = System.Windows.Forms;

namespace ConstructionControl
{
    public partial class SettingsWindow : Window
    {
        private static readonly Regex DigitsRegex = new(@"^\d+$", RegexOptions.Compiled);

        public ProjectUiSettings ResultSettings { get; private set; }

        public SettingsWindow(ProjectUiSettings source)
        {
            InitializeComponent();

            var settings = source ?? new ProjectUiSettings();
            DisableTreeCheckBox.IsChecked = settings.DisableTree;
            PinTreeCheckBox.IsChecked = settings.PinTreeByDefault;
            ReminderPopupCheckBox.IsChecked = settings.ShowReminderPopup;
            ReminderSnoozeMinutesBox.Text = settings.ReminderSnoozeMinutes > 0
                ? settings.ReminderSnoozeMinutes.ToString()
                : "15";
            AutoSaveIntervalMinutesBox.Text = settings.AutoSaveIntervalMinutes > 0
                ? settings.AutoSaveIntervalMinutes.ToString()
                : "5";
            DataRootPathBox.Text = ResolveDataRootPath(settings.DataRootDirectory);
            HideReminderDetailsCheckBox.IsChecked = settings.HideReminderDetails;
            SafeStartupModeCheckBox.IsChecked = settings.SafeStartupMode;
            SummaryReminderOverageCheckBox.IsChecked = settings.SummaryReminderOnOverage;
            SummaryReminderDeficitCheckBox.IsChecked = settings.SummaryReminderOnDeficit;
            SummaryReminderMainOnlyCheckBox.IsChecked = settings.SummaryReminderOnlyMain;
            CheckUpdatesOnStartCheckBox.IsChecked = settings.CheckUpdatesOnStart;
            UpdateFeedUrlBox.Text = settings.UpdateFeedUrl ?? string.Empty;
            RequireCodeForCriticalOperationsCheckBox.IsChecked = settings.RequireCodeForCriticalOperations;

            SelectDensityMode(NormalizeDensityMode(settings.UiDensityMode));
            SelectAccessRole(NormalizeAccessRole(settings.AccessRole));
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            var snoozeMinutes = ParseClampedMinutes(ReminderSnoozeMinutesBox.Text, 15);
            var autoSaveMinutes = ParseClampedMinutes(AutoSaveIntervalMinutesBox.Text, 5);

            ResultSettings = new ProjectUiSettings
            {
                DisableTree = DisableTreeCheckBox.IsChecked == true,
                PinTreeByDefault = DisableTreeCheckBox.IsChecked == true ? false : PinTreeCheckBox.IsChecked == true,
                ShowReminderPopup = ReminderPopupCheckBox.IsChecked != false,
                ReminderSnoozeMinutes = snoozeMinutes,
                AutoSaveIntervalMinutes = autoSaveMinutes,
                HideReminderDetails = HideReminderDetailsCheckBox.IsChecked == true,
                SafeStartupMode = SafeStartupModeCheckBox.IsChecked == true,
                SummaryReminderOnOverage = SummaryReminderOverageCheckBox.IsChecked == true,
                SummaryReminderOnDeficit = SummaryReminderDeficitCheckBox.IsChecked == true,
                SummaryReminderOnlyMain = SummaryReminderMainOnlyCheckBox.IsChecked != false,
                DataRootDirectory = ResolveDataRootPath(DataRootPathBox.Text),
                CheckUpdatesOnStart = CheckUpdatesOnStartCheckBox.IsChecked == true,
                UpdateFeedUrl = (UpdateFeedUrlBox.Text ?? string.Empty).Trim(),
                UiDensityMode = GetSelectedDensityMode(),
                AccessRole = GetSelectedAccessRole(),
                RequireCodeForCriticalOperations = RequireCodeForCriticalOperationsCheckBox.IsChecked != false
            };

            DialogResult = true;
            Close();
        }

        private static string GetDefaultDataRootPath()
        {
            var appDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var root = Path.Combine(appDataFolder, "ConstructionControl", "Data");
            Directory.CreateDirectory(root);
            return root;
        }

        private static string ResolveDataRootPath(string rawPath)
        {
            var candidate = string.IsNullOrWhiteSpace(rawPath)
                ? GetDefaultDataRootPath()
                : Environment.ExpandEnvironmentVariables(rawPath.Trim());

            try
            {
                var fullPath = Path.GetFullPath(candidate);
                Directory.CreateDirectory(fullPath);
                return fullPath;
            }
            catch
            {
                return GetDefaultDataRootPath();
            }
        }

        private static int ParseClampedMinutes(string input, int fallback)
        {
            var value = int.TryParse(input?.Trim(), out var parsed) ? parsed : fallback;
            if (value < 1)
                value = 1;
            if (value > 240)
                value = 240;
            return value;
        }

        private static string NormalizeDensityMode(string mode)
        {
            var normalized = (mode ?? string.Empty).Trim();
            if (string.Equals(normalized, "Компактный", StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("РљРѕРјРї", StringComparison.Ordinal))
            {
                return "Компактный";
            }

            return "Стандартный";
        }

        private static string NormalizeAccessRole(string role)
        {
            var normalized = (role ?? string.Empty).Trim();
            if (string.Equals(normalized, ProjectAccessRoles.View, StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("РџСЂРѕСЃ", StringComparison.OrdinalIgnoreCase))
            {
                return ProjectAccessRoles.View;
            }

            if (string.Equals(normalized, ProjectAccessRoles.Edit, StringComparison.CurrentCultureIgnoreCase)
                || normalized.Contains("Р РµРґР°РєС‚", StringComparison.OrdinalIgnoreCase))
            {
                return ProjectAccessRoles.Edit;
            }

            return ProjectAccessRoles.Critical;
        }

        private void SelectDensityMode(string mode)
        {
            DensityModeBox.SelectedIndex = string.Equals(NormalizeDensityMode(mode), "Компактный", StringComparison.CurrentCultureIgnoreCase)
                ? 1
                : 0;
        }

        private string GetSelectedDensityMode()
        {
            if (DensityModeBox.SelectedItem is ComboBoxItem item)
            {
                var value = item.Content?.ToString()?.Trim();
                return string.Equals(value, "Компактный", StringComparison.CurrentCultureIgnoreCase)
                    ? "Компактный"
                    : "Стандартный";
            }

            return "Стандартный";
        }

        private void SelectAccessRole(string role)
        {
            var normalized = NormalizeAccessRole(role);
            AccessRoleBox.SelectedIndex = normalized switch
            {
                ProjectAccessRoles.View => 0,
                ProjectAccessRoles.Edit => 1,
                _ => 2
            };
        }

        private string GetSelectedAccessRole()
        {
            if (AccessRoleBox.SelectedItem is ComboBoxItem item)
            {
                var value = item.Content?.ToString()?.Trim();
                if (string.Equals(value, ProjectAccessRoles.View, StringComparison.CurrentCultureIgnoreCase))
                    return ProjectAccessRoles.View;
                if (string.Equals(value, ProjectAccessRoles.Edit, StringComparison.CurrentCultureIgnoreCase))
                    return ProjectAccessRoles.Edit;
            }

            return ProjectAccessRoles.Critical;
        }

        private void NumericTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !DigitsRegex.IsMatch(e.Text ?? string.Empty);
        }

        private void BrowseDataRootButton_Click(object sender, RoutedEventArgs e)
        {
            using var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "Выберите папку для кэша, автосохранений и служебных файлов",
                UseDescriptionForTitle = true,
                AutoUpgradeEnabled = true,
                SelectedPath = ResolveDataRootPath(DataRootPathBox.Text)
            };

            if (dialog.ShowDialog() == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                DataRootPathBox.Text = ResolveDataRootPath(dialog.SelectedPath);
        }
    }
}
