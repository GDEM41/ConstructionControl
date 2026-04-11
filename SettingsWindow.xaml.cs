using System.Windows;

namespace ConstructionControl
{
    public partial class SettingsWindow : Window
    {
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
            HideReminderDetailsCheckBox.IsChecked = settings.HideReminderDetails;
            SafeStartupModeCheckBox.IsChecked = settings.SafeStartupMode;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            var snoozeMinutes = int.TryParse(ReminderSnoozeMinutesBox.Text?.Trim(), out var parsedMinutes)
                ? parsedMinutes
                : 15;

            if (snoozeMinutes < 1)
                snoozeMinutes = 1;
            if (snoozeMinutes > 240)
                snoozeMinutes = 240;

            var autoSaveMinutes = int.TryParse(AutoSaveIntervalMinutesBox.Text?.Trim(), out var parsedAutoSaveMinutes)
                ? parsedAutoSaveMinutes
                : 5;
            if (autoSaveMinutes < 1)
                autoSaveMinutes = 1;
            if (autoSaveMinutes > 240)
                autoSaveMinutes = 240;

            ResultSettings = new ProjectUiSettings
            {
                DisableTree = DisableTreeCheckBox.IsChecked == true,
                PinTreeByDefault = DisableTreeCheckBox.IsChecked == true ? false : PinTreeCheckBox.IsChecked == true,
                ShowReminderPopup = ReminderPopupCheckBox.IsChecked != false,
                ReminderSnoozeMinutes = snoozeMinutes,
                AutoSaveIntervalMinutes = autoSaveMinutes,
                HideReminderDetails = HideReminderDetailsCheckBox.IsChecked == true,
                SafeStartupMode = SafeStartupModeCheckBox.IsChecked == true
            };

            DialogResult = true;
            Close();
        }
    }
}
