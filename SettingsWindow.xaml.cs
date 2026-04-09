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
            HideReminderDetailsCheckBox.IsChecked = settings.HideReminderDetails;
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

            ResultSettings = new ProjectUiSettings
            {
                DisableTree = DisableTreeCheckBox.IsChecked == true,
                PinTreeByDefault = DisableTreeCheckBox.IsChecked == true ? false : PinTreeCheckBox.IsChecked == true,
                ShowReminderPopup = ReminderPopupCheckBox.IsChecked != false,
                ReminderSnoozeMinutes = snoozeMinutes,
                HideReminderDetails = HideReminderDetailsCheckBox.IsChecked == true
            };

            DialogResult = true;
            Close();
        }
    }
}
