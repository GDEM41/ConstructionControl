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
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            ResultSettings = new ProjectUiSettings
            {
                DisableTree = DisableTreeCheckBox.IsChecked == true,
                PinTreeByDefault = DisableTreeCheckBox.IsChecked == true ? false : PinTreeCheckBox.IsChecked == true,
                ShowReminderPopup = ReminderPopupCheckBox.IsChecked != false
            };

            DialogResult = true;
            Close();
        }
    }
}
