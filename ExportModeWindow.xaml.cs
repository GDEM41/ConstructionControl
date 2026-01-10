using System.Windows;

namespace ConstructionControl
{
    public partial class ExportModeWindow : Window
    {
        public ExportMode Mode { get; private set; }

        public ExportModeWindow()
        {
            InitializeComponent();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (Merged.IsChecked == true)
                Mode = ExportMode.Merged;
            else
                Mode = ExportMode.Detailed;

            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
