using System.Windows;

namespace ConstructionControl
{
    public partial class CreateObjectWindow : Window
    {
        public string ObjectName { get; private set; }

        public CreateObjectWindow()
        {
            InitializeComponent();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ObjectNameBox.Text))
            {
                MessageBox.Show("Введите название объекта");
                return;
            }

            ObjectName = ObjectNameBox.Text.Trim();
            DialogResult = true;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
