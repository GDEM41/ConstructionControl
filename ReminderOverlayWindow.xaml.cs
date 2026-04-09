using System;
using System.Windows;
using System.Windows.Controls;

namespace ConstructionControl
{
    public partial class ReminderOverlayWindow : Window
    {
        public ReminderOverlayWindow()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler SnoozeRequested;
        public event RoutedEventHandler ToggleDetailsRequested;

        public FrameworkElement RootElement => RootBorder;
        public ItemsControl SectionsHostElement => SectionsHost;
        public TextBlock StateTextElement => StateText;
        public Button SnoozeButtonElement => SnoozeButton;
        public Button ToggleDetailsButtonElement => ToggleDetailsButton;

        private void SnoozeButton_Click(object sender, RoutedEventArgs e)
            => SnoozeRequested?.Invoke(sender, e);

        private void ToggleDetailsButton_Click(object sender, RoutedEventArgs e)
            => ToggleDetailsRequested?.Invoke(sender, e);
    }
}
