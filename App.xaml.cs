using System;
using System.Windows;

namespace ConstructionControl
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
            {
                MessageBox.Show(
                    ex.ExceptionObject.ToString(),
                    "UnhandledException");
            };

            DispatcherUnhandledException += (s, ex) =>
            {
                MessageBox.Show(
                    ex.Exception.ToString(),
                    "DispatcherUnhandledException");
                ex.Handled = true;
            };

            base.OnStartup(e);
        }
    }
}
