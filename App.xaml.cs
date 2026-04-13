using System;
using System.IO;
using System.Text;
using System.Windows;

namespace ConstructionControl
{
    public partial class App : Application
    {
        private static readonly object StartupLogSync = new();

        private static void AppendStartupLog(string scope, Exception ex)
        {
            try
            {
                var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                var folder = Path.Combine(appData, "ConstructionControl");
                Directory.CreateDirectory(folder);
                var logPath = Path.Combine(folder, "startup.log");
                var text = new StringBuilder()
                    .AppendLine("========================================")
                    .AppendLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                    .AppendLine(scope)
                    .AppendLine(ex.ToString())
                    .ToString();

                lock (StartupLogSync)
                {
                    File.AppendAllText(logPath, text, Encoding.UTF8);
                }
            }
            catch
            {
                // ignore logging errors
            }
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            AppDomain.CurrentDomain.UnhandledException += (s, ex) =>
            {
                if (ex.ExceptionObject is Exception exception)
                    AppendStartupLog("AppDomain.CurrentDomain.UnhandledException", exception);

                MessageBox.Show(ex.ExceptionObject.ToString(), "UnhandledException");
            };

            DispatcherUnhandledException += (s, ex) =>
            {
                AppendStartupLog("Application.DispatcherUnhandledException", ex.Exception);
                MessageBox.Show(ex.Exception.ToString(), "DispatcherUnhandledException");
                ex.Handled = false;
            };

            base.OnStartup(e);

            try
            {
                var mainWindow = new MainWindow();
                MainWindow = mainWindow;
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                AppendStartupLog("App.OnStartup -> MainWindow()", ex);
                MessageBox.Show(
                    $"Failed to open main window.\n\n{ex}",
                    "Startup error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                Shutdown(-1);
            }
        }
    }
}
