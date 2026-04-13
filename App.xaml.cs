using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
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

                var tempLog = Path.Combine(Path.GetTempPath(), "ConstructionControl_startup.log");
                File.AppendAllText(tempLog, text, Encoding.UTF8);
            }
            catch
            {
                // ignore logging errors
            }
        }

        protected override async void OnStartup(StartupEventArgs e)
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
                // Пока показывается splash, не позволяем приложению завершиться
                // из-за временного отсутствия главного окна.
                ShutdownMode = ShutdownMode.OnExplicitShutdown;

                var splash = new SplashWindow();
                splash.Show();

                await Task.Delay(TimeSpan.FromSeconds(3));

                var mainWindow = new MainWindow
                {
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ShowInTaskbar = true,
                    WindowState = WindowState.Normal
                };

                MainWindow = mainWindow;
                mainWindow.Show();
                mainWindow.Activate();
                mainWindow.Focus();

                ShutdownMode = ShutdownMode.OnMainWindowClose;

                try { splash.Close(); } catch { }
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
