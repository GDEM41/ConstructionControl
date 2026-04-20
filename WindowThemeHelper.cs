using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace ConstructionControl
{
    internal static class WindowThemeHelper
    {
        private const int DwmaUseImmersiveDarkMode = 20;
        private const int DwmaUseImmersiveDarkModeLegacy = 19;
        private const int DwmaBorderColor = 34;
        private const int DwmaCaptionColor = 35;
        private const int DwmaTextColor = 36;

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attribute, ref int value, int valueSize);

        internal static void ApplyToWindow(Window window)
        {
            if (window == null || window.AllowsTransparency || window.WindowStyle == WindowStyle.None)
                return;

            var handle = new WindowInteropHelper(window).Handle;
            if (handle == IntPtr.Zero)
                return;

            var surfaceColor = GetThemeColor("SurfaceBrush", "#111827");
            var strokeColor = GetThemeColor("StrokeBrush", "#334155");
            var textColor = GetThemeColor("TextBrush", "#E5E7EB");
            var darkMode = GetPerceivedBrightness(surfaceColor) < 0.5;

            TrySetIntAttribute(handle, DwmaUseImmersiveDarkMode, darkMode ? 1 : 0);
            TrySetIntAttribute(handle, DwmaUseImmersiveDarkModeLegacy, darkMode ? 1 : 0);
            TrySetIntAttribute(handle, DwmaCaptionColor, ToColorRef(surfaceColor));
            TrySetIntAttribute(handle, DwmaBorderColor, ToColorRef(strokeColor));
            TrySetIntAttribute(handle, DwmaTextColor, ToColorRef(textColor));
        }

        internal static void ApplyToAllOpenWindows()
        {
            if (Application.Current == null)
                return;

            foreach (Window window in Application.Current.Windows)
                ApplyToWindow(window);
        }

        private static Color GetThemeColor(string resourceKey, string fallbackColor)
        {
            if (Application.Current?.Resources.Contains(resourceKey) == true
                && Application.Current.Resources[resourceKey] is SolidColorBrush brush)
                return brush.Color;

            try
            {
                return (Color)ColorConverter.ConvertFromString(fallbackColor);
            }
            catch
            {
                return Colors.Black;
            }
        }

        private static double GetPerceivedBrightness(Color color)
        {
            return ((0.299 * color.R) + (0.587 * color.G) + (0.114 * color.B)) / 255d;
        }

        private static int ToColorRef(Color color)
        {
            return color.R | (color.G << 8) | (color.B << 16);
        }

        private static void TrySetIntAttribute(IntPtr hwnd, int attribute, int value)
        {
            try
            {
                DwmSetWindowAttribute(hwnd, attribute, ref value, Marshal.SizeOf<int>());
            }
            catch
            {
                // Ignore on unsupported systems.
            }
        }
    }
}
