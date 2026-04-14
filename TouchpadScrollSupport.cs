using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace ConstructionControl
{
    internal static class TouchpadScrollSupport
    {
        private const int WM_MOUSEHWHEEL = 0x020E;
        private const double WheelDelta = 120.0;
        private const double MinScrollStep = 48.0;
        private const double MaxScrollStep = 160.0;
        private static bool initialized;

        public static void Initialize()
        {
            if (initialized)
                return;

            initialized = true;
            EventManager.RegisterClassHandler(
                typeof(Window),
                UIElement.PreviewMouseWheelEvent,
                new MouseWheelEventHandler(HandlePreviewMouseWheel),
                handledEventsToo: true);
            ComponentDispatcher.ThreadPreprocessMessage += HandleThreadPreprocessMessage;
        }

        private static void HandlePreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == 0)
                return;

            var originalSource = e.OriginalSource as DependencyObject;
            var scrollViewer = FindScrollableViewer(originalSource, horizontal: true);
            if (scrollViewer == null)
                return;

            if (TryScrollHorizontally(scrollViewer, e.Delta, useNaturalHorizontalDelta: false))
                e.Handled = true;
        }

        private static void HandleThreadPreprocessMessage(ref MSG msg, ref bool handled)
        {
            if (handled || msg.message != WM_MOUSEHWHEEL)
                return;

            var hwnd = msg.hwnd;
            var source = HwndSource.FromHwnd(hwnd);
            if (source?.RootVisual == null)
                return;

            var window = source.RootVisual as Window
                ?? Application.Current?.Windows
                    .OfType<Window>()
                    .FirstOrDefault(w => new WindowInteropHelper(w).Handle == hwnd);
            if (window == null)
                return;

            var point = window.PointFromScreen(GetScreenPoint(msg.lParam));
            if (point.X < 0 || point.Y < 0 || point.X > window.ActualWidth || point.Y > window.ActualHeight)
                return;

            var originalSource = window.InputHitTest(point) as DependencyObject;
            var scrollViewer = FindScrollableViewer(originalSource, horizontal: true);
            if (scrollViewer == null)
                return;

            if (TryScrollHorizontally(scrollViewer, GetWheelDelta(msg.wParam), useNaturalHorizontalDelta: true))
                handled = true;
        }

        private static ScrollViewer? FindScrollableViewer(DependencyObject? source, bool horizontal)
        {
            for (var current = source; current != null; current = GetParent(current))
            {
                if (current is ScrollViewer viewer && CanScroll(viewer, horizontal))
                    return viewer;

                if (current is DataGrid
                    || current is ListBox
                    || current is ListView
                    || current is TreeView
                    || current is TextBoxBase
                    || current is ItemsControl)
                {
                    var nestedViewer = FindDescendantScrollViewer(current, horizontal);
                    if (nestedViewer != null)
                        return nestedViewer;
                }
            }

            if (Mouse.DirectlyOver is DependencyObject mouseOver && !ReferenceEquals(mouseOver, source))
            {
                var viewer = FindScrollableViewer(mouseOver, horizontal);
                if (viewer != null)
                    return viewer;
            }

            if (Keyboard.FocusedElement is DependencyObject focused && !ReferenceEquals(focused, source))
                return FindScrollableViewer(focused, horizontal);

            return null;
        }

        private static ScrollViewer? FindDescendantScrollViewer(DependencyObject root, bool horizontal)
        {
            if (root == null)
                return null;

            var queue = new System.Collections.Generic.Queue<DependencyObject>();
            queue.Enqueue(root);

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                if (current is ScrollViewer viewer && CanScroll(viewer, horizontal))
                    return viewer;

                var childrenCount = GetChildrenCount(current);
                for (var i = 0; i < childrenCount; i++)
                {
                    var child = VisualTreeHelper.GetChild(current, i);
                    if (child != null)
                        queue.Enqueue(child);
                }
            }

            return null;
        }

        private static int GetChildrenCount(DependencyObject current)
        {
            if (current is Visual || current is Visual3D)
                return VisualTreeHelper.GetChildrenCount(current);

            return 0;
        }

        private static DependencyObject? GetParent(DependencyObject current)
        {
            if (current is Visual || current is Visual3D)
                return VisualTreeHelper.GetParent(current);

            if (current is FrameworkContentElement contentElement)
                return contentElement.Parent;

            return null;
        }

        private static bool CanScroll(ScrollViewer viewer, bool horizontal)
        {
            if (viewer == null)
                return false;

            return horizontal
                ? viewer.ScrollableWidth > 0.5
                : viewer.ScrollableHeight > 0.5;
        }

        private static bool TryScrollHorizontally(ScrollViewer viewer, int delta, bool useNaturalHorizontalDelta)
        {
            if (!CanScroll(viewer, horizontal: true) || delta == 0)
                return false;

            var currentOffset = viewer.HorizontalOffset;
            var step = GetScrollStep(viewer.ViewportWidth);
            var deltaFactor = Math.Max(1.0, Math.Abs(delta) / WheelDelta);
            var signedStep = step * deltaFactor;

            double targetOffset;
            if (useNaturalHorizontalDelta)
            {
                targetOffset = currentOffset + (delta > 0 ? signedStep : -signedStep);
            }
            else
            {
                targetOffset = currentOffset + (delta > 0 ? -signedStep : signedStep);
            }

            targetOffset = Math.Max(0, Math.Min(viewer.ScrollableWidth, targetOffset));
            if (Math.Abs(targetOffset - currentOffset) < 0.1)
                return false;

            viewer.ScrollToHorizontalOffset(targetOffset);
            return true;
        }

        private static double GetScrollStep(double viewportWidth)
        {
            if (double.IsNaN(viewportWidth) || viewportWidth <= 0)
                return MinScrollStep;

            return Math.Clamp(viewportWidth / 10.0, MinScrollStep, MaxScrollStep);
        }

        private static Point GetScreenPoint(IntPtr lParam)
        {
            var value = lParam.ToInt64();
            var x = unchecked((short)(value & 0xFFFF));
            var y = unchecked((short)((value >> 16) & 0xFFFF));
            return new Point(x, y);
        }

        private static int GetWheelDelta(IntPtr wParam)
        {
            var value = wParam.ToInt64();
            return unchecked((short)((value >> 16) & 0xFFFF));
        }
    }
}
