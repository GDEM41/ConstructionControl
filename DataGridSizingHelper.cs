using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using System.Windows.Controls.Primitives;
using System.Windows.Threading;

namespace ConstructionControl
{
    public static class DataGridSizingHelper
    {
        private static readonly DependencyPropertyDescriptor ItemsSourceDescriptor =
            DependencyPropertyDescriptor.FromProperty(ItemsControl.ItemsSourceProperty, typeof(DataGrid));

        private static readonly Dictionary<DataGrid, INotifyCollectionChanged> CollectionSubscriptions = new();

        public static readonly DependencyProperty EnableSmartSizingProperty =
            DependencyProperty.RegisterAttached(
                "EnableSmartSizing",
                typeof(bool),
                typeof(DataGridSizingHelper),
                new PropertyMetadata(false, OnEnableSmartSizingChanged));

        public static bool GetEnableSmartSizing(DependencyObject obj) => (bool)obj.GetValue(EnableSmartSizingProperty);

        public static void SetEnableSmartSizing(DependencyObject obj, bool value) => obj.SetValue(EnableSmartSizingProperty, value);

        private static void OnEnableSmartSizingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is not DataGrid grid)
                return;

            if ((bool)e.NewValue)
            {
                grid.Loaded += DataGrid_Loaded;
                grid.Unloaded += DataGrid_Unloaded;
                grid.IsVisibleChanged += DataGrid_IsVisibleChanged;
                ItemsSourceDescriptor?.AddValueChanged(grid, DataGrid_ItemsSourceChanged);

                if (grid.Columns is INotifyCollectionChanged columnsChanged)
                    CollectionChangedEventManager.AddHandler(columnsChanged, DataGrid_ColumnsChanged);

                HookItemsSource(grid);
                ScheduleApply(grid);
            }
            else
            {
                grid.Loaded -= DataGrid_Loaded;
                grid.Unloaded -= DataGrid_Unloaded;
                grid.IsVisibleChanged -= DataGrid_IsVisibleChanged;
                ItemsSourceDescriptor?.RemoveValueChanged(grid, DataGrid_ItemsSourceChanged);

                if (grid.Columns is INotifyCollectionChanged columnsChanged)
                    CollectionChangedEventManager.RemoveHandler(columnsChanged, DataGrid_ColumnsChanged);

                UnhookItemsSource(grid);
            }
        }

        private static void DataGrid_Loaded(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid grid)
                ScheduleApply(grid);
        }

        private static void DataGrid_Unloaded(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid grid)
                UnhookItemsSource(grid);
        }

        private static void DataGrid_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (sender is DataGrid grid && grid.IsVisible)
                ScheduleApply(grid);
        }

        private static void DataGrid_ItemsSourceChanged(object sender, EventArgs e)
        {
            if (sender is not DataGrid grid)
                return;

            HookItemsSource(grid);
            ScheduleApply(grid);
        }

        private static void DataGrid_ColumnsChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (sender is not IEnumerable columns)
                return;

            foreach (var grid in Application.Current?.Windows
                         .OfType<Window>()
                         .SelectMany(FindVisualChildren<DataGrid>)
                         .Where(x => ReferenceEquals(x.Columns, sender))
                     ?? Enumerable.Empty<DataGrid>())
            {
                ScheduleApply(grid);
            }
        }

        private static void HookItemsSource(DataGrid grid)
        {
            UnhookItemsSource(grid);

            if (grid.ItemsSource is INotifyCollectionChanged notifyCollection)
            {
                CollectionSubscriptions[grid] = notifyCollection;
                CollectionChangedEventManager.AddHandler(notifyCollection, DataGrid_ItemsCollectionChanged);
            }
        }

        private static void UnhookItemsSource(DataGrid grid)
        {
            if (CollectionSubscriptions.TryGetValue(grid, out var notifyCollection))
            {
                CollectionChangedEventManager.RemoveHandler(notifyCollection, DataGrid_ItemsCollectionChanged);
                CollectionSubscriptions.Remove(grid);
            }
        }

        private static void DataGrid_ItemsCollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            foreach (var pair in CollectionSubscriptions.Where(x => ReferenceEquals(x.Value, sender)).ToList())
                ScheduleApply(pair.Key);
        }

        private static void ScheduleApply(DataGrid grid)
        {
            if (grid == null)
                return;

            grid.Dispatcher.BeginInvoke(new Action(() => ApplySizing(grid)), DispatcherPriority.Background);
        }

        private static void ApplySizing(DataGrid grid)
        {
            if (grid == null || !grid.IsLoaded || grid.Columns.Count == 0)
                return;

            foreach (var column in grid.Columns)
            {
                var headerText = GetHeaderText(column.Header);
                var headerWidth = MeasureTextWidth(headerText, grid.FontFamily, grid.FontSize, FontWeights.SemiBold) + 28;

                var contentWidth = headerWidth;
                foreach (var item in grid.Items)
                {
                    if (item == null || ReferenceEquals(item, CollectionView.NewItemPlaceholder))
                        continue;

                    foreach (var text in GetColumnTexts(column, item))
                    {
                        var longestWord = GetLongestWord(text);
                        if (string.IsNullOrWhiteSpace(longestWord))
                            continue;

                        var measured = MeasureTextWidth(longestWord, grid.FontFamily, grid.FontSize, FontWeights.Normal) + 24;
                        if (measured > contentWidth)
                            contentWidth = measured;
                    }
                }

                column.MinWidth = Math.Max(40, headerWidth);
                column.Width = new DataGridLength(Math.Max(column.MinWidth, contentWidth));
            }
        }

        private static IEnumerable<string> GetColumnTexts(DataGridColumn column, object item)
        {
            if (TryGetBoundPath(column, out var path))
            {
                var rawValue = GetValueByPath(item, path);
                if (rawValue != null)
                    yield return rawValue.ToString();
            }

            if (column is DataGridTemplateColumn templateColumn)
            {
                foreach (var value in GetTextsFromTemplate(templateColumn.CellTemplate, item))
                    yield return value;

                foreach (var value in GetTextsFromTemplate(templateColumn.CellEditingTemplate, item))
                    yield return value;
            }
        }

        private static bool TryGetBoundPath(DataGridColumn column, out string path)
        {
            path = string.Empty;

            if (column is DataGridBoundColumn boundColumn && boundColumn.Binding is Binding binding)
            {
                path = binding.Path?.Path ?? string.Empty;
                return !string.IsNullOrWhiteSpace(path);
            }

            if (column is DataGridCheckBoxColumn checkBoxColumn && checkBoxColumn.Binding is Binding checkBinding)
            {
                path = checkBinding.Path?.Path ?? string.Empty;
                return !string.IsNullOrWhiteSpace(path);
            }

            return false;
        }

        private static IEnumerable<string> GetTextsFromTemplate(DataTemplate template, object item)
        {
            if (template == null)
                yield break;

            FrameworkElement element;
            try
            {
                element = template.LoadContent() as FrameworkElement;
            }
            catch (InvalidOperationException)
            {
                yield break;
            }
            catch
            {
                yield break;
            }

            if (element == null)
                yield break;

            element.DataContext = item;

            try
            {
                ApplyTemplateRecursive(element);
                element.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                element.Arrange(new Rect(0, 0, element.DesiredSize.Width, element.DesiredSize.Height));
                element.UpdateLayout();
            }
            catch
            {
                yield break;
            }

            foreach (var text in ExtractTexts(element))
            {
                if (!string.IsNullOrWhiteSpace(text))
                    yield return text;
            }
        }

        private static void ApplyTemplateRecursive(FrameworkElement element)
        {
            element.ApplyTemplate();

            foreach (var child in FindVisualChildren<FrameworkElement>(element))
                child.ApplyTemplate();
        }

        private static IEnumerable<string> ExtractTexts(object node)
        {
            switch (node)
            {
                case TextBlock textBlock when !string.IsNullOrWhiteSpace(textBlock.Text):
                    yield return textBlock.Text;
                    break;
                case TextBox textBox when !string.IsNullOrWhiteSpace(textBox.Text):
                    yield return textBox.Text;
                    break;
                case Button button when button.Content != null:
                    yield return button.Content.ToString();
                    break;
                case ContentControl contentControl when contentControl.Content != null && contentControl.Content is string text:
                    yield return text;
                    break;
                case ToggleButton toggleButton when toggleButton.Content != null:
                    yield return toggleButton.Content.ToString();
                    break;
            }

            if (node is DependencyObject dependencyObject)
            {
                for (var i = 0; i < VisualTreeHelper.GetChildrenCount(dependencyObject); i++)
                {
                    var child = VisualTreeHelper.GetChild(dependencyObject, i);
                    foreach (var text in ExtractTexts(child))
                        yield return text;
                }
            }
        }

        private static string GetHeaderText(object header)
        {
            return header switch
            {
                null => string.Empty,
                string text => text,
                TextBlock textBlock => textBlock.Text ?? string.Empty,
                _ => header.ToString() ?? string.Empty
            };
        }

        private static string GetLongestWord(object value)
        {
            var text = value?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            return Regex.Matches(text, @"[^\s,;:/\\|]+")
                .Cast<Match>()
                .Select(x => x.Value)
                .OrderByDescending(x => x.Length)
                .FirstOrDefault() ?? string.Empty;
        }

        private static object GetValueByPath(object source, string path)
        {
            if (source == null || string.IsNullOrWhiteSpace(path))
                return null;

            object current = source;
            foreach (var part in path.Split('.'))
            {
                if (current == null)
                    return null;

                if (current is IDictionary dictionary && dictionary.Contains(part))
                {
                    current = dictionary[part];
                    continue;
                }

                var property = current.GetType().GetProperty(part, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (property == null)
                    return null;

                current = property.GetValue(current);
            }

            return current;
        }

        private static double MeasureTextWidth(string text, FontFamily fontFamily, double fontSize, FontWeight fontWeight)
        {
            if (string.IsNullOrWhiteSpace(text))
                return 0;

            var formattedText = new FormattedText(
                text,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                new Typeface(fontFamily, FontStyles.Normal, fontWeight, FontStretches.Normal),
                fontSize,
                Brushes.Black,
                VisualTreeHelper.GetDpi(Application.Current.MainWindow).PixelsPerDip);

            return formattedText.WidthIncludingTrailingWhitespace;
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject parent) where T : DependencyObject
        {
            if (parent == null)
                yield break;

            for (var i = 0; i < VisualTreeHelper.GetChildrenCount(parent); i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                if (child is T match)
                    yield return match;

                foreach (var descendant in FindVisualChildren<T>(child))
                    yield return descendant;
            }
        }
    }
}
