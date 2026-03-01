// MainDialog.cs 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;

namespace ScheduleLink.Views
{
    public class MainDialog : Window
    {
        private ListBox _scheduleList;
        private TextBox _searchBox;
        private TextBlock _countText;
        private StackPanel _paramPanel;
        private TextBlock _paramTitle;
        private TextBlock _paramCount;
        private Button _exportBtn;
        private Button _importBtn;

        private readonly List<ViewSchedule> _allSchedules;
        private readonly Document _doc;

        private static readonly SolidColorBrush BrushInstance = new SolidColorBrush(System.Windows.Media.Color.FromRgb(198, 239, 206));
        private static readonly SolidColorBrush BrushType = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 235, 156));
        private static readonly SolidColorBrush BrushReadOnly = new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 199, 206));

        public ViewSchedule SelectedSchedule { get; private set; }
        public bool IsExportMode { get; private set; }

        public MainDialog(List<ViewSchedule> schedules, Document doc)
        {
            _allSchedules = schedules;
            _doc = doc;

            Title = "ScheduleLink - IB-BIM Tools";
            Width = 800;
            Height = 550;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            MinWidth = 650;
            MinHeight = 400;
            Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(245, 245, 250));

            // Set window icon from embedded resource
            SetWindowIcon();

            BuildUI();
            PopulateScheduleList(_allSchedules);
        }

        private void SetWindowIcon()
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream stream = assembly.GetManifestResourceStream("ScheduleLink.Resources.Schedule_1_32.png");
                if (stream != null)
                {
                    var bitmap = new BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = stream;
                    bitmap.CacheOption = BitmapCacheOption.OnLoad;
                    bitmap.EndInit();
                    Icon = bitmap;
                }
            }
            catch { }
        }

        private void BuildUI()
        {
            var mainGrid = new System.Windows.Controls.Grid { Margin = new Thickness(14) };
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            // Row 0: Title
            var title = new TextBlock
            {
                Text = "ScheduleLink - Export / Import Schedules",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(50, 50, 80))
            };
            System.Windows.Controls.Grid.SetRow(title, 0);
            mainGrid.Children.Add(title);

            // Row 1: Two columns
            var contentGrid = new System.Windows.Controls.Grid();
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 250 });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(5) });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 250 });
            System.Windows.Controls.Grid.SetRow(contentGrid, 1);
            mainGrid.Children.Add(contentGrid);

            var leftPanel = BuildLeftPanel();
            System.Windows.Controls.Grid.SetColumn(leftPanel, 0);
            contentGrid.Children.Add(leftPanel);

            var splitter = new GridSplitter
            {
                Width = 5,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(200, 200, 210))
            };
            System.Windows.Controls.Grid.SetColumn(splitter, 1);
            contentGrid.Children.Add(splitter);

            var rightPanel = BuildRightPanel();
            System.Windows.Controls.Grid.SetColumn(rightPanel, 2);
            contentGrid.Children.Add(rightPanel);

            // Row 2: Legend
            var legend = BuildLegend();
            System.Windows.Controls.Grid.SetRow(legend, 2);
            mainGrid.Children.Add(legend);

            // Row 3: Buttons
            var buttons = BuildButtons();
            System.Windows.Controls.Grid.SetRow(buttons, 3);
            mainGrid.Children.Add(buttons);

            Content = mainGrid;
        }

        private Border BuildLeftPanel()
        {
            var panel = new DockPanel();

            var header = new TextBlock
            {
                Text = "Select Schedule",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(60, 60, 90))
            };
            DockPanel.SetDock(header, Dock.Top);
            panel.Children.Add(header);

            // Search
            var searchBorder = new Border
            {
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 180, 200)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Margin = new Thickness(0, 0, 0, 6),
                Background = Brushes.White
            };
            _searchBox = new TextBox
            {
                Height = 26,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(6, 3, 6, 3),
                Background = Brushes.Transparent,
                Text = "Search...",
                Foreground = Brushes.Gray
            };
            _searchBox.GotFocus += (s, e) =>
            {
                if (_searchBox.Foreground == Brushes.Gray)
                { _searchBox.Text = ""; _searchBox.Foreground = Brushes.Black; }
            };
            _searchBox.LostFocus += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_searchBox.Text))
                { _searchBox.Text = "Search..."; _searchBox.Foreground = Brushes.Gray; }
            };
            _searchBox.TextChanged += SearchBox_TextChanged;
            searchBorder.Child = _searchBox;
            DockPanel.SetDock(searchBorder, Dock.Top);
            panel.Children.Add(searchBorder);

            _countText = new TextBlock
            {
                FontSize = 11,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 4, 0, 0)
            };
            DockPanel.SetDock(_countText, Dock.Bottom);
            panel.Children.Add(_countText);

            var listBorder = new Border
            {
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 180, 200)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Background = Brushes.White
            };
            _scheduleList = new ListBox
            {
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(2)
            };
            _scheduleList.SelectionChanged += ScheduleList_SelectionChanged;
            _scheduleList.MouseDoubleClick += (s, e) => DoExport();
            listBorder.Child = _scheduleList;
            panel.Children.Add(listBorder);

            var wrapper = new Border { Margin = new Thickness(0, 0, 4, 0) };
            wrapper.Child = panel;
            return wrapper;
        }

        private void SearchBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string q = (_searchBox.Text ?? "").Trim();

            // If placeholder or empty, show all schedules
            if (string.IsNullOrEmpty(q) || _searchBox.Foreground == Brushes.Gray)
            {
                PopulateScheduleList(_allSchedules);
            }
            else
            {
                var filtered = _allSchedules.Where(
                    vs => vs.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                PopulateScheduleList(filtered);
            }
        }

        private Border BuildRightPanel()
        {
            var panel = new DockPanel();

            _paramTitle = new TextBlock
            {
                Text = "Parameters",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 6),
                Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(60, 60, 90))
            };
            DockPanel.SetDock(_paramTitle, Dock.Top);
            panel.Children.Add(_paramTitle);

            _paramCount = new TextBlock
            {
                FontSize = 11,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 4, 0, 0)
            };
            DockPanel.SetDock(_paramCount, Dock.Bottom);
            panel.Children.Add(_paramCount);

            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };
            var paramBorder = new Border
            {
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(180, 180, 200)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Background = Brushes.White
            };

            _paramPanel = new StackPanel { Margin = new Thickness(4) };
            scrollViewer.Content = _paramPanel;
            paramBorder.Child = scrollViewer;
            panel.Children.Add(paramBorder);

            _paramPanel.Children.Add(new TextBlock
            {
                Text = "Select a schedule to see parameters",
                Foreground = Brushes.Gray,
                FontSize = 12,
                Margin = new Thickness(10),
                TextWrapping = TextWrapping.Wrap
            });

            var wrapper = new Border { Margin = new Thickness(4, 0, 0, 0) };
            wrapper.Child = panel;
            return wrapper;
        }

        private StackPanel BuildLegend()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 8, 0, 8),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            panel.Children.Add(CreateLegendItem(BrushInstance, "Instance"));
            panel.Children.Add(CreateLegendItem(BrushType, "Type"));
            panel.Children.Add(CreateLegendItem(BrushReadOnly, "Read-only"));
            return panel;
        }

        private StackPanel CreateLegendItem(SolidColorBrush color, string text)
        {
            var sp = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(12, 0, 12, 0) };
            sp.Children.Add(new Border
            {
                Width = 14,
                Height = 14,
                Background = color,
                BorderBrush = new SolidColorBrush(System.Windows.Media.Color.FromRgb(150, 150, 150)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(2),
                Margin = new Thickness(0, 0, 5, 0),
                VerticalAlignment = VerticalAlignment.Center
            });
            sp.Children.Add(new TextBlock { Text = text, FontSize = 12, VerticalAlignment = VerticalAlignment.Center });
            return sp;
        }

        private System.Windows.Controls.Grid BuildButtons()
        {
            var grid = new System.Windows.Controls.Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            _exportBtn = CreateButton("Export to Excel", 140,
                System.Windows.Media.Color.FromRgb(68, 114, 196),
                System.Windows.Media.Color.FromRgb(50, 90, 170));
            _exportBtn.IsEnabled = false;
            _exportBtn.Click += (s, e) => DoExport();
            System.Windows.Controls.Grid.SetColumn(_exportBtn, 1);
            grid.Children.Add(_exportBtn);

            _importBtn = CreateButton("Import from Excel", 140,
                System.Windows.Media.Color.FromRgb(76, 175, 80),
                System.Windows.Media.Color.FromRgb(56, 142, 60));
            _importBtn.Margin = new Thickness(8, 0, 8, 0);
            _importBtn.Click += (s, e) => DoImport();
            System.Windows.Controls.Grid.SetColumn(_importBtn, 2);
            grid.Children.Add(_importBtn);

            var cancelBtn = CreateButton("Cancel", 90,
                System.Windows.Media.Color.FromRgb(230, 230, 235),
                System.Windows.Media.Color.FromRgb(180, 180, 190));
            cancelBtn.Foreground = Brushes.Black;
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            System.Windows.Controls.Grid.SetColumn(cancelBtn, 3);
            grid.Children.Add(cancelBtn);

            return grid;
        }

        private Button CreateButton(string text, double width,
            System.Windows.Media.Color bg, System.Windows.Media.Color border)
        {
            return new Button
            {
                Width = width,
                Height = 36,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Content = text,
                Background = new SolidColorBrush(bg),
                Foreground = Brushes.White,
                BorderBrush = new SolidColorBrush(border),
                BorderThickness = new Thickness(1),
                Cursor = Cursors.Hand
            };
        }

        private void PopulateScheduleList(IEnumerable<ViewSchedule> schedules)
        {
            _scheduleList.SelectionChanged -= ScheduleList_SelectionChanged;
            _scheduleList.Items.Clear();
            int count = 0;
            foreach (var vs in schedules)
            {
                _scheduleList.Items.Add(new ListBoxItem
                {
                    Content = vs.Name,
                    Tag = vs,
                    Padding = new Thickness(6, 4, 6, 4)
                });
                count++;
            }
            _countText.Text = count + " schedules found";
            _exportBtn.IsEnabled = false;
            _scheduleList.SelectedIndex = -1;
            _scheduleList.SelectionChanged += ScheduleList_SelectionChanged;

            // Clear parameter preview
            _paramPanel.Children.Clear();
            _paramPanel.Children.Add(new TextBlock
            {
                Text = "Select a schedule to see parameters",
                Foreground = Brushes.Gray,
                FontSize = 12,
                Margin = new Thickness(10),
                TextWrapping = TextWrapping.Wrap
            });
            _paramTitle.Text = "Parameters";
            _paramCount.Text = "";
        }

        private void ScheduleList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _exportBtn.IsEnabled = _scheduleList.SelectedItem != null;
            if (_scheduleList.SelectedItem is ListBoxItem item && item.Tag is ViewSchedule vs)
                UpdateParameterPreview(vs);
        }

        private void UpdateParameterPreview(ViewSchedule schedule)
        {
            _paramPanel.Children.Clear();

            try
            {
                ScheduleDefinition def = schedule.Definition;
                int fieldCount = def.GetFieldCount();
                int instanceCount = 0, typeCount = 0, readOnlyCount = 0;

                // Get first element for read-only check
                Element firstElem = null;
                try
                {
                    var collector = new FilteredElementCollector(_doc, schedule.Id);
                    var firstId = collector.ToElementIds().FirstOrDefault();
                    if (firstId != null && firstId != ElementId.InvalidElementId)
                        firstElem = _doc.GetElement(firstId);
                }
                catch { }

                for (int i = 0; i < fieldCount; i++)
                {
                    ScheduleField field = def.GetField(i);
                    if (field.IsHidden) continue;

                    string name = field.ColumnHeading;
                    if (string.IsNullOrEmpty(name)) name = field.GetName();

                    ScheduleFieldType fieldType = field.FieldType;
                    bool isReadOnly = (fieldType == ScheduleFieldType.Formula || fieldType == ScheduleFieldType.Count);
                    bool isType = (fieldType == ScheduleFieldType.ElementType);

                    // Check actual read-only status
                    if (!isReadOnly && firstElem != null)
                    {
                        try
                        {
                            Parameter p = firstElem.LookupParameter(name)
                                       ?? firstElem.LookupParameter(field.GetName());
                            if (p != null && p.IsReadOnly) isReadOnly = true;
                        }
                        catch { }
                    }

                    SolidColorBrush bgColor;
                    string typeLabel;
                    if (isReadOnly) { bgColor = BrushReadOnly; typeLabel = "Read-only"; readOnlyCount++; }
                    else if (isType) { bgColor = BrushType; typeLabel = "Type"; typeCount++; }
                    else { bgColor = BrushInstance; typeLabel = "Instance"; instanceCount++; }

                    var paramRow = new Border
                    {
                        Background = bgColor,
                        CornerRadius = new CornerRadius(3),
                        Margin = new Thickness(2, 2, 2, 2),
                        Padding = new Thickness(8, 5, 8, 5)
                    };

                    var rowGrid = new System.Windows.Controls.Grid();
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    var nameText = new TextBlock { Text = name, FontSize = 12, FontWeight = FontWeights.Medium, VerticalAlignment = VerticalAlignment.Center };
                    System.Windows.Controls.Grid.SetColumn(nameText, 0);
                    rowGrid.Children.Add(nameText);

                    var typeText = new TextBlock
                    {
                        Text = typeLabel,
                        FontSize = 10,
                        Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 100, 100)),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(8, 0, 0, 0)
                    };
                    System.Windows.Controls.Grid.SetColumn(typeText, 1);
                    rowGrid.Children.Add(typeText);

                    paramRow.Child = rowGrid;
                    _paramPanel.Children.Add(paramRow);
                }

                int total = instanceCount + typeCount + readOnlyCount;
                _paramTitle.Text = "Parameters - " + schedule.Name;
                _paramCount.Text = total + " parameters: " + instanceCount + " Instance, " + typeCount + " Type, " + readOnlyCount + " Read-only";
            }
            catch (Exception ex)
            {
                _paramPanel.Children.Add(new TextBlock
                {
                    Text = "Error reading parameters:\n" + ex.Message,
                    Foreground = Brushes.Red,
                    FontSize = 11,
                    Margin = new Thickness(8),
                    TextWrapping = TextWrapping.Wrap
                });
            }
        }

        private void DoExport()
        {
            if (_scheduleList.SelectedItem is ListBoxItem item && item.Tag is ViewSchedule vs)
            {
                SelectedSchedule = vs;
                IsExportMode = true;
                DialogResult = true;
                Close();
            }
        }

        private void DoImport()
        {
            SelectedSchedule = null;
            IsExportMode = false;
            DialogResult = true;
            Close();
        }
    }
}
