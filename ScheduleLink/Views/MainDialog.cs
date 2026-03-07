// MainDialog.cs - ScheduleLink - Improved UI Design
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Xml.Linq;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Effects;
using System.Windows.Media.Imaging;
using Autodesk.Revit.DB;
using WpfColor = System.Windows.Media.Color;
using WpfGrid = System.Windows.Controls.Grid;

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

        // ==================== COLOR PALETTE ====================
        private static readonly WpfColor ThemePrimaryDark = WpfColor.FromRgb(27, 42, 74);       // #1B2A4A
        private static readonly WpfColor ThemePrimaryMedium = WpfColor.FromRgb(44, 62, 80);      // #2C3E50
        private static readonly WpfColor ThemePrimaryLight = WpfColor.FromRgb(52, 73, 94);       // #34495E
        private static readonly WpfColor ThemeAccentBlue = WpfColor.FromRgb(52, 152, 219);       // #3498DB
        private static readonly WpfColor ThemeAccentGreen = WpfColor.FromRgb(39, 174, 96);       // #27AE60
        private static readonly WpfColor ThemeAccentPurple = WpfColor.FromRgb(142, 68, 173);     // #8E44AD
        private static readonly WpfColor ThemeAccentRed = WpfColor.FromRgb(192, 57, 43);         // #C0392B
        private static readonly WpfColor ThemeAccentOrange = WpfColor.FromRgb(230, 126, 34);     // #E67E22
        private static readonly WpfColor ThemeBgLight = WpfColor.FromRgb(240, 242, 245);         // #F0F2F5
        private static readonly WpfColor ThemeBorderLight = WpfColor.FromRgb(224, 228, 232);     // #E0E4E8
        private static readonly WpfColor ThemeTextMuted = WpfColor.FromRgb(123, 143, 163);       // #7B8FA3

        private static readonly SolidColorBrush BrushInstance = new SolidColorBrush(WpfColor.FromRgb(198, 239, 206));
        private static readonly SolidColorBrush BrushType = new SolidColorBrush(WpfColor.FromRgb(255, 235, 156));
        private static readonly SolidColorBrush BrushReadOnly = new SolidColorBrush(WpfColor.FromRgb(255, 199, 206));

        public ViewSchedule SelectedSchedule { get; private set; }
        public bool IsExportMode { get; private set; }

        public MainDialog(List<ViewSchedule> schedules, Document doc)
        {
            _allSchedules = schedules;
            _doc = doc;

            Title = "ScheduleLink - IB-BIM Tools";
            Width = 820;
            Height = 580;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            MinWidth = 680;
            MinHeight = 450;
            Background = new SolidColorBrush(ThemeBgLight);

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
            var mainGrid = new WpfGrid();
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // Header
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // Legend
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });   // Footer

            // Row 0: Header
            var header = BuildHeader();
            WpfGrid.SetRow(header, 0);
            mainGrid.Children.Add(header);

            // Row 1: Two columns with content
            var contentGrid = new WpfGrid { Margin = new Thickness(14, 10, 14, 0) };
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 250 });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(5) });
            contentGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star), MinWidth = 250 });
            WpfGrid.SetRow(contentGrid, 1);
            mainGrid.Children.Add(contentGrid);

            var leftPanel = BuildLeftPanel();
            WpfGrid.SetColumn(leftPanel, 0);
            contentGrid.Children.Add(leftPanel);

            var splitter = new GridSplitter
            {
                Width = 5,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Stretch,
                Background = new SolidColorBrush(ThemeBorderLight)
            };
            WpfGrid.SetColumn(splitter, 1);
            contentGrid.Children.Add(splitter);

            var rightPanel = BuildRightPanel();
            WpfGrid.SetColumn(rightPanel, 2);
            contentGrid.Children.Add(rightPanel);

            // Row 2: Legend
            var legend = BuildLegend();
            WpfGrid.SetRow(legend, 2);
            mainGrid.Children.Add(legend);

            // Row 3: Footer with buttons
            var footer = BuildFooter();
            WpfGrid.SetRow(footer, 3);
            mainGrid.Children.Add(footer);

            Content = mainGrid;
        }

        // ==================== HEADER ====================
        private Border BuildHeader()
        {
            var headerBorder = new Border
            {
                Padding = new Thickness(16, 12, 16, 12),
                BorderThickness = new Thickness(0, 0, 0, 1),
                BorderBrush = new SolidColorBrush(WpfColor.FromRgb(26, 34, 54)),
                Background = new LinearGradientBrush(
                    ThemePrimaryDark, ThemePrimaryLight,
                    new System.Windows.Point(0, 0), new System.Windows.Point(1, 0))
            };

            var headerGrid = new WpfGrid();
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            headerGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Left: Title + Version
            var leftStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            var titleStack = new StackPanel { VerticalAlignment = VerticalAlignment.Center };
            var titleRow = new StackPanel { Orientation = Orientation.Horizontal };
            titleRow.Children.Add(new TextBlock
            {
                Text = "ScheduleLink",
                FontSize = 17,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });

            // Version text from .addin
            titleRow.Children.Add(new TextBlock
            {
                Text = "  v" + GetVersionFromAddin(),
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeTextMuted),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 2, 0, 0)
            });
            titleStack.Children.Add(titleRow);

            titleStack.Children.Add(new TextBlock
            {
                Text = "Export / Import Schedules",
                FontSize = 12,
                Foreground = new SolidColorBrush(WpfColor.FromRgb(160, 175, 190)),
                Margin = new Thickness(0, 2, 0, 0)
            });
            leftStack.Children.Add(titleStack);

            WpfGrid.SetColumn(leftStack, 0);
            headerGrid.Children.Add(leftStack);

            // Center: Credit
            var credit = new TextBlock
            {
                Text = "\u00A9 IB-BIM  \u2022  Itzik Bejarano",
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeTextMuted),
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            WpfGrid.SetColumn(credit, 1);
            headerGrid.Children.Add(credit);

            // Right: Help button
            var rightStack = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };

            var helpBtn = CreateHeaderButton("Help", ThemeAccentBlue);
            helpBtn.ToolTip = "Open online documentation";
            helpBtn.Click += (s, e) => OpenUrl("https://itzikb49.github.io/IB-BIM-ScheduleLink-Docs/UserGuide.html");
            rightStack.Children.Add(helpBtn);

            WpfGrid.SetColumn(rightStack, 2);
            headerGrid.Children.Add(rightStack);

            headerBorder.Child = headerGrid;
            return headerBorder;
        }

        private Button CreateHeaderButton(string text, WpfColor bgColor)
        {
            var btn = new Button
            {
                Height = 30,
                Padding = new Thickness(12, 4, 12, 4),
                Background = new SolidColorBrush(bgColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                FontSize = 11,
                FontWeight = FontWeights.SemiBold,
                Content = text
            };
            btn.MouseEnter += (s, e) => btn.Opacity = 0.85;
            btn.MouseLeave += (s, e) => btn.Opacity = 1.0;
            return btn;
        }

        // ==================== LEFT PANEL ====================
        private Border BuildLeftPanel()
        {
            var panel = new DockPanel();

            var header = new TextBlock
            {
                Text = "Select Schedule",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = new SolidColorBrush(ThemePrimaryMedium)
            };
            DockPanel.SetDock(header, Dock.Top);
            panel.Children.Add(header);

            // Search box with icon
            var searchBorder = new Border
            {
                BorderBrush = new SolidColorBrush(ThemeBorderLight),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 0, 8),
                Background = Brushes.White
            };

            _searchBox = new TextBox
            {
                Height = 30,
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(8, 4, 8, 4),
                Background = Brushes.Transparent,
                Text = "Search...",
                Foreground = Brushes.Gray,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            _searchBox.GotFocus += (s, e) =>
            {
                if (_searchBox.Foreground == Brushes.Gray)
                { _searchBox.Text = ""; _searchBox.Foreground = Brushes.Black; }
                searchBorder.BorderBrush = new SolidColorBrush(ThemeAccentBlue);
                searchBorder.BorderThickness = new Thickness(1.5);
            };
            _searchBox.LostFocus += (s, e) =>
            {
                searchBorder.BorderBrush = new SolidColorBrush(ThemeBorderLight);
                searchBorder.BorderThickness = new Thickness(1);
                if (string.IsNullOrWhiteSpace(_searchBox.Text))
                {
                    _searchBox.TextChanged -= SearchBox_TextChanged;
                    _searchBox.Text = "Search...";
                    _searchBox.Foreground = Brushes.Gray;
                    _searchBox.TextChanged += SearchBox_TextChanged;
                }
            };
            _searchBox.TextChanged += SearchBox_TextChanged;

            searchBorder.Child = _searchBox;
            DockPanel.SetDock(searchBorder, Dock.Top);
            panel.Children.Add(searchBorder);

            _countText = new TextBlock
            {
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeTextMuted),
                Margin = new Thickness(2, 6, 0, 0)
            };
            DockPanel.SetDock(_countText, Dock.Bottom);
            panel.Children.Add(_countText);

            var listBorder = new Border
            {
                BorderBrush = new SolidColorBrush(ThemeBorderLight),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Background = Brushes.White
            };
            _scheduleList = new ListBox
            {
                FontSize = 12,
                BorderThickness = new Thickness(0),
                Padding = new Thickness(2),
                Background = Brushes.Transparent
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
            if (string.IsNullOrEmpty(q) || _searchBox.Foreground == Brushes.Gray)
                PopulateScheduleList(_allSchedules);
            else
            {
                var filtered = _allSchedules.Where(
                    vs => vs.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
                PopulateScheduleList(filtered);
            }
        }

        // ==================== RIGHT PANEL ====================
        private Border BuildRightPanel()
        {
            var panel = new DockPanel();

            _paramTitle = new TextBlock
            {
                Text = "Parameters",
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 8),
                Foreground = new SolidColorBrush(ThemePrimaryMedium)
            };
            DockPanel.SetDock(_paramTitle, Dock.Top);
            panel.Children.Add(_paramTitle);

            _paramCount = new TextBlock
            {
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeTextMuted),
                Margin = new Thickness(2, 6, 0, 0)
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
                BorderBrush = new SolidColorBrush(ThemeBorderLight),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Background = Brushes.White
            };

            _paramPanel = new StackPanel { Margin = new Thickness(4) };
            scrollViewer.Content = _paramPanel;
            paramBorder.Child = scrollViewer;
            panel.Children.Add(paramBorder);

            _paramPanel.Children.Add(new TextBlock
            {
                Text = "Select a schedule to see parameters",
                Foreground = new SolidColorBrush(ThemeTextMuted),
                FontSize = 12,
                Margin = new Thickness(10),
                TextWrapping = TextWrapping.Wrap
            });

            var wrapper = new Border { Margin = new Thickness(4, 0, 0, 0) };
            wrapper.Child = panel;
            return wrapper;
        }

        // ==================== LEGEND ====================
        private StackPanel BuildLegend()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(14, 8, 14, 4),
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
                BorderBrush = new SolidColorBrush(WpfColor.FromRgb(180, 180, 180)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(3),
                Margin = new Thickness(0, 0, 5, 0),
                VerticalAlignment = VerticalAlignment.Center
            });
            sp.Children.Add(new TextBlock
            {
                Text = text,
                FontSize = 12,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                VerticalAlignment = VerticalAlignment.Center
            });
            return sp;
        }

        // ==================== FOOTER ====================
        private Border BuildFooter()
        {
            var footerBorder = new Border
            {
                Padding = new Thickness(16, 10, 16, 10),
                Margin = new Thickness(0),
                Background = new LinearGradientBrush(
                    ThemePrimaryMedium, ThemePrimaryLight,
                    new System.Windows.Point(0, 0), new System.Windows.Point(1, 0))
            };

            var grid = new WpfGrid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            // Left: discreet "More IB-BIM Apps" link
            var moreAppsLink = new TextBlock
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 0, 0)
            };
            var hyperlink = new System.Windows.Documents.Hyperlink
            {
                Foreground = new SolidColorBrush(WpfColor.FromRgb(241, 196, 15)),  // Yellow #F1C40F
                TextDecorations = null // no underline by default
            };
            hyperlink.Inlines.Add("More IB-BIM Apps...");
            hyperlink.Click += (s, e) => OpenUrl("https://apps.autodesk.com/en/Publisher/PublisherHomepage?ID=47V2CLG9KMHPGAHB");
            hyperlink.MouseEnter += (s, e) => hyperlink.Foreground = new SolidColorBrush(WpfColor.FromRgb(255, 224, 102));
            hyperlink.MouseLeave += (s, e) => hyperlink.Foreground = new SolidColorBrush(WpfColor.FromRgb(241, 196, 15));
            moreAppsLink.Inlines.Add(hyperlink);
            moreAppsLink.FontSize = 11;
            WpfGrid.SetColumn(moreAppsLink, 0);
            grid.Children.Add(moreAppsLink);

            _exportBtn = CreateFooterButton("Export to Excel", 145, ThemeAccentBlue);
            _exportBtn.IsEnabled = false;
            _exportBtn.ToolTip = "Export selected schedule to Excel";
            _exportBtn.Click += (s, e) => DoExport();
            WpfGrid.SetColumn(_exportBtn, 1);
            grid.Children.Add(_exportBtn);

            _importBtn = CreateFooterButton("Import from Excel", 155, ThemeAccentGreen);
            _importBtn.Margin = new Thickness(10, 0, 10, 0);
            _importBtn.ToolTip = "Import data from Excel file into a schedule";
            _importBtn.Click += (s, e) => DoImport();
            WpfGrid.SetColumn(_importBtn, 2);
            grid.Children.Add(_importBtn);

            var cancelBtn = CreateFooterButton("Cancel", 85, WpfColor.FromRgb(127, 140, 141));
            cancelBtn.ToolTip = "Close without action";
            cancelBtn.Click += (s, e) => { DialogResult = false; Close(); };
            WpfGrid.SetColumn(cancelBtn, 3);
            grid.Children.Add(cancelBtn);

            footerBorder.Child = grid;
            return footerBorder;
        }

        private Button CreateFooterButton(string text, double width, WpfColor bgColor)
        {
            var btn = new Button
            {
                Width = width,
                Height = 38,
                FontSize = 13,
                FontWeight = FontWeights.SemiBold,
                Content = text,
                Background = new SolidColorBrush(bgColor),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btn.MouseEnter += (s, e) => btn.Opacity = 0.88;
            btn.MouseLeave += (s, e) => btn.Opacity = 1.0;
            return btn;
        }

        // ==================== HELPER ====================
        private string GetVersionFromAddin()
        {
            try
            {
                string revitVersion = _doc.Application.VersionNumber; // e.g. "2024"
                string addinFolder = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData);

                // Search for the .addin file
                string[] possibleNames = { "ScheduleLink.addin", "IB-BIM-ScheduleLink.addin" };
                foreach (var name in possibleNames)
                {
                    string addinPath = Path.Combine(addinFolder, "Autodesk", "Revit", "Addins", revitVersion, name);
                    if (File.Exists(addinPath))
                    {
                        var doc = System.Xml.Linq.XDocument.Load(addinPath);
                        var versionElement = doc.Descendants("Version").FirstOrDefault();
                        if (versionElement != null && !string.IsNullOrWhiteSpace(versionElement.Value))
                            return versionElement.Value.Trim();
                    }
                }

                // Fallback: assembly version
                return Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";
            }
            catch
            {
                return Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "1.0.0";
            }
        }

        private void OpenUrl(string url)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch
            {
                MessageBox.Show("Could not open the link.", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ==================== DATA ====================
        private void PopulateScheduleList(IEnumerable<ViewSchedule> schedules)
        {
            _scheduleList.SelectionChanged -= ScheduleList_SelectionChanged;
            _scheduleList.Items.Clear();
            int count = 0;
            foreach (var vs in schedules)
            {
                var item = new ListBoxItem
                {
                    Content = vs.Name,
                    Tag = vs,
                    Padding = new Thickness(8, 5, 8, 5),
                    FontSize = 12
                };
                _scheduleList.Items.Add(item);
                count++;
            }
            _countText.Text = $"{count} schedules found";
            _exportBtn.IsEnabled = false;
            _scheduleList.SelectedIndex = -1;
            _scheduleList.SelectionChanged += ScheduleList_SelectionChanged;

            _paramPanel.Children.Clear();
            _paramPanel.Children.Add(new TextBlock
            {
                Text = "Select a schedule to see parameters",
                Foreground = new SolidColorBrush(ThemeTextMuted),
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
                        CornerRadius = new CornerRadius(4),
                        Margin = new Thickness(2, 2, 2, 2),
                        Padding = new Thickness(10, 6, 10, 6)
                    };

                    var rowGrid = new WpfGrid();
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                    var nameText = new TextBlock
                    {
                        Text = name,
                        FontSize = 12,
                        FontWeight = FontWeights.Medium,
                        VerticalAlignment = VerticalAlignment.Center,
                        Foreground = new SolidColorBrush(ThemePrimaryMedium)
                    };
                    WpfGrid.SetColumn(nameText, 0);
                    rowGrid.Children.Add(nameText);

                    // Type badge
                    var typeBadge = new Border
                    {
                        Background = new SolidColorBrush(WpfColor.FromArgb(40, 0, 0, 0)),
                        CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(6, 2, 6, 2),
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(8, 0, 0, 0)
                    };
                    typeBadge.Child = new TextBlock
                    {
                        Text = typeLabel,
                        FontSize = 10,
                        Foreground = new SolidColorBrush(WpfColor.FromRgb(80, 80, 80)),
                        FontWeight = FontWeights.Medium
                    };
                    WpfGrid.SetColumn(typeBadge, 1);
                    rowGrid.Children.Add(typeBadge);

                    paramRow.Child = rowGrid;
                    _paramPanel.Children.Add(paramRow);
                }

                int total = instanceCount + typeCount + readOnlyCount;
                _paramTitle.Text = "Parameters — " + schedule.Name;
                _paramCount.Text = $"{total} parameters: {instanceCount} Instance, {typeCount} Type, {readOnlyCount} Read-only";
            }
            catch (Exception ex)
            {
                _paramPanel.Children.Add(new TextBlock
                {
                    Text = "Error reading parameters:\n" + ex.Message,
                    Foreground = new SolidColorBrush(ThemeAccentRed),
                    FontSize = 11,
                    Margin = new Thickness(8),
                    TextWrapping = TextWrapping.Wrap
                });
            }
        }

        // ==================== ACTIONS ====================
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

        // ==================== STYLED MESSAGE DIALOGS ====================
        /// <summary>
        /// Shows a styled warning dialog (replaces standard MessageBox for warnings)
        /// </summary>
        public static bool ShowWarning(string title, string message, string detail = null)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 440,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            // Header bar
            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentOrange),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "⚠",
                FontSize = 18,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            // Message
            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 4),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            // Detail (optional)
            if (!string.IsNullOrEmpty(detail))
            {
                mainStack.Children.Add(new TextBlock
                {
                    Text = detail,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeTextMuted),
                    Margin = new Thickness(20, 2, 20, 4),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyles.Italic
                });
            }

            // Buttons
            bool result = false;
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 14, 0, 16)
            };

            var continueBtn = new Button
            {
                Content = "Continue",
                Width = 100,
                Height = 34,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(ThemeAccentBlue),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 0, 10, 0)
            };
            continueBtn.Click += (s, e) => { result = true; dialog.Close(); };
            continueBtn.MouseEnter += (s, e) => continueBtn.Opacity = 0.88;
            continueBtn.MouseLeave += (s, e) => continueBtn.Opacity = 1.0;
            btnPanel.Children.Add(continueBtn);

            var cancelBtn = new Button
            {
                Content = "Cancel",
                Width = 100,
                Height = 34,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(WpfColor.FromRgb(189, 195, 199)),
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            cancelBtn.Click += (s, e) => { result = false; dialog.Close(); };
            cancelBtn.MouseEnter += (s, e) => cancelBtn.Opacity = 0.88;
            cancelBtn.MouseLeave += (s, e) => cancelBtn.Opacity = 1.0;
            btnPanel.Children.Add(cancelBtn);

            mainStack.Children.Add(btnPanel);
            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
            return result;
        }

        /// <summary>
        /// Shows a styled success dialog
        /// </summary>
        public static void ShowSuccess(string title, string message, string detail = null)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 440,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentGreen),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "✓",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 4),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            if (!string.IsNullOrEmpty(detail))
            {
                mainStack.Children.Add(new TextBlock
                {
                    Text = detail,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeTextMuted),
                    Margin = new Thickness(20, 2, 20, 4),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyles.Italic
                });
            }

            var okBtn = new Button
            {
                Content = "OK",
                Width = 100,
                Height = 34,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(ThemeAccentGreen),
                Foreground = Brushes.White,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 14, 0, 16),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            okBtn.Click += (s, e) => dialog.Close();
            okBtn.MouseEnter += (s, e) => okBtn.Opacity = 0.88;
            okBtn.MouseLeave += (s, e) => okBtn.Opacity = 1.0;
            mainStack.Children.Add(okBtn);

            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
        }

        /// <summary>
        /// Shows a styled error dialog
        /// </summary>
        public static void ShowError(string title, string message, string detail = null)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 460,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentRed),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "✕",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 4),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            if (!string.IsNullOrEmpty(detail))
            {
                mainStack.Children.Add(new TextBlock
                {
                    Text = detail,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeTextMuted),
                    Margin = new Thickness(20, 2, 20, 4),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyles.Italic
                });
            }

            var okBtn = new Button
            {
                Content = "OK",
                Width = 100,
                Height = 34,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(WpfColor.FromRgb(189, 195, 199)),
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Margin = new Thickness(0, 14, 0, 16),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            okBtn.Click += (s, e) => dialog.Close();
            okBtn.MouseEnter += (s, e) => okBtn.Opacity = 0.88;
            okBtn.MouseLeave += (s, e) => okBtn.Opacity = 1.0;
            mainStack.Children.Add(okBtn);

            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
        }

        /// <summary>
        /// Shows a styled info dialog (replaces TaskDialog.Show for simple messages)
        /// </summary>
        public static void ShowInfo(string title, string message, string detail = null)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 440,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentBlue),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "i",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 4),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            if (!string.IsNullOrEmpty(detail))
            {
                mainStack.Children.Add(new TextBlock
                {
                    Text = detail,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeTextMuted),
                    Margin = new Thickness(20, 2, 20, 4),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyles.Italic
                });
            }

            var okBtn = CreateDialogButton("OK", ThemeAccentBlue, Brushes.White);
            okBtn.HorizontalAlignment = HorizontalAlignment.Center;
            okBtn.Margin = new Thickness(0, 14, 0, 16);
            okBtn.Click += (s, e) => dialog.Close();
            mainStack.Children.Add(okBtn);

            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
        }

        /// <summary>
        /// Shows a styled confirm dialog with Yes/No (replaces TaskDialog with CommonButtons)
        /// Returns true if Yes clicked
        /// </summary>
        public static bool ShowConfirm(string title, string message, string detail = null,
            string yesText = "Yes", string noText = "No")
        {
            var dialog = new Window
            {
                Title = title,
                Width = 460,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 500,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentBlue),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "?",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 4),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            if (!string.IsNullOrEmpty(detail))
            {
                mainStack.Children.Add(new TextBlock
                {
                    Text = detail,
                    FontSize = 11,
                    Foreground = new SolidColorBrush(ThemeTextMuted),
                    Margin = new Thickness(20, 2, 20, 4),
                    TextWrapping = TextWrapping.Wrap,
                    FontStyle = FontStyles.Italic
                });
            }

            bool result = false;
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 14, 0, 16)
            };

            var yesBtn = CreateDialogButton(yesText, ThemeAccentBlue, Brushes.White);
            yesBtn.Margin = new Thickness(0, 0, 10, 0);
            yesBtn.Click += (s, e) => { result = true; dialog.Close(); };
            btnPanel.Children.Add(yesBtn);

            var noBtn = CreateDialogButton(noText, WpfColor.FromRgb(189, 195, 199), new SolidColorBrush(ThemePrimaryMedium));
            noBtn.Click += (s, e) => { result = false; dialog.Close(); };
            btnPanel.Children.Add(noBtn);

            mainStack.Children.Add(btnPanel);
            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
            return result;
        }

        /// <summary>
        /// Shows a styled export-complete dialog with option to open file
        /// Returns true if user wants to open the file
        /// </summary>
        public static bool ShowExportComplete(string scheduleName, int rows, int columns, string filePath)
        {
            var dialog = new Window
            {
                Title = "Export Complete",
                Width = 480,
                Height = 260,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            // Green header
            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentGreen),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "\u2713",
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = "Export Completed Successfully",
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            // Details
            var detailPanel = new StackPanel { Margin = new Thickness(20, 14, 20, 0) };

            detailPanel.Children.Add(CreateDetailRow("Schedule:", scheduleName));
            detailPanel.Children.Add(CreateDetailRow("Rows:", rows.ToString()));
            detailPanel.Children.Add(CreateDetailRow("Columns:", columns.ToString()));
            detailPanel.Children.Add(new TextBlock
            {
                Text = filePath,
                FontSize = 11,
                Foreground = new SolidColorBrush(ThemeTextMuted),
                Margin = new Thickness(0, 8, 0, 0),
                TextWrapping = TextWrapping.Wrap,
                TextTrimming = TextTrimming.CharacterEllipsis
            });
            mainStack.Children.Add(detailPanel);

            // Buttons
            bool openFile = false;
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 14, 0, 16)
            };

            var openBtn = CreateDialogButton("Open File", ThemeAccentGreen, Brushes.White);
            openBtn.Margin = new Thickness(0, 0, 10, 0);
            openBtn.Click += (s, e) => { openFile = true; dialog.Close(); };
            btnPanel.Children.Add(openBtn);

            var closeBtn = CreateDialogButton("Close", WpfColor.FromRgb(189, 195, 199), new SolidColorBrush(ThemePrimaryMedium));
            closeBtn.Click += (s, e) => { openFile = false; dialog.Close(); };
            btnPanel.Children.Add(closeBtn);

            mainStack.Children.Add(btnPanel);
            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
            return openFile;
        }

        /// <summary>
        /// Shows a styled import result dialog with summary
        /// </summary>
        public static void ShowImportResult(string title, string summaryText, bool hasUpdates, int failedCount = 0)
        {
            // Determine color scheme: green=all good, yellow-orange=partial, orange=no changes
            WpfColor headerColor;
            string headerIcon;
            if (hasUpdates && failedCount == 0)
            {
                headerColor = ThemeAccentGreen;
                headerIcon = "\u2713";
            }
            else if (hasUpdates && failedCount > 0)
            {
                headerColor = WpfColor.FromRgb(230, 126, 34); // orange - partial success
                headerIcon = "\u2713";
            }
            else
            {
                headerColor = ThemeAccentOrange;
                headerIcon = "!";
            }

            // Calculate width based on longest error line
            double baseWidth = 520;
            if (failedCount > 0)
            {
                foreach (string line in summaryText.Split('\n'))
                {
                    double estimated = line.Length * 7.0 + 60; // rough char width + padding
                    if (estimated > baseWidth) baseWidth = estimated;
                }
            }
            if (baseWidth > 800) baseWidth = 800; // cap at 800

            var dialog = new Window
            {
                Title = "ScheduleLink - " + title,
                Width = baseWidth,
                Height = 420,
                MinWidth = 400,
                MinHeight = 250,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.CanResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            // Use DockPanel so OK button stays at bottom when resizing
            var dockPanel = new DockPanel { LastChildFill = true };

            // === TOP: Header bar (docked top) ===
            var headerBar = new Border
            {
                Background = new SolidColorBrush(headerColor),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = headerIcon,
                FontSize = 18,
                FontWeight = FontWeights.Bold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            DockPanel.SetDock(headerBar, Dock.Top);
            dockPanel.Children.Add(headerBar);

            // === BOTTOM: OK + Copy buttons (docked bottom) ===
            var btnBorder = new Border
            {
                Padding = new Thickness(0, 10, 0, 12),
                Background = new SolidColorBrush(WpfColor.FromRgb(245, 246, 248)),
                BorderBrush = new SolidColorBrush(ThemeBorderLight),
                BorderThickness = new Thickness(0, 1, 0, 0)
            };
            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            var copyBtn = CreateDialogButton("Copy", WpfColor.FromRgb(149, 165, 179), Brushes.White);
            copyBtn.ToolTip = "Copy report to clipboard";
            copyBtn.Margin = new Thickness(0, 0, 10, 0);
            copyBtn.Click += (s, e) =>
            {
                try
                {
                    System.Windows.Clipboard.SetText(summaryText);
                    copyBtn.Content = "Copied!";
                }
                catch { }
            };
            btnPanel.Children.Add(copyBtn);

            var okBtn = CreateDialogButton("OK", headerColor, Brushes.White);
            okBtn.Click += (s, e) => dialog.Close();
            btnPanel.Children.Add(okBtn);

            btnBorder.Child = btnPanel;
            DockPanel.SetDock(btnBorder, Dock.Bottom);
            dockPanel.Children.Add(btnBorder);

            // === CENTER: Selectable text in scrollable area (fills remaining space) ===
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Padding = new Thickness(20, 14, 20, 10)
            };
            scrollViewer.Content = new TextBox
            {
                Text = summaryText,
                FontSize = 12,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                TextWrapping = TextWrapping.Wrap,
                IsReadOnly = true,
                BorderThickness = new Thickness(0),
                Background = Brushes.Transparent,
                Cursor = Cursors.IBeam
            };
            dockPanel.Children.Add(scrollViewer);

            dialog.Content = dockPanel;
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
        }

        /// <summary>
        /// Shows a styled rating dialog with 3 options.
        /// Returns: 1=Yes/Rate, 2=Maybe later, 3=No thanks
        /// </summary>
        public static int ShowRatingDialog(string title, string message)
        {
            var dialog = new Window
            {
                Title = title,
                Width = 460,
                SizeToContent = SizeToContent.Height,
                MaxHeight = 400,
                WindowStartupLocation = WindowStartupLocation.CenterScreen,
                ResizeMode = ResizeMode.NoResize,
                Background = new SolidColorBrush(WpfColor.FromRgb(250, 250, 252)),
                WindowStyle = WindowStyle.ToolWindow,
                ShowInTaskbar = false
            };

            var mainStack = new StackPanel();

            // Purple header
            var headerBar = new Border
            {
                Background = new SolidColorBrush(ThemeAccentPurple),
                Padding = new Thickness(16, 10, 16, 10)
            };
            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };
            headerStack.Children.Add(new TextBlock
            {
                Text = "\u2605",
                FontSize = 18,
                Foreground = new SolidColorBrush(WpfColor.FromRgb(241, 196, 15)),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            });
            headerStack.Children.Add(new TextBlock
            {
                Text = title,
                FontSize = 14,
                FontWeight = FontWeights.SemiBold,
                Foreground = Brushes.White,
                VerticalAlignment = VerticalAlignment.Center
            });
            headerBar.Child = headerStack;
            mainStack.Children.Add(headerBar);

            mainStack.Children.Add(new TextBlock
            {
                Text = message,
                FontSize = 13,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Margin = new Thickness(20, 16, 20, 8),
                TextWrapping = TextWrapping.Wrap,
                LineHeight = 20
            });

            int result = 0;
            var btnPanel = new StackPanel { Margin = new Thickness(20, 6, 20, 16) };

            // Option 1: Yes - Rate
            var rateBtn = CreateDialogButton("Yes, take me to the App Store", ThemeAccentGreen, Brushes.White);
            rateBtn.Width = double.NaN; // auto width
            rateBtn.HorizontalAlignment = HorizontalAlignment.Stretch;
            rateBtn.Margin = new Thickness(0, 0, 0, 6);
            rateBtn.Click += (s, e) => { result = 1; dialog.Close(); };
            btnPanel.Children.Add(rateBtn);

            // Option 2: Maybe later
            var laterBtn = CreateDialogButton("Maybe later", ThemeAccentBlue, Brushes.White);
            laterBtn.Width = double.NaN;
            laterBtn.HorizontalAlignment = HorizontalAlignment.Stretch;
            laterBtn.Margin = new Thickness(0, 0, 0, 6);
            laterBtn.Click += (s, e) => { result = 2; dialog.Close(); };
            btnPanel.Children.Add(laterBtn);

            // Option 3: No thanks
            var noBtn = CreateDialogButton("No thanks", WpfColor.FromRgb(189, 195, 199), new SolidColorBrush(ThemePrimaryMedium));
            noBtn.Width = double.NaN;
            noBtn.HorizontalAlignment = HorizontalAlignment.Stretch;
            noBtn.Click += (s, e) => { result = 3; dialog.Close(); };
            btnPanel.Children.Add(noBtn);

            mainStack.Children.Add(btnPanel);
            dialog.Content = mainStack;
            try { dialog.Owner = System.Windows.Application.Current?.MainWindow; } catch { }
            SetRevitAsOwner(dialog);
            dialog.ShowDialog();
            return result;
        }

        // ==================== DIALOG HELPERS ====================
        /// <summary>
        /// Sets the Revit main window as owner of the dialog to prevent it from hiding behind Revit
        /// </summary>
        private static void SetRevitAsOwner(Window dialog)
        {
            try
            {
                var revitProcess = System.Diagnostics.Process.GetCurrentProcess();
                var revitHandle = revitProcess.MainWindowHandle;
                if (revitHandle != IntPtr.Zero)
                {
                    var helper = new System.Windows.Interop.WindowInteropHelper(dialog);
                    helper.Owner = revitHandle;
                }
            }
            catch { }
        }
        private static Button CreateDialogButton(string text, WpfColor bgColor, Brush fgBrush)
        {
            var btn = new Button
            {
                Content = text,
                Width = 110,
                Height = 34,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Background = new SolidColorBrush(bgColor),
                Foreground = fgBrush,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand
            };
            btn.MouseEnter += (s, e) => btn.Opacity = 0.88;
            btn.MouseLeave += (s, e) => btn.Opacity = 1.0;
            return btn;
        }

        private static StackPanel CreateDetailRow(string label, string value)
        {
            var row = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                Margin = new Thickness(0, 2, 0, 2)
            };
            row.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 12,
                FontWeight = FontWeights.SemiBold,
                Foreground = new SolidColorBrush(ThemePrimaryMedium),
                Width = 80
            });
            row.Children.Add(new TextBlock
            {
                Text = value,
                FontSize = 12,
                Foreground = new SolidColorBrush(ThemePrimaryLight)
            });
            return row;
        }
    }
}