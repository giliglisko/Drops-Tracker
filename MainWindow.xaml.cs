using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Drops_Tracker.Models;
using Drops_Tracker.Services;

namespace Drops_Tracker
{
    public partial class MainWindow : Window
    {
        private readonly DataService _dataService;
        private readonly string _assetsPath;
        private DateTime _startWeek;
        private const int WEEKS_TO_DISPLAY = 8;
        private int _overviewYear;
        private bool _isDarkMode;
        private const string LootItemDragFormat = "LootItem";
        private const string CharacterReorderDragFormat = "CharacterReorder";

        // Drag & drop state
        private Window? _dragWindow;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x00000020;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private void MakeWindowClickThrough(Window window)
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(window).Handle;
            var extendedStyle = GetWindowLong(hwnd, GWL_EXSTYLE);
            SetWindowLong(hwnd, GWL_EXSTYLE, extendedStyle | WS_EX_TRANSPARENT);
        }

        public MainWindow()
        {
            InitializeComponent();
            _assetsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets");
            _dataService = new DataService(_assetsPath);
            _startWeek = DataService.GetWeekStart(DateTime.Now).AddDays(-7 * (WEEKS_TO_DISPLAY - 1));
            _overviewYear = DateTime.Now.Year;

            // Set window icon from assets
            var iconPath = Path.Combine(_assetsPath, "Genesis Badge.png");
            if (File.Exists(iconPath))
            {
                Icon = new BitmapImage(new Uri(iconPath));
            }

            // Set up grid-level drop handling
            CalendarGrid.PreviewDragOver += CalendarGrid_PreviewDragOver;
            CalendarGrid.PreviewDrop += CalendarGrid_PreviewDrop;

            // Load dark mode setting
            _isDarkMode = _dataService.GetDarkModeSetting();
            ApplyTheme(_isDarkMode);
            UpdateDarkModeToggleVisual();

            RefreshUI();
        }

        private void CalendarGrid_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(CharacterReorderDragFormat))
            {
                e.Effects = DragDropEffects.Move;
                return;
            }

            if (e.Data.GetDataPresent(LootItemDragFormat))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
                e.Handled = false;
            }
        }

        private void CalendarGrid_PreviewDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(LootItemDragFormat))
                return;

            var item = (LootItem)e.Data.GetData(LootItemDragFormat)!;
            var position = e.GetPosition(CalendarGrid);

            // Find which cell was dropped on using hit testing
            var hitResult = VisualTreeHelper.HitTest(CalendarGrid, position);
            if (hitResult?.VisualHit != null)
            {
                // Walk up the visual tree to find a Border with our Tag
                DependencyObject current = hitResult.VisualHit;
                while (current != null)
                {
                    if (current is Border border && border.Tag is Tuple<string, string> data)
                    {
                        var characterId = data.Item1;
                        var weekKey = data.Item2;
                        _dataService.AddDropToCell(characterId, weekKey, item.Id);
                        RefreshCalendarGrid();
                        e.Handled = true;
                        return;
                    }
                    current = VisualTreeHelper.GetParent(current);
                }
            }
        }

        private void RefreshUI()
        {
            UpdateWeekRangeDisplay();
            RefreshCalendarGrid();
            RefreshInventoryPanel();
        }

        private void UpdateWeekRangeDisplay()
        {
            // Week range display removed - navigation buttons only
        }

        private void PrevWeek_Click(object sender, RoutedEventArgs e)
        {
            _startWeek = _startWeek.AddDays(-7);
            RefreshUI();
        }

        private void NextWeek_Click(object sender, RoutedEventArgs e)
        {
            _startWeek = _startWeek.AddDays(7);
            RefreshUI();
        }

        private void GoToToday_Click(object sender, RoutedEventArgs e)
        {
            _startWeek = DataService.GetWeekStart(DateTime.Now).AddDays(-7 * (WEEKS_TO_DISPLAY - 1));
            RefreshUI();
        }

        private int _previousTabIndex = 0;

        private void MainTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MainTabControl.SelectedIndex == 1)
            {
                RefreshOverview();
                _previousTabIndex = 1;
            }
            else if (MainTabControl.SelectedIndex == 2)
            {
                // Settings tab
                _previousTabIndex = 2;
            }
            else if (MainTabControl.SelectedIndex == 3)
            {
                // "New Character" tab clicked - open dialog and go back
                MainTabControl.SelectedIndex = _previousTabIndex;
                AddCharacter_Click(sender, new RoutedEventArgs());
            }
            else
            {
                _previousTabIndex = MainTabControl.SelectedIndex;
            }
        }

        private void PrevYear_Click(object sender, RoutedEventArgs e)
        {
            _overviewYear--;
            RefreshOverview();
        }

        private void NextYear_Click(object sender, RoutedEventArgs e)
        {
            _overviewYear++;
            RefreshOverview();
        }

        private void RefreshOverview()
        {
            YearDisplay.Text = _overviewYear.ToString();
            RefreshYearlyOverview();
            RefreshMonthlyOverview();
        }

        private void RefreshYearlyOverview()
        {
            YearlyOverviewGrid.Children.Clear();
            YearlyOverviewGrid.RowDefinitions.Clear();
            YearlyOverviewGrid.ColumnDefinitions.Clear();

            var characters = _dataService.Characters;
            var items = _dataService.GetItemsFromAssets();

            if (characters.Count == 0)
            {
                YearlyOverviewGrid.Children.Add(new TextBlock
                {
                    Text = "No characters to display",
                    Foreground = (SolidColorBrush)FindResource("TextMutedBrush"),
                    FontStyle = FontStyles.Italic
                });
                return;
            }

            // Create columns: Character name + one per item
            YearlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            foreach (var item in items)
            {
                YearlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) });
            }
            YearlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(80) }); // Total column

            // Create rows: Header + one per character
            YearlyOverviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            foreach (var _ in characters)
            {
                YearlyOverviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }

            // Header row
            var charHeader = CreateOverviewHeaderCell("Character");
            Grid.SetRow(charHeader, 0);
            Grid.SetColumn(charHeader, 0);
            YearlyOverviewGrid.Children.Add(charHeader);

            for (int i = 0; i < items.Count; i++)
            {
                var itemHeader = CreateOverviewItemHeader(items[i]);
                Grid.SetRow(itemHeader, 0);
                Grid.SetColumn(itemHeader, i + 1);
                YearlyOverviewGrid.Children.Add(itemHeader);
            }

            var totalHeader = CreateOverviewHeaderCell("Total");
            Grid.SetRow(totalHeader, 0);
            Grid.SetColumn(totalHeader, items.Count + 1);
            YearlyOverviewGrid.Children.Add(totalHeader);

            // Character rows
            for (int row = 0; row < characters.Count; row++)
            {
                var character = characters[row];
                var yearDrops = GetDropsForYear(character.Id, _overviewYear);

                var nameCell = CreateOverviewCell(character.Name, true);
                Grid.SetRow(nameCell, row + 1);
                Grid.SetColumn(nameCell, 0);
                YearlyOverviewGrid.Children.Add(nameCell);

                int totalDrops = 0;
                for (int col = 0; col < items.Count; col++)
                {
                    var item = items[col];
                    var count = yearDrops.Where(d => d.ItemId == item.Id).Sum(d => d.Quantity);

                    // Exclude Grindstone items from total count
                    if (!item.Name.Contains("Grindstone", StringComparison.OrdinalIgnoreCase))
                    {
                        totalDrops += count;
                    }

                    var countCell = CreateOverviewCell(count > 0 ? count.ToString() : "-", false);
                    Grid.SetRow(countCell, row + 1);
                    Grid.SetColumn(countCell, col + 1);
                    YearlyOverviewGrid.Children.Add(countCell);
                }

                var totalCell = CreateOverviewCell(totalDrops.ToString(), false, true);
                Grid.SetRow(totalCell, row + 1);
                Grid.SetColumn(totalCell, items.Count + 1);
                YearlyOverviewGrid.Children.Add(totalCell);
            }
        }

        private void RefreshMonthlyOverview()
        {
            MonthlyOverviewGrid.Children.Clear();
            MonthlyOverviewGrid.RowDefinitions.Clear();
            MonthlyOverviewGrid.ColumnDefinitions.Clear();

            var characters = _dataService.Characters;

            if (characters.Count == 0)
            {
                MonthlyOverviewGrid.Children.Add(new TextBlock
                {
                    Text = "No characters to display",
                    Foreground = (SolidColorBrush)FindResource("TextMutedBrush"),
                    FontStyle = FontStyles.Italic
                });
                return;
            }

            // Create columns: Character name + 12 months + total
            MonthlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(150) });
            for (int i = 0; i < 12; i++)
            {
                MonthlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
            }
            MonthlyOverviewGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) }); // Total

            // Create rows: Header + one per character
            MonthlyOverviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            foreach (var _ in characters)
            {
                MonthlyOverviewGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            }

            // Header row
            var charHeader = CreateOverviewHeaderCell("Character");
            Grid.SetRow(charHeader, 0);
            Grid.SetColumn(charHeader, 0);
            MonthlyOverviewGrid.Children.Add(charHeader);

            string[] monthNames = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            for (int i = 0; i < 12; i++)
            {
                var monthHeader = CreateOverviewHeaderCell(monthNames[i]);
                Grid.SetRow(monthHeader, 0);
                Grid.SetColumn(monthHeader, i + 1);
                MonthlyOverviewGrid.Children.Add(monthHeader);
            }

            var totalHeader = CreateOverviewHeaderCell("Total");
            Grid.SetRow(totalHeader, 0);
            Grid.SetColumn(totalHeader, 13);
            MonthlyOverviewGrid.Children.Add(totalHeader);

            // Character rows
            var items = _dataService.GetItemsFromAssets();
            var grindstoneItemIds = items
                .Where(i => i.Name.Contains("Grindstone", StringComparison.OrdinalIgnoreCase))
                .Select(i => i.Id)
                .ToHashSet();

            for (int row = 0; row < characters.Count; row++)
            {
                var character = characters[row];

                var nameCell = CreateOverviewCell(character.Name, true);
                Grid.SetRow(nameCell, row + 1);
                Grid.SetColumn(nameCell, 0);
                MonthlyOverviewGrid.Children.Add(nameCell);

                int totalDrops = 0;
                for (int month = 1; month <= 12; month++)
                {
                    var monthDrops = GetDropsForMonth(character.Id, _overviewYear, month);
                    var count = monthDrops.Sum(d => d.Quantity);
                    // Exclude Grindstone items from total count
                    var countExcludingGrindstones = monthDrops
                        .Where(d => !grindstoneItemIds.Contains(d.ItemId))
                        .Sum(d => d.Quantity);
                    totalDrops += countExcludingGrindstones;

                    var countCell = CreateOverviewCell(count > 0 ? count.ToString() : "-", false);
                    Grid.SetRow(countCell, row + 1);
                    Grid.SetColumn(countCell, month);
                    MonthlyOverviewGrid.Children.Add(countCell);
                }

                var totalCell = CreateOverviewCell(totalDrops.ToString(), false, true);
                Grid.SetRow(totalCell, row + 1);
                Grid.SetColumn(totalCell, 13);
                MonthlyOverviewGrid.Children.Add(totalCell);
            }
        }

        private List<WeeklyDrop> GetDropsForYear(string characterId, int year)
        {
            return _dataService.Drops.Where(d =>
                d.CharacterId == characterId &&
                d.IncludeInSummary &&
                DateTime.TryParse(d.WeekKey, out var weekDate) &&
                weekDate.Year == year).ToList();
        }

        private List<WeeklyDrop> GetDropsForMonth(string characterId, int year, int month)
        {
            return _dataService.Drops.Where(d =>
                d.CharacterId == characterId &&
                d.IncludeInSummary &&
                DateTime.TryParse(d.WeekKey, out var weekDate) &&
                weekDate.Year == year &&
                weekDate.Month == month).ToList();
        }

        private Border CreateOverviewHeaderCell(string text)
        {
            return new Border
            {
                Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                Padding = new Thickness(8, 6, 8, 6),
                Child = new TextBlock
                {
                    Text = text,
                    FontWeight = FontWeights.SemiBold,
                    FontSize = 12,
                    Foreground = Brushes.White,
                    TextAlignment = TextAlignment.Center
                }
            };
        }

        private Border CreateOverviewItemHeader(LootItem item)
        {
            var stack = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };

            var imagePath = GetImagePath(item.ImageFileName);
            if (!string.IsNullOrEmpty(imagePath))
            {
                var image = new Image
                {
                    Source = new BitmapImage(new Uri(imagePath)),
                    Width = 24,
                    Height = 24,
                    ToolTip = item.Name
                };
                RenderOptions.SetBitmapScalingMode(image, BitmapScalingMode.HighQuality);
                stack.Children.Add(image);
            }
            else
            {
                stack.Children.Add(new TextBlock
                {
                    Text = item.Name.Length > 3 ? item.Name.Substring(0, 3) : item.Name,
                    FontSize = 10,
                    Foreground = Brushes.White,
                    ToolTip = item.Name
                });
            }

            return new Border
            {
                Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                Padding = new Thickness(4),
                Child = stack
            };
        }

        private Border CreateOverviewCell(string text, bool isName, bool isTotal = false)
        {
            return new Border
            {
                Background = isTotal
                    ? (SolidColorBrush)FindResource("CharacterCellBrush")
                    : (SolidColorBrush)FindResource("CardBrush"),
                BorderBrush = (SolidColorBrush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(8, 6, 8, 6),
                Child = new TextBlock
                {
                    Text = text,
                    FontWeight = isName || isTotal ? FontWeights.SemiBold : FontWeights.Normal,
                    FontSize = 12,
                    Foreground = (SolidColorBrush)FindResource("TextBrush"),
                    TextAlignment = isName ? TextAlignment.Left : TextAlignment.Center,
                    TextTrimming = TextTrimming.CharacterEllipsis
                }
            };
        }

        private void RefreshCalendarGrid()
        {
            CalendarGrid.Children.Clear();
            CalendarGrid.RowDefinitions.Clear();
            CalendarGrid.ColumnDefinitions.Clear();

            var characters = _dataService.Characters;

            if (characters.Count == 0)
            {
                NoDataPanel.Visibility = Visibility.Visible;
                return;
            }

            NoDataPanel.Visibility = Visibility.Collapsed;

            // Create columns
            CalendarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
            for (int i = 0; i < WEEKS_TO_DISPLAY; i++)
            {
                CalendarGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            }

            // Create rows
            CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            foreach (var _ in characters)
            {
                CalendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(75) });
            }

            // Header - Character column
            var charHeader = CreateHeaderCell("Character");
            Grid.SetRow(charHeader, 0);
            Grid.SetColumn(charHeader, 0);
            CalendarGrid.Children.Add(charHeader);

            // Header - Week columns
            for (int i = 0; i < WEEKS_TO_DISPLAY; i++)
            {
                var weekDate = _startWeek.AddDays(7 * i);
                var weekHeader = CreateWeekHeaderCell(weekDate);
                Grid.SetRow(weekHeader, 0);
                Grid.SetColumn(weekHeader, i + 1);
                CalendarGrid.Children.Add(weekHeader);
            }

            // Character rows
            var items = _dataService.GetItemsFromAssets();
            for (int row = 0; row < characters.Count; row++)
            {
                var character = characters[row];

                var charCell = CreateCharacterCell(character);
                Grid.SetRow(charCell, row + 1);
                Grid.SetColumn(charCell, 0);
                CalendarGrid.Children.Add(charCell);

                for (int col = 0; col < WEEKS_TO_DISPLAY; col++)
                {
                    var weekDate = _startWeek.AddDays(7 * col);
                    var weekKey = DataService.GetWeekKey(weekDate);
                    var drops = _dataService.GetDropsForCharacterWeek(character.Id, weekKey);

                    var weekCell = CreateWeekCell(character, weekKey, drops, items);
                    Grid.SetRow(weekCell, row + 1);
                    Grid.SetColumn(weekCell, col + 1);
                    CalendarGrid.Children.Add(weekCell);
                }
            }

            // Add orange border lines for current week column
            for (int i = 0; i < WEEKS_TO_DISPLAY; i++)
            {
                var weekDate = _startWeek.AddDays(7 * i);
                var isCurrentWeek = DataService.GetWeekKey(weekDate) == DataService.GetWeekKey(DateTime.Now);
                if (isCurrentWeek)
                {
                    // Left orange line
                    var leftLine = new Border
                    {
                        Width = 3,
                        Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                        HorizontalAlignment = HorizontalAlignment.Left,
                        IsHitTestVisible = false
                    };
                    Grid.SetColumn(leftLine, i + 1);
                    Grid.SetRow(leftLine, 0);
                    Grid.SetRowSpan(leftLine, characters.Count + 1);
                    CalendarGrid.Children.Add(leftLine);

                    // Right orange line
                    var rightLine = new Border
                    {
                        Width = 3,
                        Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                        HorizontalAlignment = HorizontalAlignment.Right,
                        IsHitTestVisible = false
                    };
                    Grid.SetColumn(rightLine, i + 1);
                    Grid.SetRow(rightLine, 0);
                    Grid.SetRowSpan(rightLine, characters.Count + 1);
                    CalendarGrid.Children.Add(rightLine);

                    // Bottom orange line
                    var bottomLine = new Border
                    {
                        Height = 3,
                        Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                        VerticalAlignment = VerticalAlignment.Bottom,
                        IsHitTestVisible = false
                    };
                    Grid.SetColumn(bottomLine, i + 1);
                    Grid.SetRow(bottomLine, characters.Count);
                    CalendarGrid.Children.Add(bottomLine);
                    break;
                }
            }
        }

        private Border CreateHeaderCell(string text)
        {
            return new Border
            {
                Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                Padding = new Thickness(8, 6, 8, 6),
                Child = new TextBlock
                {
                    Text = text,
                    FontWeight = FontWeights.Bold,
                    FontSize = 11,
                    Foreground = Brushes.White
                }
            };
        }

        private Border CreateWeekHeaderCell(DateTime weekStart)
        {
            return new Border
            {
                Background = (SolidColorBrush)FindResource("PrimaryBrush"),
                Padding = new Thickness(4, 4, 4, 4),
                Child = new TextBlock
                {
                    Text = weekStart.ToString("MMM dd"),
                    FontWeight = FontWeights.SemiBold,
                    FontSize = 11,
                    Foreground = Brushes.White,
                    HorizontalAlignment = HorizontalAlignment.Center
                }
            };
        }

        private Border CreateCharacterCell(Character character)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(18) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var handleLines = new StackPanel
            {
                Orientation = Orientation.Vertical,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };
            for (int i = 0; i < 3; i++)
            {
                handleLines.Children.Add(new Border
                {
                    Width = 9,
                    Height = 1.5,
                    Background = (SolidColorBrush)FindResource("TextMutedBrush"),
                    Margin = new Thickness(0, 1.5, 0, 1.5),
                    CornerRadius = new CornerRadius(1)
                });
            }

            var dragHandle = new Border
            {
                Width = 14,
                Height = 32,
                Background = Brushes.Transparent,
                Cursor = Cursors.SizeNS,
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = "Drag to reorder character rows",
                Tag = character.Id,
                Child = handleLines
            };
            dragHandle.MouseLeftButtonDown += CharacterHandle_MouseLeftButtonDown;
            dragHandle.MouseMove += CharacterHandle_MouseMove;
            Grid.SetColumn(dragHandle, 0);
            grid.Children.Add(dragHandle);

            var infoStack = new StackPanel
            {
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(4, 0, 18, 0)
            };
            infoStack.Children.Add(new TextBlock
            {
                Text = character.Name,
                FontWeight = FontWeights.SemiBold,
                FontSize = 11,
                Foreground = (SolidColorBrush)FindResource("TextBrush"),
                TextTrimming = TextTrimming.CharacterEllipsis
            });

            if (character.Level > 0 || !string.IsNullOrEmpty(character.Class))
            {
                var details = new List<string>();
                if (character.Level > 0) details.Add($"Lv.{character.Level}");
                if (!string.IsNullOrEmpty(character.Class)) details.Add(character.Class);

                infoStack.Children.Add(new TextBlock
                {
                    Text = string.Join(" | ", details),
                    FontSize = 9,
                    Foreground = (SolidColorBrush)FindResource("TextMutedBrush"),
                    TextTrimming = TextTrimming.CharacterEllipsis
                });
            }

            Grid.SetColumn(infoStack, 1);
            grid.Children.Add(infoStack);

            var editBtn = new Button
            {
                Content = "\u270E",
                Style = (Style)FindResource("EditButton"),
                HorizontalAlignment = HorizontalAlignment.Left,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(-3, -3, 0, 0),
                ToolTip = "Edit character",
                Tag = character.Id
            };
            editBtn.Click += EditCharacter_Click;
            Grid.SetColumn(editBtn, 1);
            grid.Children.Add(editBtn);

            var deleteBtn = new Button
            {
                Content = "X",
                Style = (Style)FindResource("DeleteButton"),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, -3, -3, 0),
                ToolTip = "Delete character",
                Tag = character.Id
            };
            deleteBtn.Click += DeleteCharacter_Click;
            Grid.SetColumn(deleteBtn, 1);
            grid.Children.Add(deleteBtn);

            var characterCell = new Border
            {
                Background = (SolidColorBrush)FindResource("CharacterCellBrush"),
                BorderBrush = (SolidColorBrush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(6),
                AllowDrop = true,
                Tag = character.Id,
                Child = grid
            };
            characterCell.PreviewDragEnter += CharacterCell_DragEnter;
            characterCell.PreviewDragOver += CharacterCell_DragOver;
            characterCell.PreviewDragLeave += CharacterCell_DragLeave;
            characterCell.PreviewDrop += CharacterCell_Drop;

            return characterCell;
        }

        private void CharacterHandle_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _characterDragStartPoint = e.GetPosition(null);
            _isCharacterDragging = false;
        }

        private void CharacterHandle_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            if (sender is not FrameworkElement element || element.Tag is not string characterId || _isCharacterDragging)
                return;

            Point mousePos = e.GetPosition(null);
            Vector diff = _characterDragStartPoint - mousePos;

            if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                _isCharacterDragging = true;
                var dataObject = new DataObject(CharacterReorderDragFormat, characterId);
                DragDrop.DoDragDrop(element, dataObject, DragDropEffects.Move);
                _isCharacterDragging = false;
            }
        }

        private void CharacterCell_DragEnter(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(CharacterReorderDragFormat))
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            if (sender is Border border)
            {
                border.Background = new SolidColorBrush(Color.FromArgb(0x55, 0xFF, 0x8C, 0x00));
            }

            e.Effects = DragDropEffects.Move;
            e.Handled = true;
        }

        private void CharacterCell_DragOver(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(CharacterReorderDragFormat))
            {
                e.Effects = DragDropEffects.None;
                return;
            }

            if (sender is Border border &&
                border.Tag is string targetCharacterId &&
                e.Data.GetData(CharacterReorderDragFormat) is string sourceCharacterId &&
                sourceCharacterId == targetCharacterId)
            {
                e.Effects = DragDropEffects.None;
            }
            else
            {
                e.Effects = DragDropEffects.Move;
            }

            e.Handled = true;
        }

        private void CharacterCell_DragLeave(object sender, DragEventArgs e)
        {
            if (sender is Border border)
            {
                border.Background = (SolidColorBrush)FindResource("CharacterCellBrush");
            }
        }

        private void CharacterCell_Drop(object sender, DragEventArgs e)
        {
            if (sender is Border border)
            {
                border.Background = (SolidColorBrush)FindResource("CharacterCellBrush");
            }

            if (!e.Data.GetDataPresent(CharacterReorderDragFormat))
                return;

            if (sender is not Border targetBorder ||
                targetBorder.Tag is not string targetCharacterId ||
                e.Data.GetData(CharacterReorderDragFormat) is not string sourceCharacterId)
            {
                return;
            }

            if (_dataService.SwapCharacterOrder(sourceCharacterId, targetCharacterId))
            {
                RefreshCalendarGrid();
            }

            e.Handled = true;
        }

        private Border CreateWeekCell(Character character, string weekKey, List<WeeklyDrop> drops, List<LootItem> items)
        {
            var wrapPanel = new WrapPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };

            foreach (var drop in drops)
            {
                var item = items.FirstOrDefault(i => i.Id == drop.ItemId);
                if (item != null)
                {
                    for (int i = 0; i < drop.Quantity; i++)
                    {
                        var itemDisplay = CreateDroppedItemDisplay(item, drop);
                        wrapPanel.Children.Add(itemDisplay);
                    }
                }
            }

            var border = new Border
            {
                Background = (SolidColorBrush)FindResource("CardBrush"),
                BorderBrush = (SolidColorBrush)FindResource("BorderBrush"),
                BorderThickness = new Thickness(0, 0, 1, 1),
                Padding = new Thickness(2),
                AllowDrop = true,
                Tag = new Tuple<string, string>(character.Id, weekKey),
                Child = wrapPanel
            };

            border.PreviewDragEnter += WeekCell_DragEnter;
            border.PreviewDragOver += WeekCell_DragOver;
            border.PreviewDragLeave += WeekCell_DragLeave;
            border.PreviewDrop += WeekCell_Drop;

            return border;
        }

        private UIElement CreateDroppedItemDisplay(LootItem item, WeeklyDrop drop)
        {
            var imagePath = GetImagePath(item.ImageFileName);
            UIElement content;

            if (!string.IsNullOrEmpty(imagePath))
            {
                content = new Image
                {
                    Source = new BitmapImage(new Uri(imagePath)),
                    Width = 26,
                    Height = 26,
                    ToolTip = item.Name
                };
                RenderOptions.SetBitmapScalingMode((Image)content, BitmapScalingMode.HighQuality);
            }
            else
            {
                content = new TextBlock
                {
                    Text = item.Name.Length > 2 ? item.Name.Substring(0, 2) : item.Name,
                    FontSize = 8,
                    Foreground = (SolidColorBrush)FindResource("TextBrush"),
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    ToolTip = item.Name
                };
            }

            var border = new Border
            {
                Width = 30,
                Height = 30,
                Background = (SolidColorBrush)FindResource("ItemBackgroundBrush"),
                CornerRadius = new CornerRadius(4),
                Child = content,
                Cursor = Cursors.Hand,
                Tag = new Tuple<string, string, string>(drop.CharacterId, drop.WeekKey, drop.ItemId)
            };

            border.MouseRightButtonUp += DroppedItem_RightClick;
            border.ToolTip = $"{item.Name}\nRight-click to remove";

            var includeCheckbox = new CheckBox
            {
                Width = 9,
                Height = 9,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 1, 1, 0),
                IsChecked = drop.IncludeInSummary,
                Tag = new Tuple<string, string, string>(drop.CharacterId, drop.WeekKey, drop.ItemId),
                ToolTip = "Include in Overview?\nFor items acquired through Restore, Mystic, or Pitch Crystals.",
                Style = (Style)FindResource("OverlayItemCheckBox")
            };
            includeCheckbox.Checked += IncludeInSummaryCheckbox_Changed;
            includeCheckbox.Unchecked += IncludeInSummaryCheckbox_Changed;

            var container = new Grid
            {
                Width = 34,
                Height = 34,
                Margin = new Thickness(1)
            };
            container.Children.Add(border);
            container.Children.Add(includeCheckbox);

            return container;
        }

        private void DroppedItem_RightClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is Border border && border.Tag is Tuple<string, string, string> data)
            {
                _dataService.RemoveDropFromCell(data.Item1, data.Item2, data.Item3);
                RefreshCalendarGrid();
            }
        }

        private void IncludeInSummaryCheckbox_Changed(object sender, RoutedEventArgs e)
        {
            if (sender is not CheckBox checkBox || checkBox.Tag is not Tuple<string, string, string> data)
            {
                return;
            }

            var includeInSummary = checkBox.IsChecked != false;
            _dataService.SetDropIncludeInSummary(data.Item1, data.Item2, data.Item3, includeInSummary);

            if (MainTabControl.SelectedIndex == 1)
            {
                RefreshOverview();
            }
        }

        private void RefreshInventoryPanel()
        {
            InventoryPanel.Items.Clear();
            var items = _dataService.GetItemsFromAssets();

            if (items.Count == 0)
            {
                NoItemsText.Visibility = Visibility.Visible;
                return;
            }

            NoItemsText.Visibility = Visibility.Collapsed;

            // Sort items: regular items, then grindstones, then hammers (rightmost)
            var sortedItems = items
                .OrderBy(i =>
                {
                    if (i.Name.Contains("Hammer", StringComparison.OrdinalIgnoreCase)) return 2;
                    if (i.Name.Contains("Grindstone", StringComparison.OrdinalIgnoreCase)) return 1;
                    return 0;
                })
                .ThenBy(i => i.Name)
                .ToList();

            foreach (var item in sortedItems)
            {
                var itemControl = CreateInventoryItem(item);
                InventoryPanel.Items.Add(itemControl);
            }
        }

        private Border CreateInventoryItem(LootItem item)
        {
            var stack = new StackPanel { HorizontalAlignment = HorizontalAlignment.Center };

            var imagePath = GetImagePath(item.ImageFileName);
            if (!string.IsNullOrEmpty(imagePath))
            {
                var image = new Image
                {
                    Source = new BitmapImage(new Uri(imagePath)),
                    Width = 36,
                    Height = 36
                };
                RenderOptions.SetBitmapScalingMode(image, BitmapScalingMode.HighQuality);
                stack.Children.Add(image);
            }
            else
            {
                stack.Children.Add(new Border
                {
                    Width = 36,
                    Height = 36,
                    Background = (SolidColorBrush)FindResource("BackgroundDarkBrush"),
                    CornerRadius = new CornerRadius(6),
                    Child = new TextBlock
                    {
                        Text = "?",
                        FontSize = 16,
                        Foreground = (SolidColorBrush)FindResource("TextMutedBrush"),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    }
                });
            }

            stack.Children.Add(new TextBlock
            {
                Text = item.Name,
                FontSize = 9,
                LineHeight = 11,
                Foreground = (SolidColorBrush)FindResource("TextBrush"),
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                MaxWidth = 62,
                Margin = new Thickness(0, 2, 0, 0)
            });

            var border = new Border
            {
                Width = 70,
                Background = (SolidColorBrush)FindResource("ItemBackgroundBrush"),
                CornerRadius = new CornerRadius(6),
                Padding = new Thickness(5),
                Margin = new Thickness(3),
                Cursor = Cursors.Hand,
                Child = stack,
                Tag = item
            };

            border.MouseLeftButtonDown += InventoryItem_MouseLeftButtonDown;
            border.MouseMove += InventoryItem_MouseMove;

            return border;
        }

        private Point _dragStartPoint;
        private bool _isDragging;
        private Point _characterDragStartPoint;
        private bool _isCharacterDragging;

        private void InventoryItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragStartPoint = e.GetPosition(null);
            _isDragging = false;
        }

        private void InventoryItem_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
                return;

            Point mousePos = e.GetPosition(null);
            Vector diff = _dragStartPoint - mousePos;

            if (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance)
            {
                if (sender is Border border && border.Tag is LootItem item && !_isDragging)
                {
                    _isDragging = true;

                    // Create drag visual window
                    var imagePath = GetImagePath(item.ImageFileName);
                    var dragContent = new Image
                    {
                        Width = 36,
                        Height = 36,
                        Opacity = 0.85
                    };

                    if (!string.IsNullOrEmpty(imagePath))
                    {
                        dragContent.Source = new BitmapImage(new Uri(imagePath));
                        RenderOptions.SetBitmapScalingMode(dragContent, BitmapScalingMode.HighQuality);
                    }

                    _dragWindow = new Window
                    {
                        Content = dragContent,
                        WindowStyle = WindowStyle.None,
                        AllowsTransparency = true,
                        Background = Brushes.Transparent,
                        Width = 36,
                        Height = 36,
                        Topmost = true,
                        ShowInTaskbar = false,
                        IsHitTestVisible = false,
                        ShowActivated = false
                    };

                    // Position at cursor
                    if (GetCursorPos(out POINT pt))
                    {
                        _dragWindow.Left = pt.X - 18;
                        _dragWindow.Top = pt.Y - 18;
                    }
                    _dragWindow.Show();
                    MakeWindowClickThrough(_dragWindow);

                    this.QueryContinueDrag += MainWindow_QueryContinueDrag;

                    var dataObject = new DataObject("LootItem", item);
                    DragDrop.DoDragDrop(border, dataObject, DragDropEffects.Copy);

                    this.QueryContinueDrag -= MainWindow_QueryContinueDrag;

                    if (_dragWindow != null)
                    {
                        _dragWindow.Close();
                        _dragWindow = null;
                    }

                    _isDragging = false;
                }
            }
        }

        private void MainWindow_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
        {
            if (_dragWindow != null && GetCursorPos(out POINT pt))
            {
                _dragWindow.Left = pt.X - 18;
                _dragWindow.Top = pt.Y - 18;
            }
        }

        private void WeekCell_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("LootItem"))
            {
                if (sender is Border border)
                {
                    border.Background = new SolidColorBrush(Color.FromArgb(0x55, 0xFF, 0x8C, 0x00));
                }
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void WeekCell_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("LootItem"))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void WeekCell_DragLeave(object sender, DragEventArgs e)
        {
            if (sender is Border border)
            {
                border.Background = (SolidColorBrush)FindResource("CardBrush");
            }
            e.Handled = true;
        }

        private void WeekCell_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent("LootItem"))
            {
                if (sender is Border border && border.Tag is Tuple<string, string> data)
                {
                    var characterId = data.Item1;
                    var weekKey = data.Item2;
                    var item = (LootItem)e.Data.GetData("LootItem")!;
                    _dataService.AddDropToCell(characterId, weekKey, item.Id);
                    RefreshCalendarGrid();
                }
            }
            e.Handled = true;
        }

        private string GetImagePath(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
                return string.Empty;

            var fullPath = Path.Combine(_assetsPath, fileName);
            return File.Exists(fullPath) ? fullPath : string.Empty;
        }

        private void AddCharacter_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new AddCharacterDialog();
            dialog.Owner = this;
            if (dialog.ShowDialog() == true)
            {
                _dataService.AddCharacter(
                    dialog.CharacterName,
                    dialog.CharacterClass,
                    dialog.CharacterLevel);
                RefreshUI();
            }
        }

        private void EditCharacter_Click(object sender, RoutedEventArgs e)
        {
            if (sender is not Button button || button.Tag is not string id)
            {
                return;
            }

            var character = _dataService.Characters.FirstOrDefault(c => c.Id == id);
            if (character == null)
            {
                return;
            }

            var dialog = new AddCharacterDialog(
                character.Name,
                character.Class,
                character.Level,
                isEditMode: true);
            dialog.Owner = this;

            if (dialog.ShowDialog() == true)
            {
                _dataService.UpdateCharacter(
                    id,
                    dialog.CharacterName,
                    dialog.CharacterClass,
                    dialog.CharacterLevel);
                RefreshUI();
            }
        }

        private void DeleteCharacter_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string id)
            {
                var result = MessageBox.Show(
                    "Delete this character? This will also remove their drop history.",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    _dataService.DeleteCharacter(id);
                    RefreshUI();
                }
            }
        }

        private void DarkModeToggle_Click(object sender, MouseButtonEventArgs e)
        {
            _isDarkMode = !_isDarkMode;
            ApplyTheme(_isDarkMode);
            UpdateDarkModeToggleVisual();
            _dataService.SetDarkModeSetting(_isDarkMode);
            RefreshUI();
        }

        private void UpdateDarkModeToggleVisual()
        {
            if (_isDarkMode)
            {
                DarkModeToggle.Background = (SolidColorBrush)FindResource("PrimaryBrush");
                DarkModeThumb.HorizontalAlignment = HorizontalAlignment.Right;
                DarkModeThumb.Margin = new Thickness(0, 0, 2, 0);
            }
            else
            {
                DarkModeToggle.Background = (SolidColorBrush)FindResource("BorderBrush");
                DarkModeThumb.HorizontalAlignment = HorizontalAlignment.Left;
                DarkModeThumb.Margin = new Thickness(2, 0, 0, 0);
            }
        }

        private void ApplyTheme(bool isDark)
        {
            var resources = Application.Current.Resources;

            if (isDark)
            {
                // Dark mode colors
                resources["BackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
                resources["BackgroundDarkBrush"] = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
                resources["CardBrush"] = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
                resources["TextBrush"] = (SolidColorBrush)FindResource("BorderBrush");
                resources["TextMutedBrush"] = new SolidColorBrush(Color.FromRgb(0x9E, 0x9E, 0x9E));
                resources["BorderBrush"] = new SolidColorBrush(Color.FromRgb(0x40, 0x40, 0x40));
                resources["InputBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x3D, 0x3D));
                resources["CellBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
                resources["CharacterCellBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x2D, 0x1A));
                resources["ItemBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x3D, 0x3D));
                resources["TotalCellBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x2D, 0x1A));
                resources["ScrollBarTrackBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x3D, 0x3D));
                resources["TabInactiveBrush"] = new SolidColorBrush(Color.FromRgb(0x3D, 0x3D, 0x3D));
                resources["DeleteHoverBrush"] = new SolidColorBrush(Color.FromRgb(0x4A, 0x2A, 0x2A));
                resources["EditHoverBrush"] = new SolidColorBrush(Color.FromRgb(0x5A, 0x43, 0x28));
            }
            else
            {
                // Light mode colors
                resources["BackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA));
                resources["BackgroundDarkBrush"] = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0));
                resources["CardBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
                resources["TextBrush"] = new SolidColorBrush(Color.FromRgb(0x2D, 0x2D, 0x2D));
                resources["TextMutedBrush"] = new SolidColorBrush(Color.FromRgb(0x75, 0x75, 0x75));
                resources["BorderBrush"] = (SolidColorBrush)FindResource("BorderBrush");
                resources["InputBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
                resources["CellBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
                resources["CharacterCellBrush"] = (SolidColorBrush)FindResource("CharacterCellBrush");
                resources["ItemBackgroundBrush"] = new SolidColorBrush(Color.FromRgb(0xFA, 0xFA, 0xFA));
                resources["TotalCellBrush"] = (SolidColorBrush)FindResource("CharacterCellBrush");
                resources["ScrollBarTrackBrush"] = new SolidColorBrush(Color.FromRgb(0xF0, 0xF0, 0xF0));
                resources["TabInactiveBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF));
                resources["DeleteHoverBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xEB, 0xEE));
                resources["EditHoverBrush"] = new SolidColorBrush(Color.FromRgb(0xFF, 0xF5, 0xE6));
            }
        }
    }
}
