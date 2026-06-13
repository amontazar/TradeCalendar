// TradeCalendarAddOn_v3.cs
// NinjaTrader 8 AddOn
// v2 adds CSV backfill/import for historical missed days and colors positive/negative
// day-grid rows with light green / light red backgrounds.
//
// Place this file in:
// Documents\NinjaTrader 8\bin\Custom\AddOns\TradeCalendarAddOn_v3.cs
//
// Compile from NinjaTrader's NinjaScript editor, then open from New > Trade Calendar.

#region Using declarations
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml.Linq;
using System.Xml.Serialization;
using Microsoft.Win32;
using NinjaTrader.Cbi;
using NinjaTrader.Gui;
using NinjaTrader.Gui.Tools;
using NinjaTrader.NinjaScript;
#endregion

namespace NinjaTrader.NinjaScript.AddOns
{
    public class TradeCalendarAddOn_v3 : AddOnBase
    {
        private NTMenuItem menuItem;
        private NTMenuItem newMenu;
        private static TradeCalendarWindow_v3 windowRef;

        protected override void OnStateChange()
        {
            if (State == State.SetDefaults)
            {
                Name = "TradeCalendarAddOn_v3";
                Description = "Calendar-style realized PnL dashboard for account executions with CSV backfill.";
            }
            else if (State == State.Active)
            {
                TradeLedgerService_v3.Instance.EnsureInitialized();
            }
            else if (State == State.Terminated)
            {
                TradeLedgerService_v3.Instance.Shutdown();

                if (menuItem != null)
                    menuItem.Click -= OnMenuItemClick;
            }
        }

        protected override void OnWindowCreated(Window window)
        {
            ControlCenter cc = window as ControlCenter;
            if (cc == null)
                return;

            if (menuItem != null)
                return;

            newMenu = cc.FindFirst("ControlCenterMenuItemNew") as NTMenuItem;
            if (newMenu == null)
                return;

            menuItem = new NTMenuItem
            {
                Header = "Trade Calendar",
                Style = Application.Current.TryFindResource("MainMenuItem") as Style
            };
            AutomationProperties.SetAutomationId(menuItem, "TradeCalendarMenuItemV3");
            menuItem.Click += OnMenuItemClick;

            bool alreadyAdded = false;
            foreach (object item in newMenu.Items)
            {
                NTMenuItem existing = item as NTMenuItem;
                if (existing != null && Equals(existing.Header, menuItem.Header))
                {
                    alreadyAdded = true;
                    break;
                }
            }

            if (!alreadyAdded)
                newMenu.Items.Add(menuItem);
        }

        protected override void OnWindowDestroyed(Window window)
        {
            if (!(window is ControlCenter))
                return;

            if (menuItem != null)
                menuItem.Click -= OnMenuItemClick;

            if (newMenu != null && menuItem != null && newMenu.Items.Contains(menuItem))
                newMenu.Items.Remove(menuItem);

            menuItem = null;
            newMenu = null;
        }

        private void OnMenuItemClick(object sender, RoutedEventArgs e)
        {
            try
            {
                TradeLedgerService_v3.Instance.EnsureInitialized();

                if (windowRef == null || !windowRef.IsLoaded)
                {
                    windowRef = new TradeCalendarWindow_v3();
                    windowRef.Closed += (o, args) => windowRef = null;
                    windowRef.Show();
                }
                else
                {
                    if (windowRef.WindowState == WindowState.Minimized)
                        windowRef.WindowState = WindowState.Normal;

                    windowRef.Activate();
                    windowRef.Focus();
                }
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendarAddOn_v3 error: " + ex, PrintTo.OutputTab1);
            }
        }
    }

    public class TradeCalendarWindow_v3 : NTWindow, IWorkspacePersistence
    {
        public WorkspaceOptions WorkspaceOptions { get; set; }

        public TradeCalendarWindow_v3()
        {
            Caption = "Trade Calendar";
            Width = 1400;
            Height = 940;
            MinWidth = 1050;
            MinHeight = 720;

            Content = new TradeCalendarControl_v3();

            Loaded += (o, e) =>
            {
                if (WorkspaceOptions == null)
                    WorkspaceOptions = new WorkspaceOptions("TradeCalendarWindowV3-" + Guid.NewGuid().ToString("N"), this);
            };
        }

        public void Restore(XDocument document, XElement element)
        {
        }

        public void Save(XDocument document, XElement element)
        {
        }
    }


    public class TradeCalendarControl_v3 : Grid
    {
        private readonly TradeLedgerService_v3 service = TradeLedgerService_v3.Instance;

        private AccountSelector accountSelector;
        private CheckBox allAccountsCheckBox;
        private ComboBox instrumentComboBox;
        private ComboBox sourceComboBox;
        private ComboBox sideComboBox;
        private ComboBox resultComboBox;
        private DatePicker fromDatePicker;
        private DatePicker toDatePicker;
        private Button thisMonthFilterButton;
        private Button last30DaysFilterButton;
        private Button clearFiltersButton;
        private Button prevButton;
        private Button nextButton;
        private Button todayButton;
        private Button refreshButton;
        private Button importCsvButton;
        private Button exportCsvButton;
        private TextBlock monthTitle;
        private TextBlock monthTotalText;
        private TextBlock accountCashValueText;
        private Grid calendarGrid;
        private TextBlock dayHeaderText;
        private TextBlock dayPnlText;
        private StackPanel instrumentSummaryPanel;
        private ComboBox tradeListViewComboBox;
        private Button exportCurrentTradeListCsvButton;
        private TextBlock tradesHeaderText;
        private DataGrid dayGrid;
        private TextBlock activeFilterSummaryText;

        private DateTime currentMonth;
        private DateTime selectedDate;
        private bool isApplyingUiState;
        private bool isRefreshingFilterChoices;
        private bool isApplyingAccountSelection;
        private string lastSelectedAccountName;
        private string pendingRestoreAccountName;
        private DispatcherTimer accountRestoreTimer;
        private int accountRestoreAttempts;
        private bool isStartupAccountRestorePending;

        private const int AccountRestoreMaxAttempts = 20;
        private static readonly TimeSpan AccountRestoreInterval = TimeSpan.FromMilliseconds(250);

        private readonly Brush positiveBrush = new SolidColorBrush(Color.FromRgb(56, 142, 60));
        private readonly Brush negativeBrush = new SolidColorBrush(Color.FromRgb(198, 40, 40));
        private readonly Brush neutralBrush = new SolidColorBrush(Color.FromRgb(120, 120, 120));
        private readonly Brush softPositiveBrush = new SolidColorBrush(Color.FromRgb(219, 238, 220));
        private readonly Brush softNegativeBrush = new SolidColorBrush(Color.FromRgb(248, 224, 224));
        private readonly Brush selectedBorderBrush = new SolidColorBrush(Color.FromRgb(31, 96, 196));

        private const string AllInstrumentsText = "All instruments";
        private const string AllSourcesText = "All sources";
        private const string AllSidesText = "All sides";
        private const string AllResultsText = "All results";
        private const string MatchedTradesViewText = "Matched Trades";
        private const string ExecutionLegsViewText = "Execution Legs";

        public TradeCalendarControl_v3()
        {
            ResetCalendarToToday();

            isApplyingUiState = true;
            try
            {
                BuildUi();
            }
            finally
            {
                isApplyingUiState = false;
            }

            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            service.DataChanged += OnServiceDataChanged;
            service.EnsureInitialized();
            service.RefreshCurrentSessionSnapshots();

            isApplyingUiState = true;
            try
            {
                LoadSavedUiState();
                RefreshFilterChoices();
            }
            finally
            {
                isApplyingUiState = false;
            }

            RefreshAll();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            StopAccountRestoreTimer();
            service.DataChanged -= OnServiceDataChanged;
            if (!isStartupAccountRestorePending)
                SaveCurrentUiState();
        }

        private void OnServiceDataChanged(object sender, EventArgs e)
        {
            Dispatcher.InvokeAsync(() =>
            {
                RefreshFilterChoices();
                if (isStartupAccountRestorePending && TryRestorePendingAccountSelection())
                    return;

                RefreshAll();
            });
        }

        private void BuildUi()
        {
            RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            RowDefinitions.Add(new RowDefinition { Height = new GridLength(280) });

            BuildToolbar();
            BuildFilterBar();
            BuildActiveFilterSummary();
            BuildMonthSummary();
            BuildCalendarArea();
            BuildDetailArea();
        }

        private void BuildToolbar()
        {
            Grid toolbar = new Grid { Margin = new Thickness(10, 10, 10, 6) };
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            toolbar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            prevButton = MakeButton("<");
            nextButton = MakeButton(">");
            todayButton = MakeButton("Today");
            refreshButton = MakeButton("Refresh");
            importCsvButton = MakeButton("Import CSV");
            exportCsvButton = MakeButton("Export CSV");

            prevButton.Click += (s, e) => { currentMonth = currentMonth.AddMonths(-1); RefreshAll(); };
            nextButton.Click += (s, e) => { currentMonth = currentMonth.AddMonths(1); RefreshAll(); };
            todayButton.Click += (s, e) =>
            {
                ResetCalendarToToday();
                RefreshAll();
            };
            refreshButton.Click += (s, e) =>
            {
                service.RefreshCurrentSessionSnapshots();
                RefreshFilterChoices();
                RefreshAll();
            };
            importCsvButton.Click += OnImportCsvClick;
            exportCsvButton.Click += OnExportCsvClick;

            monthTitle = new TextBlock
            {
                FontSize = 30,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center
            };

            Grid.SetColumn(prevButton, 0);
            Grid.SetColumn(nextButton, 1);
            Grid.SetColumn(todayButton, 2);
            Grid.SetColumn(refreshButton, 3);
            Grid.SetColumn(importCsvButton, 4);
            Grid.SetColumn(exportCsvButton, 5);
            Grid.SetColumn(monthTitle, 7);

            toolbar.Children.Add(prevButton);
            toolbar.Children.Add(nextButton);
            toolbar.Children.Add(todayButton);
            toolbar.Children.Add(refreshButton);
            toolbar.Children.Add(importCsvButton);
            toolbar.Children.Add(exportCsvButton);
            toolbar.Children.Add(monthTitle);

            Children.Add(toolbar);
            Grid.SetRow(toolbar, 0);
        }

        private void BuildFilterBar()
        {
            Grid filterBar = new Grid { Margin = new Thickness(10, 0, 10, 6) };
            for (int i = 0; i < 12; i++)
                filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            filterBar.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            accountSelector = new AccountSelector
            {
                Width = 170,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center
            };
            accountSelector.SelectionChanged += OnAccountSelectionChanged;

            allAccountsCheckBox = new CheckBox
            {
                Content = "All accounts",
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 10, 0)
            };
            allAccountsCheckBox.Checked += OnAllAccountsChanged;
            allAccountsCheckBox.Unchecked += OnAllAccountsChanged;

            instrumentComboBox = MakeFilterComboBox(170, AllInstrumentsText);
            sourceComboBox = MakeFilterComboBox(120, AllSourcesText, new[] { AllSourcesText, "Live", "Imported" });
            sideComboBox = MakeFilterComboBox(110, AllSidesText, new[] { AllSidesText, "Long", "Short" });
            resultComboBox = MakeFilterComboBox(120, AllResultsText, new[] { AllResultsText, "Winning", "Losing", "Breakeven", "Open" });

            fromDatePicker = MakeDatePicker();
            toDatePicker = MakeDatePicker();
            fromDatePicker.SelectedDateChanged += OnAnyFilterChanged;
            toDatePicker.SelectedDateChanged += OnAnyFilterChanged;

            thisMonthFilterButton = MakeButton("This Month");
            last30DaysFilterButton = MakeButton("Last 30");
            clearFiltersButton = MakeButton("Clear Filters");

            thisMonthFilterButton.Click += (s, e) =>
            {
                isApplyingUiState = true;
                try
                {
                    DateTime first = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                    fromDatePicker.SelectedDate = first;
                    toDatePicker.SelectedDate = first.AddMonths(1).AddDays(-1);
                }
                finally
                {
                    isApplyingUiState = false;
                }
                OnAnyFilterChanged(s, e);
            };

            last30DaysFilterButton.Click += (s, e) =>
            {
                isApplyingUiState = true;
                try
                {
                    fromDatePicker.SelectedDate = DateTime.Today.AddDays(-29);
                    toDatePicker.SelectedDate = DateTime.Today;
                }
                finally
                {
                    isApplyingUiState = false;
                }
                OnAnyFilterChanged(s, e);
            };

            clearFiltersButton.Click += (s, e) => ClearNonAccountFilters();

            AddLabeledControl(filterBar, 0, "Account", accountSelector);
            Grid.SetColumn(allAccountsCheckBox, 1);
            filterBar.Children.Add(allAccountsCheckBox);
            AddLabeledControl(filterBar, 2, "Instrument", instrumentComboBox);
            AddLabeledControl(filterBar, 3, "Source", sourceComboBox);
            AddLabeledControl(filterBar, 4, "Side", sideComboBox);
            AddLabeledControl(filterBar, 5, "Result", resultComboBox);
            AddLabeledControl(filterBar, 6, "From", fromDatePicker);
            AddLabeledControl(filterBar, 7, "To", toDatePicker);

            Grid.SetColumn(thisMonthFilterButton, 8);
            Grid.SetColumn(last30DaysFilterButton, 9);
            Grid.SetColumn(clearFiltersButton, 10);
            filterBar.Children.Add(thisMonthFilterButton);
            filterBar.Children.Add(last30DaysFilterButton);
            filterBar.Children.Add(clearFiltersButton);

            Children.Add(filterBar);
            Grid.SetRow(filterBar, 1);
        }

        private void BuildActiveFilterSummary()
        {
            Border border = new Border
            {
                Margin = new Thickness(10, 0, 10, 6),
                Padding = new Thickness(10, 6, 10, 6),
                BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Background = new SolidColorBrush(Color.FromRgb(249, 249, 249))
            };

            activeFilterSummaryText = new TextBlock
            {
                FontSize = 13,
                Foreground = neutralBrush,
                Text = "Filters: none"
            };

            border.Child = activeFilterSummaryText;
            Children.Add(border);
            Grid.SetRow(border, 2);
        }

        private void OnAccountSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isApplyingUiState || isRefreshingFilterChoices || isApplyingAccountSelection || !IsLoaded || !IsUiReady())
                return;

            string selectedAccountName = GetSelectedAccountNameFromEvent(e) ?? GetCurrentlySelectedAccountName();
            if (string.IsNullOrWhiteSpace(selectedAccountName))
                return;

            if (isStartupAccountRestorePending)
            {
                if (string.Equals(selectedAccountName, pendingRestoreAccountName, StringComparison.OrdinalIgnoreCase))
                {
                    TryRestorePendingAccountSelection();
                    return;
                }

                if (accountSelector == null || (!accountSelector.IsKeyboardFocusWithin && !accountSelector.IsMouseOver))
                    return;

                StopAccountRestoreTimer();
                pendingRestoreAccountName = null;
                isStartupAccountRestorePending = false;
            }

            lastSelectedAccountName = selectedAccountName;
            SaveCurrentUiState();
            RefreshFilterChoices();
            RefreshAll();
        }

        private string GetSelectedAccountNameFromEvent(SelectionChangedEventArgs e)
        {
            if (e == null || e.AddedItems == null)
                return null;

            foreach (object item in e.AddedItems)
            {
                Account accountItem = item as Account;
                if (accountItem != null && !string.IsNullOrWhiteSpace(accountItem.Name))
                    return accountItem.Name;

                Account knownAccount = FindAccountByName(GetAccountNameFromItem(item));
                if (knownAccount != null)
                    return knownAccount.Name;
            }

            return null;
        }

        private void OnAllAccountsChanged(object sender, RoutedEventArgs e)
        {
            if (isApplyingUiState || isApplyingAccountSelection || !IsLoaded || !IsUiReady())
                return;

            if (isStartupAccountRestorePending)
            {
                StopAccountRestoreTimer();
                pendingRestoreAccountName = null;
                isStartupAccountRestorePending = false;
            }

            bool useAllAccounts = allAccountsCheckBox.IsChecked == true;
            accountSelector.IsEnabled = !useAllAccounts;

            if (!useAllAccounts)
            {
                string selectedAccountName = GetCurrentlySelectedAccountName();
                if (!string.IsNullOrWhiteSpace(lastSelectedAccountName)
                    && !string.Equals(selectedAccountName, lastSelectedAccountName, StringComparison.OrdinalIgnoreCase))
                {
                    TrySelectAccountByNameSimple(lastSelectedAccountName);
                    selectedAccountName = GetCurrentlySelectedAccountName();
                }

                if (string.IsNullOrWhiteSpace(selectedAccountName))
                {
                    if (!TrySelectFirstAvailableAccount())
                    {
                        SelectFallbackAccountOrAllAccounts();
                        useAllAccounts = true;
                    }
                    selectedAccountName = GetCurrentlySelectedAccountName();
                }

                if (!string.IsNullOrWhiteSpace(selectedAccountName))
                    lastSelectedAccountName = selectedAccountName;
            }

            SaveCurrentUiState();
            RefreshFilterChoices();
            RefreshAll();
        }

        private void OnAnyFilterChanged(object sender, RoutedEventArgs e)
        {
            if (isApplyingUiState || isRefreshingFilterChoices || !IsLoaded || !IsUiReady())
                return;

            if (accountSelector != null && allAccountsCheckBox != null)
                accountSelector.IsEnabled = allAccountsCheckBox.IsChecked != true;

            SaveCurrentUiState();
            RefreshAll();
        }

        private ComboBox MakeFilterComboBox(double width, string defaultItem, IEnumerable<string> items = null)
        {
            ComboBox combo = new ComboBox
            {
                Width = width,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                IsEditable = false
            };
            if (items != null)
            {
                foreach (string item in items)
                    combo.Items.Add(item);
            }
            else
            {
                combo.Items.Add(defaultItem);
            }

            combo.SelectedIndex = 0;
            combo.SelectionChanged += OnAnyFilterChanged;
            return combo;
        }

        private DatePicker MakeDatePicker()
        {
            return new DatePicker
            {
                Width = 122,
                Margin = new Thickness(0, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                SelectedDateFormat = DatePickerFormat.Short
            };
        }

        private void AddLabeledControl(Grid grid, int column, string label, Control control)
        {
            StackPanel panel = new StackPanel
            {
                Orientation = Orientation.Vertical,
                Margin = new Thickness(0, 0, 8, 0)
            };

            panel.Children.Add(new TextBlock
            {
                Text = label,
                FontSize = 11,
                Foreground = neutralBrush,
                Margin = new Thickness(0, 0, 0, 2)
            });
            panel.Children.Add(control);

            Grid.SetColumn(panel, column);
            grid.Children.Add(panel);
        }

        private void ClearNonAccountFilters()
        {
            isApplyingUiState = true;
            try
            {
                if (instrumentComboBox.Items.Count > 0)
                    instrumentComboBox.SelectedIndex = 0;
                if (sourceComboBox.Items.Count > 0)
                    sourceComboBox.SelectedIndex = 0;
                if (sideComboBox.Items.Count > 0)
                    sideComboBox.SelectedIndex = 0;
                if (resultComboBox.Items.Count > 0)
                    resultComboBox.SelectedIndex = 0;
                fromDatePicker.SelectedDate = null;
                toDatePicker.SelectedDate = null;
            }
            finally
            {
                isApplyingUiState = false;
            }

            OnAnyFilterChanged(this, new RoutedEventArgs());
        }

        private void RefreshFilterChoices()
        {
            isRefreshingFilterChoices = true;
            try
            {
                string preservedInstrument = instrumentComboBox != null && instrumentComboBox.SelectedItem != null
                    ? instrumentComboBox.SelectedItem.ToString()
                    : AllInstrumentsText;

                string accountName = GetSelectedAccountName();
                List<string> instruments = service.GetAvailableInstruments(accountName);

                instrumentComboBox.Items.Clear();
                instrumentComboBox.Items.Add(AllInstrumentsText);
                foreach (string instrument in instruments)
                    instrumentComboBox.Items.Add(instrument);

                if (!string.IsNullOrWhiteSpace(preservedInstrument) && instrumentComboBox.Items.Cast<object>().Any(i => string.Equals(i.ToString(), preservedInstrument, StringComparison.OrdinalIgnoreCase)))
                    instrumentComboBox.SelectedItem = preservedInstrument;
                else
                    instrumentComboBox.SelectedIndex = 0;

                accountSelector.IsEnabled = allAccountsCheckBox.IsChecked != true;
            }
            finally
            {
                isRefreshingFilterChoices = false;
            }
        }

        private void LoadSavedUiState()
        {
            TradeCalendarUiState state = TradeCalendarUiState.Load();
            ResetCalendarToToday();

            lastSelectedAccountName = GetSavedAccountName(state);
            bool useAllAccounts = state != null && state.AllAccounts;
            allAccountsCheckBox.IsChecked = useAllAccounts;
            accountSelector.IsEnabled = !useAllAccounts;
            pendingRestoreAccountName = null;
            isStartupAccountRestorePending = false;
            StopAccountRestoreTimer();

            if (!useAllAccounts)
            {
                if (!string.IsNullOrWhiteSpace(lastSelectedAccountName))
                {
                    pendingRestoreAccountName = lastSelectedAccountName;
                    isStartupAccountRestorePending = true;
                    if (!TryRestorePendingAccountSelection())
                        StartAccountRestoreTimer();
                }
                else if (!TrySelectFirstAvailableAccount())
                {
                    SelectFallbackAccountOrAllAccounts();
                }
            }

            ApplyComboSelection(sourceComboBox, state == null || string.IsNullOrWhiteSpace(state.SourceFilter) ? AllSourcesText : state.SourceFilter);
            ApplyComboSelection(sideComboBox, state == null || string.IsNullOrWhiteSpace(state.SideFilter) ? AllSidesText : state.SideFilter);
            ApplyComboSelection(resultComboBox, state == null || string.IsNullOrWhiteSpace(state.ResultFilter) ? AllResultsText : state.ResultFilter);

            fromDatePicker.SelectedDate = state != null ? state.FromDate : null;
            toDatePicker.SelectedDate = state != null ? state.ToDate : null;

            RefreshFilterChoices();

            string wantedInstrument = state == null || string.IsNullOrWhiteSpace(state.InstrumentFilter) ? AllInstrumentsText : state.InstrumentFilter;
            if (instrumentComboBox.Items.Cast<object>().Any(i => string.Equals(i.ToString(), wantedInstrument, StringComparison.OrdinalIgnoreCase)))
                instrumentComboBox.SelectedItem = wantedInstrument;
            else if (instrumentComboBox.Items.Count > 0)
                instrumentComboBox.SelectedIndex = 0;

            string wantedTradeListView = state == null || string.IsNullOrWhiteSpace(state.TradeListViewMode)
                ? MatchedTradesViewText
                : state.TradeListViewMode;
            ApplyComboSelection(tradeListViewComboBox, wantedTradeListView);
            ConfigureDayGridColumns();
        }

        private void ResetCalendarToToday()
        {
            DateTime today = DateTime.Today;
            currentMonth = new DateTime(today.Year, today.Month, 1);
            selectedDate = today;
        }

        private void SaveCurrentUiState()
        {
            if (isStartupAccountRestorePending)
                return;

            try
            {
                bool useAllAccounts = allAccountsCheckBox != null && allAccountsCheckBox.IsChecked == true;
                string selectedAccountName = GetCurrentlySelectedAccountName();
                if (!useAllAccounts && !string.IsNullOrWhiteSpace(selectedAccountName))
                    lastSelectedAccountName = selectedAccountName;
                string persistedAccountName = useAllAccounts
                    ? lastSelectedAccountName ?? string.Empty
                    : selectedAccountName ?? lastSelectedAccountName ?? string.Empty;

                TradeCalendarUiState state = new TradeCalendarUiState
                {
                    AccountName = useAllAccounts ? string.Empty : persistedAccountName,
                    AllAccounts = useAllAccounts,
                    LastSelectedAccountName = lastSelectedAccountName ?? string.Empty,
                    LastSelectedAccountIndex = null,
                    InstrumentFilter = GetComboSelection(instrumentComboBox, AllInstrumentsText),
                    SourceFilter = GetComboSelection(sourceComboBox, AllSourcesText),
                    SideFilter = GetComboSelection(sideComboBox, AllSidesText),
                    ResultFilter = GetComboSelection(resultComboBox, AllResultsText),
                    FromDate = fromDatePicker != null ? fromDatePicker.SelectedDate : null,
                    ToDate = toDatePicker != null ? toDatePicker.SelectedDate : null,
                    CurrentMonth = currentMonth,
                    SelectedDate = selectedDate,
                    TradeListViewMode = GetTradeListViewMode()
                };
                state.Save();
                NinjaTrader.Code.Output.Process(
                    "TradeCalendar saved account='" + persistedAccountName + "', allAccounts=" + useAllAccounts,
                    PrintTo.OutputTab1);
            }
            catch
            {
            }
        }

        private void ApplyComboSelection(ComboBox combo, string value)
        {
            if (combo == null)
                return;

            foreach (object item in combo.Items)
            {
                if (string.Equals(item.ToString(), value, StringComparison.OrdinalIgnoreCase))
                {
                    combo.SelectedItem = item;
                    return;
                }
            }

            if (combo.Items.Count > 0)
                combo.SelectedIndex = 0;
        }

        private string GetComboSelection(ComboBox combo, string defaultLabel)
        {
            if (combo == null || combo.SelectedItem == null)
                return defaultLabel;
            return combo.SelectedItem.ToString();
        }

        private Account FindAccountByName(string accountName)
        {
            if (string.IsNullOrWhiteSpace(accountName))
                return null;

            lock (Account.All)
            {
                foreach (Account account in Account.All)
                {
                    if (account != null && string.Equals(account.Name, accountName, StringComparison.OrdinalIgnoreCase))
                        return account;
                }
            }

            return null;
        }

        private void OnImportCsvClick(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.Multiselect = true;
                dialog.Title = "Import historical trade/trade-performance CSV";

                bool? ok = dialog.ShowDialog();
                if (ok != true)
                    return;

                string defaultAccount = GetSelectedAccountName();
                CsvImportResult result = service.ImportCsvTrades(dialog.FileNames, defaultAccount);
                RefreshFilterChoices();
                RefreshAll();

                MessageBox.Show(result.ToDisplayText(), "Trade Calendar Import", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed: " + ex.Message, "Trade Calendar Import", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void OnExportCsvClick(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveFileDialog dialog = new SaveFileDialog();
                dialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*";
                dialog.Title = "Export trade calendar ledgers to CSV";
                dialog.FileName = "TradeCalendarLedgerExport_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".csv";
                dialog.OverwritePrompt = true;

                bool? ok = dialog.ShowDialog();
                if (ok != true)
                    return;

                LedgerExportResult result = service.ExportLedgersToCsv(dialog.FileName);
                MessageBox.Show(result.ToDisplayText(), "Trade Calendar Export", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed: " + ex.Message, "Trade Calendar Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BuildMonthSummary()
        {
            Border border = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(215, 215, 215)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(10, 0, 10, 8),
                Padding = new Thickness(12, 8, 12, 8)
            };

            Grid panel = new Grid();
            panel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            panel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            StackPanel leftPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center
            };

            TextBlock monthLabel = new TextBlock
            {
                Text = "Month total",
                FontSize = 16,
                Margin = new Thickness(0, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center
            };

            monthTotalText = new TextBlock
            {
                FontSize = 24,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };

            leftPanel.Children.Add(monthLabel);
            leftPanel.Children.Add(monthTotalText);

            StackPanel rightPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                VerticalAlignment = VerticalAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Right
            };

            TextBlock cashLabel = new TextBlock
            {
                Text = "Account cash value",
                FontSize = 14,
                Margin = new Thickness(18, 0, 10, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = neutralBrush
            };

            accountCashValueText = new TextBlock
            {
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = neutralBrush
            };

            rightPanel.Children.Add(cashLabel);
            rightPanel.Children.Add(accountCashValueText);

            panel.Children.Add(leftPanel);
            Grid.SetColumn(leftPanel, 0);
            panel.Children.Add(rightPanel);
            Grid.SetColumn(rightPanel, 1);
            border.Child = panel;

            Children.Add(border);
            Grid.SetRow(border, 3);
        }

        private void BuildCalendarArea()
        {
            Grid host = new Grid { Margin = new Thickness(10, 0, 10, 8) };
            host.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            host.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            Grid header = new Grid();
            for (int i = 0; i < 7; i++)
                header.ColumnDefinitions.Add(new ColumnDefinition());

            string[] days = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
            for (int i = 0; i < 7; i++)
            {
                Border b = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
                    BorderThickness = new Thickness(1),
                    Background = new SolidColorBrush(Color.FromRgb(247, 247, 247)),
                    Padding = new Thickness(6)
                };

                TextBlock t = new TextBlock
                {
                    Text = days[i],
                    FontSize = 16,
                    FontWeight = FontWeights.SemiBold,
                    Foreground = selectedBorderBrush,
                    HorizontalAlignment = HorizontalAlignment.Center
                };

                b.Child = t;
                Grid.SetColumn(b, i);
                header.Children.Add(b);
            }

            calendarGrid = new Grid();
            for (int i = 0; i < 7; i++)
                calendarGrid.ColumnDefinitions.Add(new ColumnDefinition());
            for (int i = 0; i < 6; i++)
                calendarGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star), MinHeight = 56 });

            host.Children.Add(header);
            Grid.SetRow(header, 0);

            host.Children.Add(calendarGrid);
            Grid.SetRow(calendarGrid, 1);

            Children.Add(host);
            Grid.SetRow(host, 4);
        }

        private void BuildDetailArea()
        {
            Grid detail = new Grid { Margin = new Thickness(10, 0, 10, 10) };
            detail.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(320) });
            detail.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            Border leftCard = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Margin = new Thickness(0, 0, 12, 0),
                Padding = new Thickness(12)
            };

            StackPanel leftPanel = new StackPanel();

            dayHeaderText = new TextBlock
            {
                FontSize = 20,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 0, 18)
            };

            TextBlock pnlLabel = new TextBlock
            {
                Text = "Total PnL",
                FontSize = 14,
                Margin = new Thickness(0, 0, 0, 6)
            };

            dayPnlText = new TextBlock
            {
                FontSize = 42,
                FontWeight = FontWeights.Light,
                Margin = new Thickness(0, 0, 0, 18)
            };

            TextBlock byInstrumentLabel = new TextBlock
            {
                Text = "By instrument",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 4, 0, 8)
            };

            instrumentSummaryPanel = new StackPanel();

            leftPanel.Children.Add(dayHeaderText);
            leftPanel.Children.Add(pnlLabel);
            leftPanel.Children.Add(dayPnlText);
            leftPanel.Children.Add(byInstrumentLabel);
            leftPanel.Children.Add(instrumentSummaryPanel);

            leftCard.Child = leftPanel;

            Border rightCard = new Border
            {
                BorderBrush = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(4),
                Padding = new Thickness(0)
            };

            Grid rightPanel = new Grid();
            rightPanel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            rightPanel.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });

            Grid tradeListHeader = new Grid
            {
                Margin = new Thickness(12, 8, 12, 8)
            };
            tradeListHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            tradeListHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            tradeListHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            tradeListHeader.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            tradesHeaderText = new TextBlock
            {
                Text = "Matched trades",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                VerticalAlignment = VerticalAlignment.Center
            };

            TextBlock tradeListViewLabel = new TextBlock
            {
                Text = "Trade List View",
                Margin = new Thickness(12, 0, 8, 0),
                VerticalAlignment = VerticalAlignment.Center,
                Foreground = neutralBrush
            };

            tradeListViewComboBox = new ComboBox
            {
                Width = 145,
                MinHeight = 26,
                VerticalContentAlignment = VerticalAlignment.Center
            };
            tradeListViewComboBox.Items.Add(MatchedTradesViewText);
            tradeListViewComboBox.Items.Add(ExecutionLegsViewText);
            tradeListViewComboBox.SelectedItem = MatchedTradesViewText;
            tradeListViewComboBox.SelectionChanged += OnTradeListViewChanged;

            exportCurrentTradeListCsvButton = MakeButton("Export CSV");
            exportCurrentTradeListCsvButton.Margin = new Thickness(8, 0, 0, 0);
            exportCurrentTradeListCsvButton.Click += OnExportCurrentTradeListCsvClick;

            tradeListHeader.Children.Add(tradesHeaderText);
            Grid.SetColumn(tradesHeaderText, 0);
            tradeListHeader.Children.Add(tradeListViewLabel);
            Grid.SetColumn(tradeListViewLabel, 1);
            tradeListHeader.Children.Add(tradeListViewComboBox);
            Grid.SetColumn(tradeListViewComboBox, 2);
            tradeListHeader.Children.Add(exportCurrentTradeListCsvButton);
            Grid.SetColumn(exportCurrentTradeListCsvButton, 3);

            dayGrid = new DataGrid
            {
                AutoGenerateColumns = false,
                IsReadOnly = true,
                Margin = new Thickness(0),
                HeadersVisibility = DataGridHeadersVisibility.Column,
                GridLinesVisibility = DataGridGridLinesVisibility.Horizontal,
                CanUserAddRows = false,
                CanUserDeleteRows = false,
                SelectionMode = DataGridSelectionMode.Single,
                AlternationCount = 2
            };
            dayGrid.LoadingRow += OnDayGridLoadingRow;

            ConfigureDayGridColumns();

            rightPanel.Children.Add(tradeListHeader);
            Grid.SetRow(tradeListHeader, 0);

            rightPanel.Children.Add(dayGrid);
            Grid.SetRow(dayGrid, 1);

            rightCard.Child = rightPanel;

            detail.Children.Add(leftCard);
            Grid.SetColumn(leftCard, 0);

            detail.Children.Add(rightCard);
            Grid.SetColumn(rightCard, 1);

            Children.Add(detail);
            Grid.SetRow(detail, 5);
        }

        private void OnDayGridLoadingRow(object sender, DataGridRowEventArgs e)
        {
            double? pnl = null;
            MatchedTradeRow matchedTrade = e.Row.Item as MatchedTradeRow;
            if (matchedTrade != null)
                pnl = matchedTrade.Pnl;

            ExecutionBreakdownRow executionLeg = e.Row.Item as ExecutionBreakdownRow;
            if (executionLeg != null)
                pnl = executionLeg.Pnl;

            if (!pnl.HasValue)
            {
                e.Row.Background = Brushes.White;
                e.Row.Foreground = Brushes.Black;
                return;
            }

            if (pnl.Value > 0)
            {
                e.Row.Background = softPositiveBrush;
                e.Row.Foreground = Brushes.Black;
            }
            else if (pnl.Value < 0)
            {
                e.Row.Background = softNegativeBrush;
                e.Row.Foreground = Brushes.Black;
            }
            else
            {
                e.Row.Background = Brushes.WhiteSmoke;
                e.Row.Foreground = Brushes.Black;
            }
        }

        private void OnTradeListViewChanged(object sender, SelectionChangedEventArgs e)
        {
            if (isApplyingUiState || !IsLoaded || !IsUiReady())
                return;

            ConfigureDayGridColumns();
            SaveCurrentUiState();
            RefreshAll();
        }

        private void OnExportCurrentTradeListCsvClick(object sender, RoutedEventArgs e)
        {
            try
            {
                string viewMode = GetTradeListViewMode();
                List<MatchedTradeRow> matchedTrades = dayGrid.Items.Cast<object>().OfType<MatchedTradeRow>().ToList();
                List<ExecutionBreakdownRow> executionLegs = dayGrid.Items.Cast<object>().OfType<ExecutionBreakdownRow>().ToList();
                int rowCount = string.Equals(viewMode, ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase)
                    ? executionLegs.Count
                    : matchedTrades.Count;

                if (rowCount == 0)
                {
                    MessageBox.Show(
                        "No rows to export for the selected day and filters.",
                        "Trade Calendar Export",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                    Title = "Export current trade list to CSV",
                    FileName = BuildCurrentTradeListExportFileName(),
                    OverwritePrompt = true
                };

                bool? ok = dialog.ShowDialog();
                if (ok != true)
                    return;

                if (string.Equals(viewMode, ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase))
                    CurrentTradeListCsvExporter.ExportExecutionLegs(dialog.FileName, executionLegs);
                else
                    CurrentTradeListCsvExporter.ExportMatchedTrades(dialog.FileName, matchedTrades);

                MessageBox.Show(
                    "Exported " + rowCount.ToString(CultureInfo.InvariantCulture) + " rows to:" + Environment.NewLine + dialog.FileName,
                    "Trade Calendar Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                OpenExportFolder(dialog.FileName);
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar current trade list export error: " + ex, PrintTo.OutputTab1);
                MessageBox.Show(
                    "Export failed: " + ex.Message,
                    "Trade Calendar Export",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void OpenExportFolder(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    return;

                Process.Start(new ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "/select,\"" + Path.GetFullPath(filePath) + "\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar open export folder error: " + ex, PrintTo.OutputTab1);
            }
        }

        private string BuildCurrentTradeListExportFileName()
        {
            string instrument = GetComboSelection(instrumentComboBox, AllInstrumentsText);
            if (string.Equals(instrument, AllInstrumentsText, StringComparison.OrdinalIgnoreCase))
                instrument = "AllInstruments";

            string viewMode = string.Equals(GetTradeListViewMode(), ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase)
                ? "ExecutionLegs"
                : "MatchedTrades";

            return "TradeCalendar_"
                + CurrentTradeListCsvExporter.SanitizeFileNamePart(instrument)
                + "_"
                + selectedDate.ToString("yyyyMMdd", CultureInfo.InvariantCulture)
                + "_"
                + viewMode
                + ".csv";
        }

        private string GetTradeListViewMode()
        {
            string selected = GetComboSelection(tradeListViewComboBox, MatchedTradesViewText);
            return string.Equals(selected, ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase)
                ? ExecutionLegsViewText
                : MatchedTradesViewText;
        }

        private void ConfigureDayGridColumns()
        {
            if (dayGrid == null)
                return;

            dayGrid.Columns.Clear();

            if (string.Equals(GetTradeListViewMode(), ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase))
            {
                if (tradesHeaderText != null)
                    tradesHeaderText.Text = "Trades / execution breakdown";
                AddExecutionLegColumns();
                return;
            }

            if (tradesHeaderText != null)
                tradesHeaderText.Text = "Matched trades";
            AddMatchedTradeColumns();
        }

        private void AddMatchedTradeColumns()
        {
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Entry Time",
                Binding = new Binding("EntryTimeDisplay"),
                Width = new DataGridLength(150)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Exit Time",
                Binding = new Binding("ExitTimeDisplay"),
                Width = new DataGridLength(150)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Symbol",
                Binding = new Binding("Symbol"),
                Width = new DataGridLength(85)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Side",
                Binding = new Binding("Direction"),
                Width = new DataGridLength(70)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Qty",
                Binding = new Binding("Qty"),
                Width = new DataGridLength(55)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Entry Price",
                Binding = new Binding("EntryPriceDisplay"),
                Width = new DataGridLength(95)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Exit Price",
                Binding = new Binding("ExitPriceDisplay"),
                Width = new DataGridLength(95)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Commission",
                Binding = new Binding("CommissionDisplay"),
                Width = new DataGridLength(95)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "P&L",
                Binding = new Binding("PnlDisplay"),
                Width = new DataGridLength(90)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Source",
                Binding = new Binding("Source"),
                Width = new DataGridLength(85)
            });
        }

        private void AddExecutionLegColumns()
        {
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Time",
                Binding = new Binding("Time") { StringFormat = "HH:mm:ss" },
                Width = new DataGridLength(85)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Opened On",
                Binding = new Binding("OpenedOnDisplay"),
                Width = new DataGridLength(170)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Closed On",
                Binding = new Binding("ClosedOnDisplay"),
                Width = new DataGridLength(170)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Symbol",
                Binding = new Binding("Symbol"),
                Width = new DataGridLength(80)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Side",
                Binding = new Binding("Direction"),
                Width = new DataGridLength(85)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Action",
                Binding = new Binding("Action"),
                Width = new DataGridLength(85)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Qty",
                Binding = new Binding("Qty"),
                Width = new DataGridLength(55)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Price",
                Binding = new Binding("Price") { StringFormat = "N2" },
                Width = new DataGridLength(95)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Commission",
                Binding = new Binding("Commission") { StringFormat = "N2" },
                Width = new DataGridLength(95)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "P&L",
                Binding = new Binding("PnlDisplay"),
                Width = new DataGridLength(90)
            });
            dayGrid.Columns.Add(new DataGridTextColumn
            {
                Header = "Source",
                Binding = new Binding("Source"),
                Width = new DataGridLength(95)
            });
        }

        private Button MakeButton(string text)
        {
            return new Button
            {
                Content = text,
                Margin = new Thickness(0, 0, 6, 0),
                Padding = new Thickness(12, 5, 12, 5),
                MinWidth = 42,
                VerticalAlignment = VerticalAlignment.Center
            };
        }

        private bool IsUiReady()
        {
            return monthTitle != null
                && monthTotalText != null
                && accountCashValueText != null
                && calendarGrid != null
                && dayHeaderText != null
                && dayPnlText != null
                && instrumentSummaryPanel != null
                && tradeListViewComboBox != null
                && exportCurrentTradeListCsvButton != null
                && tradesHeaderText != null
                && dayGrid != null
                && activeFilterSummaryText != null
                && instrumentComboBox != null
                && sourceComboBox != null
                && sideComboBox != null
                && resultComboBox != null
                && fromDatePicker != null
                && toDatePicker != null
                && allAccountsCheckBox != null
                && accountSelector != null;
        }

        private void RefreshAll()
        {
            if (!IsUiReady())
                return;

            TradeQueryFilters filters = GetCurrentFilters();
            TradeAccountBook book = service.BuildAccountBook(filters.AccountName).ApplyFilters(filters);

            monthTitle.Text = currentMonth.ToString("MMMM yyyy", CultureInfo.InvariantCulture);
            monthTotalText.Text = FormatCurrency(book.GetMonthTotal(currentMonth));
            monthTotalText.Foreground = ChooseBrush(book.GetMonthTotal(currentMonth));
            accountCashValueText.Text = GetAccountCashValueDisplay();
            accountCashValueText.Foreground = neutralBrush;

            activeFilterSummaryText.Text = BuildFilterSummary(filters);

            RenderCalendar(book);
            RefreshDetail(book, filters);
            SaveCurrentUiState();
        }

        private TradeQueryFilters GetCurrentFilters()
        {
            string accountName = GetSelectedAccountName();

            return new TradeQueryFilters
            {
                AccountName = accountName,
                Instrument = GetComboSelection(instrumentComboBox, AllInstrumentsText),
                Source = GetComboSelection(sourceComboBox, AllSourcesText),
                Side = GetComboSelection(sideComboBox, AllSidesText),
                Result = GetComboSelection(resultComboBox, AllResultsText),
                FromDate = fromDatePicker != null ? fromDatePicker.SelectedDate : null,
                ToDate = toDatePicker != null ? toDatePicker.SelectedDate : null
            };
        }

        private string BuildFilterSummary(TradeQueryFilters filters)
        {
            List<string> parts = new List<string>();

            parts.Add(string.IsNullOrWhiteSpace(filters.AccountName) ? "All accounts" : "Account: " + filters.AccountName);

            if (filters.HasInstrumentFilter)
                parts.Add("Instrument: " + filters.Instrument);
            if (filters.HasSourceFilter)
                parts.Add("Source: " + filters.Source);
            if (filters.HasSideFilter)
                parts.Add("Side: " + filters.Side);
            if (filters.HasResultFilter)
                parts.Add("Result: " + filters.Result);
            if (filters.FromDate.HasValue || filters.ToDate.HasValue)
            {
                string fromText = filters.FromDate.HasValue ? filters.FromDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "...";
                string toText = filters.ToDate.HasValue ? filters.ToDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) : "...";
                parts.Add("Range: " + fromText + " to " + toText);
            }

            return "Filters: " + string.Join(" | ", parts);
        }

        private string GetSelectedAccountName()
        {
            if (allAccountsCheckBox != null && allAccountsCheckBox.IsChecked == true)
                return null;

            string selectedAccountName = GetCurrentlySelectedAccountName();
            if (!string.IsNullOrWhiteSpace(selectedAccountName))
                return selectedAccountName;

            Account rememberedAccount = FindAccountByName(lastSelectedAccountName);
            if (rememberedAccount != null)
                return rememberedAccount.Name;

            return null;
        }

        private string GetAccountCashValueDisplay()
        {
            if (allAccountsCheckBox != null && allAccountsCheckBox.IsChecked == true)
                return "All accounts";

            string accountName = GetSelectedAccountName();
            if (string.IsNullOrWhiteSpace(accountName))
                return "N/A";

            Account account = FindAccountByName(accountName);
            if (account == null)
                return "N/A";

            try
            {
                double cashValue = account.Get(AccountItem.CashValue, Currency.UsDollar);
                if (double.IsNaN(cashValue) || double.IsInfinity(cashValue))
                    return "N/A";

                return FormatCurrencyWithCents(cashValue);
            }
            catch
            {
                return "N/A";
            }
        }

        private string GetSavedAccountName(TradeCalendarUiState state)
        {
            if (state == null)
                return null;

            string savedName = !string.IsNullOrWhiteSpace(state.LastSelectedAccountName)
                ? state.LastSelectedAccountName
                : state.AccountName;

            return string.IsNullOrWhiteSpace(savedName) ? null : savedName.Trim();
        }

        private bool TryRestorePendingAccountSelection()
        {
            if (string.IsNullOrWhiteSpace(pendingRestoreAccountName))
                return true;

            string requestedAccountName = pendingRestoreAccountName;
            if (!TrySelectAccountByNameSimple(requestedAccountName))
                return false;

            lastSelectedAccountName = GetCurrentlySelectedAccountName() ?? requestedAccountName;
            pendingRestoreAccountName = null;
            isStartupAccountRestorePending = false;
            StopAccountRestoreTimer();
            SaveCurrentUiState();
            RefreshFilterChoices();
            RefreshAll();
            NinjaTrader.Code.Output.Process("TradeCalendar restored account='" + lastSelectedAccountName + "'", PrintTo.OutputTab1);
            return true;
        }

        private void StartAccountRestoreTimer()
        {
            if (!isStartupAccountRestorePending || string.IsNullOrWhiteSpace(pendingRestoreAccountName))
                return;

            if (accountRestoreTimer == null)
            {
                accountRestoreTimer = new DispatcherTimer();
                accountRestoreTimer.Interval = AccountRestoreInterval;
                accountRestoreTimer.Tick += OnAccountRestoreTimerTick;
            }

            accountRestoreAttempts = 0;
            accountRestoreTimer.Start();
        }

        private void StopAccountRestoreTimer()
        {
            if (accountRestoreTimer != null && accountRestoreTimer.IsEnabled)
                accountRestoreTimer.Stop();
        }

        private void OnAccountRestoreTimerTick(object sender, EventArgs e)
        {
            accountRestoreAttempts++;
            if (TryRestorePendingAccountSelection())
                return;

            if (accountRestoreAttempts < AccountRestoreMaxAttempts)
                return;

            StopAccountRestoreTimer();
            isStartupAccountRestorePending = false;
            pendingRestoreAccountName = null;
            SelectFallbackAccountOrAllAccounts();
            SaveCurrentUiState();
            RefreshFilterChoices();
            RefreshAll();
            NinjaTrader.Code.Output.Process(
                "TradeCalendar account restore failed after retries; fallback to All accounts.",
                PrintTo.OutputTab1);
        }

        private string GetCurrentlySelectedAccountName()
        {
            if (accountSelector == null)
                return null;

            Account selectedItemAccount = accountSelector.SelectedItem as Account;
            if (selectedItemAccount != null && !string.IsNullOrWhiteSpace(selectedItemAccount.Name))
                return selectedItemAccount.Name;

            string selectedItemName = GetAccountNameFromItem(accountSelector.SelectedItem);
            Account knownAccount = FindAccountByName(selectedItemName);
            if (knownAccount != null)
                return knownAccount.Name;

            if (accountSelector.SelectedAccount != null && !string.IsNullOrWhiteSpace(accountSelector.SelectedAccount.Name))
                return accountSelector.SelectedAccount.Name;

            return null;
        }

        private bool TrySelectAccountByNameSimple(string accountName)
        {
            if (string.IsNullOrWhiteSpace(accountName) || accountSelector == null)
                return false;

            string target = accountName.Trim();
            object matchingItem = null;
            Account matchingAccount = null;
            int matchingIndex = -1;

            for (int i = 0; i < accountSelector.Items.Count; i++)
            {
                object item = accountSelector.Items[i];
                string itemName = GetAccountNameFromItem(item);
                if (!string.Equals(itemName, target, StringComparison.OrdinalIgnoreCase))
                    continue;

                matchingItem = item;
                matchingAccount = item as Account;
                matchingIndex = i;
                break;
            }

            if (matchingAccount == null)
                matchingAccount = FindAccountByName(target);

            if (matchingItem == null && matchingAccount == null)
                return false;

            bool previousApplyingAccountSelection = isApplyingAccountSelection;
            isApplyingAccountSelection = true;
            try
            {
                if (matchingItem != null)
                {
                    accountSelector.SelectedIndex = matchingIndex;
                    accountSelector.SelectedItem = matchingItem;
                }
                if (matchingAccount != null)
                    accountSelector.SelectedAccount = matchingAccount;
            }
            finally
            {
                isApplyingAccountSelection = previousApplyingAccountSelection;
            }

            if (matchingItem == null)
                return false;

            string selectedName = GetCurrentlySelectedAccountName();
            string expectedName = matchingAccount != null ? matchingAccount.Name : GetAccountNameFromItem(matchingItem);
            return accountSelector.SelectedIndex == matchingIndex
                && !string.IsNullOrWhiteSpace(selectedName)
                && string.Equals(selectedName, expectedName, StringComparison.OrdinalIgnoreCase);
        }

        private string GetAccountNameFromItem(object item)
        {
            if (item == null)
                return null;

            Account account = item as Account;
            if (account != null)
                return account.Name;

            string itemText = item.ToString();
            return string.IsNullOrWhiteSpace(itemText) ? null : itemText.Trim();
        }

        private bool TrySelectFirstAvailableAccount()
        {
            if (accountSelector == null)
                return false;

            foreach (object item in accountSelector.Items)
            {
                string itemName = GetAccountNameFromItem(item);
                if (TrySelectAccountByNameSimple(itemName))
                {
                    lastSelectedAccountName = GetCurrentlySelectedAccountName();
                    return true;
                }
            }

            List<string> accountNames;
            lock (Account.All)
            {
                accountNames = Account.All
                    .Where(account => account != null && !string.IsNullOrWhiteSpace(account.Name))
                    .Select(account => account.Name)
                    .ToList();
            }

            foreach (string accountName in accountNames)
            {
                if (TrySelectAccountByNameSimple(accountName))
                {
                    lastSelectedAccountName = GetCurrentlySelectedAccountName();
                    return true;
                }
            }

            return false;
        }

        private void SelectFallbackAccountOrAllAccounts()
        {
            NinjaTrader.Code.Output.Process(
                "TradeCalendar fallback to All accounts; account not available='" + (lastSelectedAccountName ?? string.Empty) + "'",
                PrintTo.OutputTab1);

            bool previousApplyingAccountSelection = isApplyingAccountSelection;
            isApplyingAccountSelection = true;
            try
            {
                if (allAccountsCheckBox != null)
                    allAccountsCheckBox.IsChecked = true;
                if (accountSelector != null)
                    accountSelector.IsEnabled = false;
            }
            finally
            {
                isApplyingAccountSelection = previousApplyingAccountSelection;
            }
        }

        private void RenderCalendar(TradeAccountBook book)
        {
            calendarGrid.Children.Clear();

            DateTime first = new DateTime(currentMonth.Year, currentMonth.Month, 1);
            DateTime start = first.AddDays(-(int)first.DayOfWeek);

            for (int i = 0; i < 42; i++)
            {
                DateTime day = start.AddDays(i);
                DateTime cellDay = day.Date;
                int row = i / 7;
                int col = i % 7;

                Border border = new Border
                {
                    BorderBrush = new SolidColorBrush(Color.FromRgb(210, 210, 210)),
                    BorderThickness = new Thickness(1),
                    Background = cellDay == selectedDate.Date
                        ? new SolidColorBrush(Color.FromRgb(238, 245, 255))
                        : Brushes.White,
                    Padding = new Thickness(6),
                    Margin = new Thickness(0)
                };

                if (cellDay == selectedDate.Date)
                {
                    border.BorderBrush = selectedBorderBrush;
                    border.BorderThickness = new Thickness(2);
                }

                Grid cell = new Grid();
                cell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                cell.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
                cell.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                border.Child = cell;

                TextBlock dateText = new TextBlock
                {
                    Text = cellDay.Day.ToString(CultureInfo.InvariantCulture),
                    FontSize = 16,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Top,
                    Foreground = cellDay.Month == currentMonth.Month ? Brushes.Black : new SolidColorBrush(Color.FromRgb(165, 165, 165)),
                    Cursor = Cursors.Hand
                };
                dateText.MouseLeftButtonUp += (s, e) =>
                {
                    selectedDate = cellDay;
                    RefreshAll();
                };
                Grid.SetRow(dateText, 0);
                cell.Children.Add(dateText);

                bool hasOpenOrCloseActivity = book.DayHasAnyActivity(cellDay);
                double dayPnl = book.GetDayTotal(cellDay);

                if (hasOpenOrCloseActivity || Math.Abs(dayPnl) > 0.000001)
                {
                    Border pnlBadge = new Border
                    {
                        Margin = new Thickness(0, 6, 0, 0),
                        Padding = new Thickness(4, 2, 4, 2),
                        Height = 22,
                        HorizontalAlignment = HorizontalAlignment.Stretch,
                        VerticalAlignment = VerticalAlignment.Bottom,
                        CornerRadius = new CornerRadius(2),
                        Background = dayPnl > 0 ? softPositiveBrush : (dayPnl < 0 ? softNegativeBrush : Brushes.WhiteSmoke),
                        BorderBrush = dayPnl > 0 ? positiveBrush : (dayPnl < 0 ? negativeBrush : neutralBrush),
                        BorderThickness = new Thickness(1),
                        Cursor = Cursors.Hand,
                        SnapsToDevicePixels = true,
                        ClipToBounds = true
                    };

                    Viewbox pnlViewbox = new Viewbox
                    {
                        Stretch = Stretch.Uniform,
                        StretchDirection = StretchDirection.DownOnly,
                        Height = 14,
                        Child = new TextBlock
                        {
                            Text = "PnL: " + FormatCurrency(dayPnl),
                            FontWeight = FontWeights.Medium,
                            TextAlignment = TextAlignment.Center,
                            Foreground = dayPnl > 0 ? positiveBrush : (dayPnl < 0 ? negativeBrush : neutralBrush)
                        }
                    };

                    pnlBadge.Child = pnlViewbox;
                    pnlBadge.MouseLeftButtonUp += (s, e) =>
                    {
                        selectedDate = cellDay;
                        RefreshAll();
                    };
                    Grid.SetRow(pnlBadge, 2);
                    cell.Children.Add(pnlBadge);
                }

                if (cellDay.DayOfWeek == DayOfWeek.Saturday)
                {
                    double weekTotal = book.GetWeekTotalForSaturday(cellDay);
                    Border weekBadge = new Border
                    {
                        Background = weekTotal > 0 ? positiveBrush : (weekTotal < 0 ? negativeBrush : neutralBrush),
                        CornerRadius = new CornerRadius(3),
                        Margin = new Thickness(0, 8, 0, 0),
                        Padding = new Thickness(6, 3, 6, 3)
                    };

                    TextBlock weekText = new TextBlock
                    {
                        Text = "Week: " + FormatCurrency(weekTotal),
                        Foreground = Brushes.White,
                        FontWeight = FontWeights.SemiBold
                    };

                    weekBadge.Child = weekText;
                    Grid.SetRow(weekBadge, 2);
                    cell.Children.Add(weekBadge);
                }

                if (cellDay.Month != currentMonth.Month)
                    border.Opacity = 0.72;

                Grid.SetRow(border, row);
                Grid.SetColumn(border, col);
                calendarGrid.Children.Add(border);
            }
        }

        private void RefreshDetail(TradeAccountBook book, TradeQueryFilters filters)
        {
            DayBook day = book.GetDayBook(selectedDate);

            dayHeaderText.Text = "Trading Day - " + selectedDate.ToString("ddd, MMM d, yyyy", CultureInfo.InvariantCulture);
            dayPnlText.Text = FormatCurrencyWithCents(day.TotalPnl);
            dayPnlText.Foreground = ChooseBrush(day.TotalPnl);

            instrumentSummaryPanel.Children.Clear();
            if (day.InstrumentPnl.Count == 0)
            {
                instrumentSummaryPanel.Children.Add(new TextBlock
                {
                    Text = "No realized PnL for this day under the current filters.",
                    Foreground = neutralBrush
                });
            }
            else
            {
                foreach (KeyValuePair<string, double> kv in day.InstrumentPnl.OrderByDescending(k => Math.Abs(k.Value)))
                {
                    Border rowBorder = new Border
                    {
                        BorderBrush = new SolidColorBrush(Color.FromRgb(220, 220, 220)),
                        BorderThickness = new Thickness(1),
                        Background = kv.Value > 0 ? softPositiveBrush : (kv.Value < 0 ? softNegativeBrush : Brushes.WhiteSmoke),
                        CornerRadius = new CornerRadius(3),
                        Padding = new Thickness(8, 4, 8, 4),
                        Margin = new Thickness(0, 0, 0, 6)
                    };

                    DockPanel row = new DockPanel();

                    TextBlock symbol = new TextBlock
                    {
                        Text = kv.Key,
                        FontWeight = FontWeights.SemiBold
                    };

                    TextBlock pnl = new TextBlock
                    {
                        Text = FormatCurrencyWithCents(kv.Value),
                        Foreground = ChooseBrush(kv.Value),
                        HorizontalAlignment = HorizontalAlignment.Right
                    };

                    DockPanel.SetDock(pnl, Dock.Right);
                    row.Children.Add(pnl);
                    row.Children.Add(symbol);

                    rowBorder.Child = row;
                    instrumentSummaryPanel.Children.Add(rowBorder);
                }
            }

            if (string.Equals(GetTradeListViewMode(), ExecutionLegsViewText, StringComparison.OrdinalIgnoreCase))
            {
                dayGrid.ItemsSource = day.Rows.OrderByDescending(r => r.Time).ThenBy(r => r.Symbol).ToList();
                return;
            }

            List<MatchedTradeRow> matchedTrades = service.BuildMatchedTradeRows(filters.AccountName)
                .Where(filters.MatchesMatchedTrade)
                .Where(r => TradeSessionCalendar_v3.ToTradingDate(r.ExitTime) == selectedDate.Date)
                .OrderByDescending(r => r.ExitTime)
                .ThenBy(r => r.Symbol)
                .ToList();
            dayGrid.ItemsSource = matchedTrades;
        }

        private Brush ChooseBrush(double value)
        {
            if (value > 0)
                return positiveBrush;
            if (value < 0)
                return negativeBrush;
            return neutralBrush;
        }

        private string FormatCurrency(double value)
        {
            return string.Format(CultureInfo.InvariantCulture, "${0:N0}", value);
        }

        private string FormatCurrencyWithCents(double value)
        {
            return string.Format(CultureInfo.InvariantCulture, "${0:N2}", value);
        }
    }

    public sealed class TradeLedgerService_v3
    {
        private static readonly Lazy<TradeLedgerService_v3> lazy = new Lazy<TradeLedgerService_v3>(() => new TradeLedgerService_v3());
        public static TradeLedgerService_v3 Instance { get { return lazy.Value; } }

        private readonly object sync = new object();
        private readonly Dictionary<string, Account> subscribedAccounts = new Dictionary<string, Account>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, PersistedExecutionRecord> ledgerByKey = new Dictionary<string, PersistedExecutionRecord>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, ImportedTradeRecord> importedTradesByKey = new Dictionary<string, ImportedTradeRecord>(StringComparer.OrdinalIgnoreCase);

        private bool initialized;
        private bool isSavingLive;
        private bool isSavingImported;

        public event EventHandler DataChanged;

        private TradeLedgerService_v3()
        {
        }

        private string LiveLedgerPath
        {
            get { return Path.Combine(NinjaTrader.Core.Globals.UserDataDir, "TradeCalendarLedger.xml"); }
        }

        private string ImportedLedgerPath
        {
            get { return Path.Combine(NinjaTrader.Core.Globals.UserDataDir, "TradeCalendarImportedTrades.xml"); }
        }

        private static readonly DateTime MinimumTradeDate = new DateTime(2026, 4, 8, 0, 0, 0, DateTimeKind.Unspecified);

        public void EnsureInitialized()
        {
            lock (sync)
            {
                if (initialized)
                    return;

                LoadLiveLedger_NoLock();
                LoadImportedLedger_NoLock();
                PruneRecordsBeforeCutoff_NoLock();
                PruneDuplicateLiveRecords_NoLock();
                SubscribeAllKnownAccounts_NoLock();
                SnapshotCurrentSessionExecutions_NoLock();
                PruneDuplicateLiveRecords_NoLock();
                PruneRecordsBeforeCutoff_NoLock();
                PruneDuplicateLiveRecords_NoLock();
                initialized = true;
            }

            RaiseDataChanged();
        }

        public void Shutdown()
        {
            lock (sync)
            {
                foreach (KeyValuePair<string, Account> kv in subscribedAccounts.ToList())
                {
                    try
                    {
                        kv.Value.ExecutionUpdate -= OnExecutionUpdate;
                    }
                    catch
                    {
                    }
                }

                subscribedAccounts.Clear();
                SaveLiveLedger_NoLock();
                SaveImportedLedger_NoLock();
                initialized = false;
            }
        }

        public void RefreshCurrentSessionSnapshots()
        {
            if (!initialized)
                EnsureInitialized();

            lock (sync)
            {
                SubscribeAllKnownAccounts_NoLock();
                SnapshotCurrentSessionExecutions_NoLock();
                PruneDuplicateLiveRecords_NoLock();
            }

            RaiseDataChanged();
        }

        public CsvImportResult ImportCsvTrades(IEnumerable<string> filePaths, string defaultAccountName)
        {
            CsvImportResult result = new CsvImportResult();

            if (filePaths == null)
            {
                result.Errors.Add("No files selected.");
                return result;
            }

            EnsureInitialized();

            lock (sync)
            {
                foreach (string path in filePaths)
                {
                    if (string.IsNullOrWhiteSpace(path))
                        continue;

                    result.FilesProcessed++;

                    CsvParseResult parsed = CsvTradeImportParser.ParseFile(path);
                    result.RowsRead += parsed.RowsRead;
                    result.RowsSkipped += parsed.RowsSkipped;
                    result.Errors.AddRange(parsed.Errors);

                    foreach (ImportedTradeRecord trade in parsed.Trades)
                    {
                        if (trade == null)
                            continue;

                        if (string.IsNullOrWhiteSpace(trade.AccountName))
                            trade.AccountName = defaultAccountName ?? string.Empty;

                        if (string.IsNullOrWhiteSpace(trade.Action))
                            trade.Action = "TRADE";

                        if (!IsImportedTradeOnOrAfterCutoff(trade))
                        {
                            result.RowsSkipped++;
                            continue;
                        }

                        trade.Key = BuildImportedTradeKey(trade);

                        if (importedTradesByKey.ContainsKey(trade.Key))
                        {
                            result.RowsDuplicate++;
                            continue;
                        }

                        importedTradesByKey[trade.Key] = trade;
                        result.RowsImported++;
                    }
                }

                SaveImportedLedger_NoLock();
            }

            RaiseDataChanged();
            return result;
        }

        public TradeAccountBook BuildAccountBook(string accountName)
        {
            EnsureInitialized();

            List<PersistedExecutionRecord> liveRecords;
            List<ImportedTradeRecord> importedRecords;

            lock (sync)
            {
                liveRecords = ledgerByKey.Values
                    .Where(r => IsLiveRecordOnOrAfterCutoff(r)
                        && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                    .OrderBy(r => r.Time)
                    .ThenBy(r => r.ExecutionId)
                    .ThenBy(r => r.OrderId)
                    .ToList();

                importedRecords = importedTradesByKey.Values
                    .Where(r => IsImportedTradeOnOrAfterCutoff(r)
                        && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                    .OrderBy(r => r.ClosedOn)
                    .ThenBy(r => r.Symbol)
                    .ToList();
            }

            return TradeAccountBookBuilder_v3.Build(liveRecords, importedRecords);
        }

        public List<MatchedTradeRow> BuildMatchedTradeRows(string accountName)
        {
            EnsureInitialized();

            List<PersistedExecutionRecord> liveRecords;
            List<ImportedTradeRecord> importedRecords;

            lock (sync)
            {
                liveRecords = ledgerByKey.Values
                    .Where(r => IsLiveRecordOnOrAfterCutoff(r)
                        && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                    .OrderBy(r => r.AccountName)
                    .ThenBy(r => r.Instrument)
                    .ThenBy(r => r.Time)
                    .ThenBy(r => r.ExecutionId)
                    .ToList();

                importedRecords = importedTradesByKey.Values
                    .Where(r => IsImportedTradeOnOrAfterCutoff(r)
                        && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                    .OrderBy(r => r.AccountName)
                    .ThenBy(r => r.Symbol)
                    .ThenBy(r => r.ClosedOn)
                    .ToList();
            }

            MatchedTradeBuildResult result = TradeAccountBookBuilder_v3.BuildMatchedTradeRows(liveRecords, importedRecords);
            NinjaTrader.Code.Output.Process(
                "TradeCalendar matched trades built: " + result.Rows.Count.ToString(CultureInfo.InvariantCulture)
                + "; ID match rows: " + result.IdMatchRows.ToString(CultureInfo.InvariantCulture)
                + "; fallback match rows: " + result.FallbackMatchRows.ToString(CultureInfo.InvariantCulture)
                + "; unmatched executions: " + result.UnmatchedExecutions.ToString(CultureInfo.InvariantCulture),
                PrintTo.OutputTab1);
            return result.Rows;
        }

        public List<string> GetAvailableInstruments(string accountName)
        {
            EnsureInitialized();

            lock (sync)
            {
                return ledgerByKey.Values
                    .Where(r => IsLiveRecordOnOrAfterCutoff(r)
                        && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                    .Select(r => r.Instrument)
                    .Concat(importedTradesByKey.Values
                        .Where(r => IsImportedTradeOnOrAfterCutoff(r)
                            && (string.IsNullOrEmpty(accountName) || string.Equals(r.AccountName, accountName, StringComparison.OrdinalIgnoreCase)))
                        .Select(r => r.Symbol))
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(s => s)
                    .ToList();
            }
        }

        private void SubscribeAllKnownAccounts_NoLock()
        {
            lock (Account.All)
            {
                foreach (Account account in Account.All)
                {
                    if (account == null || string.IsNullOrWhiteSpace(account.Name))
                        continue;

                    if (subscribedAccounts.ContainsKey(account.Name))
                        continue;

                    account.ExecutionUpdate += OnExecutionUpdate;
                    subscribedAccounts[account.Name] = account;
                }
            }
        }

        private void SnapshotCurrentSessionExecutions_NoLock()
        {
            foreach (Account account in subscribedAccounts.Values)
            {
                try
                {
                    lock (account.Executions)
                    {
                        foreach (Execution execution in account.Executions)
                            UpsertExecution_NoLock(account, execution);
                    }
                }
                catch (Exception ex)
                {
                    NinjaTrader.Code.Output.Process("TradeCalendar snapshot error: " + ex, PrintTo.OutputTab1);
                }
            }

            SaveLiveLedger_NoLock();
        }

        private void OnExecutionUpdate(object sender, ExecutionEventArgs e)
        {
            try
            {
                Account account = sender as Account;
                if (account == null || e == null || e.Execution == null)
                    return;

                lock (sync)
                {
                    UpsertExecution_NoLock(account, e.Execution);
                    PruneDuplicateLiveRecords_NoLock();
                    SaveLiveLedger_NoLock();
                }

                RaiseDataChanged();
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar execution update error: " + ex, PrintTo.OutputTab1);
            }
        }

        private void UpsertExecution_NoLock(Account account, Execution execution)
        {
            if (execution == null || account == null || execution.Instrument == null)
                return;

            if (!IsOnOrAfterCutoff(execution.Time))
                return;

            string action = GetOrderActionText(execution);
            string key = BuildExecutionKey(account.Name, execution);

            PersistedExecutionRecord record = new PersistedExecutionRecord
            {
                Key = key,
                AccountName = account.Name ?? string.Empty,
                Instrument = SafeInstrumentName(execution),
                Action = action,
                Quantity = execution.Quantity,
                Price = execution.Price,
                Commission = execution.Commission,
                Time = execution.Time,
                OrderId = execution.OrderId ?? string.Empty,
                ExecutionId = execution.ExecutionId ?? string.Empty,
                MarketPosition = execution.MarketPosition.ToString(),
                OrderName = execution.Name ?? string.Empty,
                PointValue = execution.Instrument.MasterInstrument != null ? execution.Instrument.MasterInstrument.PointValue : 1d
            };

            ledgerByKey[key] = record;
        }

        private string SafeInstrumentName(Execution execution)
        {
            if (execution == null || execution.Instrument == null)
                return string.Empty;

            return !string.IsNullOrWhiteSpace(execution.Instrument.MasterInstrument != null ? execution.Instrument.MasterInstrument.Name : null)
                ? execution.Instrument.MasterInstrument.Name
                : execution.Instrument.FullName;
        }

        private string GetOrderActionText(Execution execution)
        {
            try
            {
                if (execution.Order != null)
                    return execution.Order.OrderAction.ToString();
            }
            catch
            {
            }

            switch (execution.MarketPosition)
            {
                case MarketPosition.Long:
                    return "Buy";
                case MarketPosition.Short:
                    return "SellShort";
                default:
                    return "Unknown";
            }
        }

        private string BuildExecutionKey(string accountName, Execution execution)
        {
            if (execution == null)
                return string.Empty;

            string execId = execution.ExecutionId ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(execId))
                return string.Join("|",
                    accountName ?? string.Empty,
                    execId);

            return string.Join("|",
                accountName ?? string.Empty,
                execution.OrderId ?? string.Empty,
                execution.Time.Ticks.ToString(CultureInfo.InvariantCulture),
                execution.Quantity.ToString(CultureInfo.InvariantCulture),
                execution.Price.ToString("R", CultureInfo.InvariantCulture),
                execution.Name ?? string.Empty);
        }

        private string BuildImportedTradeKey(ImportedTradeRecord trade)
        {
            return string.Join("|",
                trade.AccountName ?? string.Empty,
                trade.Symbol ?? string.Empty,
                trade.Action ?? string.Empty,
                trade.OpenedOn.Ticks.ToString(CultureInfo.InvariantCulture),
                trade.ClosedOn.Ticks.ToString(CultureInfo.InvariantCulture),
                trade.Quantity.ToString(CultureInfo.InvariantCulture),
                trade.Price.ToString("R", CultureInfo.InvariantCulture),
                trade.Commission.ToString("R", CultureInfo.InvariantCulture),
                trade.Pnl.ToString("R", CultureInfo.InvariantCulture));
        }

        private string BuildLiveCanonicalKey(PersistedExecutionRecord record)
        {
            if (record == null)
                return string.Empty;

            if (!string.IsNullOrWhiteSpace(record.ExecutionId))
                return string.Join("|",
                    record.AccountName ?? string.Empty,
                    record.ExecutionId ?? string.Empty);

            return string.Join("|",
                record.AccountName ?? string.Empty,
                record.Instrument ?? string.Empty,
                (record.Action ?? string.Empty).Trim().ToUpperInvariant(),
                record.Quantity.ToString(CultureInfo.InvariantCulture),
                Math.Round(record.Price, 8).ToString("R", CultureInfo.InvariantCulture),
                Math.Round(record.Commission, 8).ToString("R", CultureInfo.InvariantCulture),
                record.Time.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture),
                record.MarketPosition ?? string.Empty,
                Math.Round(record.PointValue, 8).ToString("R", CultureInfo.InvariantCulture));
        }

        private static int GetActionPreferenceScore(string action)
        {
            string normalized = (action ?? string.Empty).Trim().ToUpperInvariant();
            switch (normalized)
            {
                case "BUY":
                case "SELL":
                    return 4;
                case "BUYTOCOVER":
                case "SELLSHORT":
                    return 3;
                case "UNKNOWN":
                    return 1;
                default:
                    return 2;
            }
        }

        private PersistedExecutionRecord SelectBestDuplicateRecord(IEnumerable<PersistedExecutionRecord> records)
        {
            return records
                .Where(x => x != null)
                .OrderByDescending(x => GetActionPreferenceScore(x.Action))
                .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.ExecutionId))
                .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.OrderId))
                .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.OrderName))
                .ThenByDescending(x => x.Time)
                .ThenBy(x => x.Key)
                .FirstOrDefault();
        }

        private void PruneDuplicateLiveRecords_NoLock()
        {
            List<IGrouping<string, PersistedExecutionRecord>> groups = ledgerByKey.Values
                .Where(r => r != null)
                .GroupBy(BuildLiveCanonicalKey, StringComparer.OrdinalIgnoreCase)
                .Where(g => !string.IsNullOrWhiteSpace(g.Key) && g.Count() > 1)
                .ToList();

            if (groups.Count == 0)
                return;

            foreach (IGrouping<string, PersistedExecutionRecord> group in groups)
            {
                PersistedExecutionRecord keep = SelectBestDuplicateRecord(group);
                if (keep == null)
                    continue;

                foreach (PersistedExecutionRecord duplicate in group)
                {
                    if (duplicate == null || duplicate.Key == keep.Key)
                        continue;

                    ledgerByKey.Remove(duplicate.Key);
                }
            }

            SaveLiveLedger_NoLock();
        }

        private static bool IsOnOrAfterCutoff(DateTime value)
        {
            return TradeSessionCalendar_v3.ToTradingDate(value) >= MinimumTradeDate.Date;
        }

        private static bool IsLiveRecordOnOrAfterCutoff(PersistedExecutionRecord record)
        {
            return record != null && IsOnOrAfterCutoff(record.Time);
        }

        private static bool IsImportedTradeOnOrAfterCutoff(ImportedTradeRecord trade)
        {
            return trade != null
                && (IsOnOrAfterCutoff(trade.OpenedOn)
                    || IsOnOrAfterCutoff(trade.ClosedOn));
        }

        private void PruneRecordsBeforeCutoff_NoLock()
        {
            List<string> liveKeysToRemove = ledgerByKey.Values
                .Where(r => !IsLiveRecordOnOrAfterCutoff(r))
                .Select(r => r.Key)
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToList();

            foreach (string key in liveKeysToRemove)
                ledgerByKey.Remove(key);

            List<string> importedKeysToRemove = importedTradesByKey.Values
                .Where(r => !IsImportedTradeOnOrAfterCutoff(r))
                .Select(r => r.Key)
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToList();

            foreach (string key in importedKeysToRemove)
                importedTradesByKey.Remove(key);

            if (liveKeysToRemove.Count > 0)
                SaveLiveLedger_NoLock();
            if (importedKeysToRemove.Count > 0)
                SaveImportedLedger_NoLock();
        }

        private void LoadLiveLedger_NoLock()
        {
            try
            {
                if (!File.Exists(LiveLedgerPath))
                    return;

                XmlSerializer serializer = new XmlSerializer(typeof(List<PersistedExecutionRecord>));
                using (FileStream stream = File.OpenRead(LiveLedgerPath))
                {
                    List<PersistedExecutionRecord> list = serializer.Deserialize(stream) as List<PersistedExecutionRecord>;
                    ledgerByKey.Clear();

                    if (list != null)
                    {
                        foreach (PersistedExecutionRecord item in list)
                        {
                            if (item == null || string.IsNullOrWhiteSpace(item.Key) || !IsLiveRecordOnOrAfterCutoff(item))
                                continue;

                            ledgerByKey[item.Key] = item;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar live load error: " + ex, PrintTo.OutputTab1);
            }
        }

        private void LoadImportedLedger_NoLock()
        {
            try
            {
                if (!File.Exists(ImportedLedgerPath))
                    return;

                XmlSerializer serializer = new XmlSerializer(typeof(List<ImportedTradeRecord>));
                using (FileStream stream = File.OpenRead(ImportedLedgerPath))
                {
                    List<ImportedTradeRecord> list = serializer.Deserialize(stream) as List<ImportedTradeRecord>;
                    importedTradesByKey.Clear();

                    if (list != null)
                    {
                        foreach (ImportedTradeRecord item in list)
                        {
                            if (item == null || string.IsNullOrWhiteSpace(item.Key) || !IsImportedTradeOnOrAfterCutoff(item))
                                continue;

                            importedTradesByKey[item.Key] = item;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar imported load error: " + ex, PrintTo.OutputTab1);
            }
        }

        private void SaveLiveLedger_NoLock()
        {
            if (isSavingLive)
                return;

            try
            {
                isSavingLive = true;

                string directory = Path.GetDirectoryName(LiveLedgerPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                List<PersistedExecutionRecord> list = ledgerByKey.Values.OrderBy(r => r.Time).ThenBy(r => r.ExecutionId).ToList();
                XmlSerializer serializer = new XmlSerializer(typeof(List<PersistedExecutionRecord>));
                using (FileStream stream = File.Create(LiveLedgerPath))
                    serializer.Serialize(stream, list);
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar live save error: " + ex, PrintTo.OutputTab1);
            }
            finally
            {
                isSavingLive = false;
            }
        }

        private void SaveImportedLedger_NoLock()
        {
            if (isSavingImported)
                return;

            try
            {
                isSavingImported = true;

                string directory = Path.GetDirectoryName(ImportedLedgerPath);
                if (!Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                List<ImportedTradeRecord> list = importedTradesByKey.Values.OrderBy(r => r.ClosedOn).ThenBy(r => r.Symbol).ToList();
                XmlSerializer serializer = new XmlSerializer(typeof(List<ImportedTradeRecord>));
                using (FileStream stream = File.Create(ImportedLedgerPath))
                    serializer.Serialize(stream, list);
            }
            catch (Exception ex)
            {
                NinjaTrader.Code.Output.Process("TradeCalendar imported save error: " + ex, PrintTo.OutputTab1);
            }
            finally
            {
                isSavingImported = false;
            }
        }


        public LedgerExportResult ExportLedgersToCsv(string filePath)
        {
            LedgerExportResult result = new LedgerExportResult();

            if (string.IsNullOrWhiteSpace(filePath))
            {
                result.Errors.Add("No output file was selected.");
                return result;
            }

            EnsureInitialized();

            List<PersistedExecutionRecord> liveRecords;
            List<ImportedTradeRecord> importedRecords;

            lock (sync)
            {
                liveRecords = ledgerByKey.Values
                    .Where(r => IsLiveRecordOnOrAfterCutoff(r))
                    .OrderBy(r => r.Time)
                    .ThenBy(r => r.ExecutionId)
                    .ThenBy(r => r.OrderId)
                    .ToList();

                importedRecords = importedTradesByKey.Values
                    .Where(r => IsImportedTradeOnOrAfterCutoff(r))
                    .OrderBy(r => r.ClosedOn)
                    .ThenBy(r => r.Symbol)
                    .ThenBy(r => r.AccountName)
                    .ToList();
            }

            try
            {
                string directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                    Directory.CreateDirectory(directory);

                using (StreamWriter writer = new StreamWriter(filePath, false, new UTF8Encoding(true)))
                {
                    writer.WriteLine("RowType,Source,Key,Account,Instrument,Action,Qty,Time,OpenedOn,ClosedOn,Price,Commission,PnL,OrderId,ExecutionId,MarketPosition,OrderName,PointValue,SourceFile");

                    foreach (PersistedExecutionRecord record in liveRecords)
                    {
                        writer.WriteLine(string.Join(",",
                            Csv("LiveExecution"),
                            Csv("Live"),
                            Csv(record.Key),
                            Csv(record.AccountName),
                            Csv(record.Instrument),
                            Csv(record.Action),
                            record.Quantity.ToString(CultureInfo.InvariantCulture),
                            Csv(FormatDateTime(record.Time)),
                            Csv(string.Empty),
                            Csv(string.Empty),
                            record.Price.ToString(CultureInfo.InvariantCulture),
                            record.Commission.ToString(CultureInfo.InvariantCulture),
                            Csv(string.Empty),
                            Csv(record.OrderId),
                            Csv(record.ExecutionId),
                            Csv(record.MarketPosition),
                            Csv(record.OrderName),
                            record.PointValue.ToString(CultureInfo.InvariantCulture),
                            Csv(string.Empty)));
                        result.LiveRowsExported++;
                    }

                    foreach (ImportedTradeRecord trade in importedRecords)
                    {
                        writer.WriteLine(string.Join(",",
                            Csv("ImportedTrade"),
                            Csv("Imported"),
                            Csv(trade.Key),
                            Csv(trade.AccountName),
                            Csv(trade.Symbol),
                            Csv(trade.Action),
                            trade.Quantity.ToString(CultureInfo.InvariantCulture),
                            Csv(string.Empty),
                            Csv(FormatDateTime(trade.OpenedOn)),
                            Csv(FormatDateTime(trade.ClosedOn)),
                            trade.Price.ToString(CultureInfo.InvariantCulture),
                            trade.Commission.ToString(CultureInfo.InvariantCulture),
                            trade.Pnl.ToString(CultureInfo.InvariantCulture),
                            Csv(string.Empty),
                            Csv(string.Empty),
                            Csv(string.Empty),
                            Csv(string.Empty),
                            Csv(string.Empty),
                            Csv(trade.SourceFile)));
                        result.ImportedRowsExported++;
                    }
                }

                result.OutputPath = filePath;
            }
            catch (Exception ex)
            {
                result.Errors.Add(ex.Message);
            }

            return result;
        }

        private static string FormatDateTime(DateTime value)
        {
            return value == DateTime.MinValue
                ? string.Empty
                : value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string Csv(string value)
        {
            string safe = value ?? string.Empty;
            safe = safe.Replace("\r", " ").Replace("\n", " ");
            if (safe.Contains(",") || safe.Contains("\"") || safe.Contains(" "))
                return "\"" + safe.Replace("\"", "\"\"") + "\"";
            return safe;
        }
        private void RaiseDataChanged()
        {
            EventHandler handler = DataChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }
    }

    public static class TradeAccountBookBuilder_v3
    {
        private const double FallbackCommissionPerContract = 0.51d;

        public static TradeAccountBook Build(List<PersistedExecutionRecord> liveRecords, List<ImportedTradeRecord> importedTrades)
        {
            List<PersistedExecutionRecord> sanitizedLiveRecords = DeduplicateLiveRecords(liveRecords ?? new List<PersistedExecutionRecord>());
            TradeAccountBook book = BuildFromExecutions(sanitizedLiveRecords);
            AddImportedTrades(book, importedTrades ?? new List<ImportedTradeRecord>());
            return book;
        }

        public static MatchedTradeBuildResult BuildMatchedTradeRows(
            List<PersistedExecutionRecord> liveRecords,
            List<ImportedTradeRecord> importedTrades)
        {
            MatchedTradeBuildResult result = new MatchedTradeBuildResult();
            List<PersistedExecutionRecord> sanitizedLiveRecords = DeduplicateLiveRecords(
                liveRecords ?? new List<PersistedExecutionRecord>());

            BuildMatchedTradesFromExecutions(sanitizedLiveRecords, result);
            BuildMatchedTradesFromImportedTrades(importedTrades ?? new List<ImportedTradeRecord>(), result);

            result.Rows.Sort((left, right) =>
            {
                int accountCompare = string.Compare(left.AccountName, right.AccountName, StringComparison.OrdinalIgnoreCase);
                if (accountCompare != 0)
                    return accountCompare;

                int symbolCompare = string.Compare(left.Symbol, right.Symbol, StringComparison.OrdinalIgnoreCase);
                if (symbolCompare != 0)
                    return symbolCompare;

                return left.ExitTime.CompareTo(right.ExitTime);
            });

            result.IdMatchRows = result.Rows.Count(r =>
                string.Equals(r.MatchMethod, "Order ID", StringComparison.OrdinalIgnoreCase)
                || string.Equals(r.MatchMethod, "Order metadata", StringComparison.OrdinalIgnoreCase));
            result.FallbackMatchRows = result.Rows.Count(r =>
                string.Equals(r.MatchMethod, "Lot fallback", StringComparison.OrdinalIgnoreCase));
            return result;
        }

        private static void BuildMatchedTradesFromExecutions(
            List<PersistedExecutionRecord> records,
            MatchedTradeBuildResult result)
        {
            Dictionary<string, List<MatchedOpenLot>> openLotsByAccountAndSymbol =
                new Dictionary<string, List<MatchedOpenLot>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, MatchedTradeRow> aggregatedRows =
                new Dictionary<string, MatchedTradeRow>(StringComparer.OrdinalIgnoreCase);
            int uniqueFallbackSequence = 0;
            int invalidExecutions = 0;

            foreach (PersistedExecutionRecord record in records
                .OrderBy(r => r.AccountName)
                .ThenBy(r => r.Instrument)
                .ThenBy(r => r.Time)
                .ThenBy(r => r.ExecutionId)
                .ThenBy(r => r.OrderId))
            {
                if (record == null || record.Quantity <= 0 || string.IsNullOrWhiteSpace(record.Instrument))
                {
                    invalidExecutions++;
                    continue;
                }

                int signedQty = GetSignedQuantity(record.Action, record.Quantity);
                if (signedQty == 0)
                {
                    invalidExecutions++;
                    continue;
                }

                string lotKey = string.Join("|", record.AccountName ?? string.Empty, record.Instrument);
                List<MatchedOpenLot> openLots;
                if (!openLotsByAccountAndSymbol.TryGetValue(lotKey, out openLots))
                {
                    openLots = new List<MatchedOpenLot>();
                    openLotsByAccountAndSymbol[lotKey] = openLots;
                }

                int incomingSign = Math.Sign(signedQty);
                int remaining = Math.Abs(signedQty);
                double executionCommissionPerUnit = GetCommissionPerUnit(record);

                while (remaining > 0)
                {
                    int openLotIndex = FindBestOpenLotIndex(openLots, incomingSign, record);
                    if (openLotIndex < 0)
                        break;

                    MatchedOpenLot lot = openLots[openLotIndex];
                    int matchedQty = Math.Min(remaining, lot.RemainingQty);
                    string matchMethod = GetMatchMethod(lot, record);
                    double commission = Round2((lot.EntryCommissionPerUnit + executionCommissionPerUnit) * matchedQty);
                    double pnl = Round2(CalculateMatchedTradePnl(lot, record.Price, matchedQty, executionCommissionPerUnit));

                    MatchedTradeRow row = new MatchedTradeRow
                    {
                        EntryTime = lot.EntryTime,
                        ExitTime = record.Time,
                        Symbol = record.Instrument,
                        Direction = lot.Sign > 0 ? "Long" : "Short",
                        Qty = matchedQty,
                        EntryPrice = lot.EntryPrice,
                        ExitPrice = record.Price,
                        Commission = commission,
                        Pnl = pnl,
                        Source = "Live",
                        AccountName = record.AccountName ?? string.Empty,
                        EntryOrderId = lot.OrderId ?? string.Empty,
                        ExitOrderId = record.OrderId ?? string.Empty,
                        EntryExecutionIds = lot.ExecutionId ?? string.Empty,
                        ExitExecutionIds = record.ExecutionId ?? string.Empty,
                        MatchMethod = matchMethod
                    };

                    string aggregateKey = GetMatchedTradeAggregateKey(row, ref uniqueFallbackSequence);
                    MatchedTradeRow existing;
                    if (aggregatedRows.TryGetValue(aggregateKey, out existing))
                        AggregateMatchedTrade(existing, row);
                    else
                    {
                        aggregatedRows[aggregateKey] = row;
                        result.Rows.Add(row);
                    }

                    lot.RemainingQty -= matchedQty;
                    remaining -= matchedQty;

                    if (lot.RemainingQty <= 0)
                        openLots.RemoveAt(openLotIndex);
                }

                if (remaining > 0)
                {
                    openLots.Add(new MatchedOpenLot
                    {
                        EntryTime = record.Time,
                        EntryPrice = record.Price,
                        RemainingQty = remaining,
                        Sign = incomingSign,
                        EntryCommissionPerUnit = executionCommissionPerUnit,
                        PointValue = record.PointValue,
                        AccountName = record.AccountName ?? string.Empty,
                        Symbol = record.Instrument,
                        OrderId = record.OrderId ?? string.Empty,
                        ExecutionId = record.ExecutionId ?? string.Empty,
                        OrderName = record.OrderName ?? string.Empty
                    });
                }
            }

            result.UnmatchedExecutions = invalidExecutions
                + openLotsByAccountAndSymbol.Values.Sum(lots => lots.Count);
        }

        private static void BuildMatchedTradesFromImportedTrades(
            List<ImportedTradeRecord> importedTrades,
            MatchedTradeBuildResult result)
        {
            foreach (ImportedTradeRecord trade in importedTrades)
            {
                if (trade == null || string.IsNullOrWhiteSpace(trade.Symbol))
                    continue;

                result.Rows.Add(new MatchedTradeRow
                {
                    EntryTime = trade.OpenedOn,
                    ExitTime = trade.ClosedOn,
                    Symbol = trade.Symbol,
                    Direction = InferDirectionFromImportedAction(trade.Action),
                    Qty = trade.Quantity,
                    EntryPrice = trade.Price,
                    ExitPrice = trade.Price,
                    Commission = Round2(trade.Commission),
                    Pnl = Round2(trade.Pnl),
                    Source = "Imported",
                    AccountName = trade.AccountName ?? string.Empty,
                    EntryOrderId = string.Empty,
                    ExitOrderId = string.Empty,
                    EntryExecutionIds = string.Empty,
                    ExitExecutionIds = string.Empty,
                    MatchMethod = "Imported"
                });
            }
        }

        private static int FindBestOpenLotIndex(
            List<MatchedOpenLot> openLots,
            int incomingSign,
            PersistedExecutionRecord closingRecord)
        {
            for (int i = 0; i < openLots.Count; i++)
            {
                if (openLots[i].Sign != incomingSign
                    && !string.Equals(GetMatchMethod(openLots[i], closingRecord), "Lot fallback", StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            for (int i = 0; i < openLots.Count; i++)
            {
                if (openLots[i].Sign != incomingSign)
                    return i;
            }

            return -1;
        }

        private static string GetMatchMethod(MatchedOpenLot lot, PersistedExecutionRecord closingRecord)
        {
            if (lot == null || closingRecord == null)
                return "Lot fallback";

            if (!string.IsNullOrWhiteSpace(lot.OrderId)
                && !string.IsNullOrWhiteSpace(closingRecord.OrderId)
                && string.Equals(lot.OrderId, closingRecord.OrderId, StringComparison.OrdinalIgnoreCase))
                return "Order ID";

            string entryRelationship = GetOrderRelationshipKey(lot.OrderName);
            string exitRelationship = GetOrderRelationshipKey(closingRecord.OrderName);
            if (!string.IsNullOrWhiteSpace(entryRelationship)
                && string.Equals(entryRelationship, exitRelationship, StringComparison.OrdinalIgnoreCase))
                return "Order metadata";

            return "Lot fallback";
        }

        private static string GetOrderRelationshipKey(string orderName)
        {
            if (string.IsNullOrWhiteSpace(orderName))
                return string.Empty;

            StringBuilder normalized = new StringBuilder(orderName.Length);
            foreach (char c in orderName.ToUpperInvariant())
                normalized.Append(char.IsLetterOrDigit(c) ? c : ' ');

            HashSet<string> relationshipNoise = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ENTRY", "EXIT", "OPEN", "CLOSE", "LONG", "SHORT", "BUY", "SELL",
                "BOT", "SLD", "BTC", "SSHORT", "STOP", "TARGET", "PROFIT", "LOSS", "ORDER"
            };

            string[] tokens = normalized.ToString()
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", tokens.Where(token => !relationshipNoise.Contains(token)));
        }

        private static string GetMatchedTradeAggregateKey(MatchedTradeRow row, ref int uniqueFallbackSequence)
        {
            if (!string.IsNullOrWhiteSpace(row.EntryOrderId) && !string.IsNullOrWhiteSpace(row.ExitOrderId))
            {
                return string.Join("|",
                    row.AccountName ?? string.Empty,
                    row.Symbol ?? string.Empty,
                    row.Direction ?? string.Empty,
                    row.EntryOrderId,
                    row.ExitOrderId);
            }

            uniqueFallbackSequence++;
            return "UNIQUE|" + uniqueFallbackSequence.ToString(CultureInfo.InvariantCulture);
        }

        private static void AggregateMatchedTrade(MatchedTradeRow target, MatchedTradeRow addition)
        {
            int totalQty = target.Qty + addition.Qty;
            if (totalQty <= 0)
                return;

            target.EntryTime = target.EntryTime <= addition.EntryTime ? target.EntryTime : addition.EntryTime;
            target.ExitTime = target.ExitTime >= addition.ExitTime ? target.ExitTime : addition.ExitTime;
            target.EntryPrice = ((target.EntryPrice * target.Qty) + (addition.EntryPrice * addition.Qty)) / totalQty;
            target.ExitPrice = ((target.ExitPrice * target.Qty) + (addition.ExitPrice * addition.Qty)) / totalQty;
            target.Qty = totalQty;
            target.Commission = Round2(target.Commission + addition.Commission);
            target.Pnl = Round2(target.Pnl + addition.Pnl);
            target.EntryExecutionIds = CombineIds(target.EntryExecutionIds, addition.EntryExecutionIds);
            target.ExitExecutionIds = CombineIds(target.ExitExecutionIds, addition.ExitExecutionIds);
            if (string.Equals(target.MatchMethod, "Lot fallback", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(addition.MatchMethod, "Lot fallback", StringComparison.OrdinalIgnoreCase))
                target.MatchMethod = addition.MatchMethod;
        }

        private static string CombineIds(string existingIds, string additionalIds)
        {
            return string.Join(",",
                (existingIds ?? string.Empty)
                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                    .Concat((additionalIds ?? string.Empty).Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    .Select(id => id.Trim())
                    .Where(id => id.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase));
        }

        private static double CalculateMatchedTradePnl(
            MatchedOpenLot lot,
            double exitPrice,
            int qty,
            double exitCommissionPerUnit)
        {
            double grossPerUnit = lot.Sign > 0
                ? (exitPrice - lot.EntryPrice) * lot.PointValue
                : (lot.EntryPrice - exitPrice) * lot.PointValue;
            return (grossPerUnit * qty)
                - ((lot.EntryCommissionPerUnit + exitCommissionPerUnit) * qty);
        }

        private static int GetActionPreferenceScoreForBuild(string action)
        {
            string normalized = (action ?? string.Empty).Trim().ToUpperInvariant();
            switch (normalized)
            {
                case "BUY":
                case "SELL":
                    return 4;
                case "BUYTOCOVER":
                case "SELLSHORT":
                    return 3;
                case "UNKNOWN":
                    return 1;
                default:
                    return 2;
            }
        }

        private static List<PersistedExecutionRecord> DeduplicateLiveRecords(List<PersistedExecutionRecord> records)
        {
            return records
                .Where(r => r != null)
                .GroupBy(r =>
                {
                    if (!string.IsNullOrWhiteSpace(r.ExecutionId))
                        return string.Join("|",
                            r.AccountName ?? string.Empty,
                            r.ExecutionId ?? string.Empty);

                    return string.Join("|",
                        r.AccountName ?? string.Empty,
                        r.Instrument ?? string.Empty,
                        (r.Action ?? string.Empty).Trim().ToUpperInvariant(),
                        r.Quantity.ToString(CultureInfo.InvariantCulture),
                        Math.Round(r.Price, 8).ToString("R", CultureInfo.InvariantCulture),
                        Math.Round(r.Commission, 8).ToString("R", CultureInfo.InvariantCulture),
                        r.Time.ToString("yyyy-MM-dd HH:mm:ss.fffffff", CultureInfo.InvariantCulture),
                        r.MarketPosition ?? string.Empty,
                        Math.Round(r.PointValue, 8).ToString("R", CultureInfo.InvariantCulture));
                })
                .Select(g => g
                    .OrderByDescending(x => GetActionPreferenceScoreForBuild(x.Action))
                    .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.ExecutionId))
                    .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.OrderId))
                    .ThenByDescending(x => !string.IsNullOrWhiteSpace(x.OrderName))
                    .ThenByDescending(x => x.Time)
                    .ThenBy(x => x.Key)
                    .First())
                .OrderBy(r => r.Time)
                .ThenBy(r => r.ExecutionId)
                .ThenBy(r => r.OrderId)
                .ToList();
        }

        private static TradeAccountBook BuildFromExecutions(List<PersistedExecutionRecord> records)
        {
            TradeAccountBook book = new TradeAccountBook();
            Dictionary<string, Queue<OpenLot>> openLotsBySymbol = new Dictionary<string, Queue<OpenLot>>(StringComparer.OrdinalIgnoreCase);

            foreach (PersistedExecutionRecord record in records)
            {
                if (record == null || record.Quantity <= 0 || string.IsNullOrWhiteSpace(record.Instrument))
                    continue;

                int signedQty = GetSignedQuantity(record.Action, record.Quantity);
                if (signedQty == 0)
                    continue;

                string symbol = record.Instrument;
                Queue<OpenLot> queue;
                if (!openLotsBySymbol.TryGetValue(symbol, out queue))
                {
                    queue = new Queue<OpenLot>();
                    openLotsBySymbol[symbol] = queue;
                }

                double execCommissionPerUnit = GetCommissionPerUnit(record);
                int incomingSign = Math.Sign(signedQty);
                int remaining = Math.Abs(signedQty);

                while (remaining > 0 && queue.Count > 0 && queue.Peek().Sign != incomingSign)
                {
                    OpenLot lot = queue.Peek();
                    int matchedQty = Math.Min(remaining, lot.RemainingQty);

                    double pnl = CalculateRealizedPnl(lot, record.Price, matchedQty, execCommissionPerUnit);
                    DateTime closeDate = TradeSessionCalendar_v3.ToTradingDate(record.Time);

                    book.AddRealized(closeDate, symbol, pnl);
                    book.AddRow(new ExecutionBreakdownRow
                    {
                        Time = record.Time,
                        OpenedOn = lot.EntryTime,
                        ClosedOn = record.Time,
                        Symbol = symbol,
                        Action = ToShortAction(record.Action),
                        Qty = matchedQty,
                        Price = record.Price,
                        Commission = Round2((lot.EntryCommissionPerUnit + execCommissionPerUnit) * matchedQty),
                        Pnl = Round2(pnl),
                        Source = "Live",
                        Direction = lot.Sign > 0 ? "Long" : "Short"
                    });

                    lot.RemainingQty -= matchedQty;
                    remaining -= matchedQty;

                    if (lot.RemainingQty <= 0)
                        queue.Dequeue();
                }

                if (remaining > 0)
                {
                    OpenLot newLot = new OpenLot
                    {
                        EntryTime = record.Time,
                        EntryPrice = record.Price,
                        RemainingQty = remaining,
                        EntryCommissionPerUnit = execCommissionPerUnit,
                        Sign = incomingSign,
                        PointValue = record.PointValue,
                        Symbol = symbol
                    };
                    queue.Enqueue(newLot);

                    book.AddActivityDate(TradeSessionCalendar_v3.ToTradingDate(record.Time));
                    book.AddRow(new ExecutionBreakdownRow
                    {
                        Time = record.Time,
                        OpenedOn = record.Time,
                        ClosedOn = null,
                        Symbol = symbol,
                        Action = ToShortAction(record.Action),
                        Qty = remaining,
                        Price = record.Price,
                        Commission = Round2(execCommissionPerUnit * remaining),
                        Pnl = null,
                        Source = "Live",
                        Direction = incomingSign > 0 ? "Long" : "Short"
                    });
                }
                else
                {
                    book.AddActivityDate(TradeSessionCalendar_v3.ToTradingDate(record.Time));
                }
            }

            return book;
        }

        private static void AddImportedTrades(TradeAccountBook book, List<ImportedTradeRecord> importedTrades)
        {
            foreach (ImportedTradeRecord trade in importedTrades)
            {
                if (trade == null || string.IsNullOrWhiteSpace(trade.Symbol))
                    continue;

                book.AddActivityDate(TradeSessionCalendar_v3.ToTradingDate(trade.OpenedOn));
                book.AddActivityDate(TradeSessionCalendar_v3.ToTradingDate(trade.ClosedOn));
                book.AddRealized(TradeSessionCalendar_v3.ToTradingDate(trade.ClosedOn), trade.Symbol, trade.Pnl);

                book.AddRow(new ExecutionBreakdownRow
                {
                    Time = trade.ClosedOn,
                    OpenedOn = trade.OpenedOn,
                    ClosedOn = trade.ClosedOn,
                    Symbol = trade.Symbol,
                    Action = NormalizeImportedAction(trade.Action),
                    Qty = trade.Quantity,
                    Price = trade.Price,
                    Commission = Round2(trade.Commission),
                    Pnl = Round2(trade.Pnl),
                    Source = "Imported",
                    Direction = InferDirectionFromImportedAction(trade.Action)
                });
            }
        }

        private static string NormalizeImportedAction(string action)
        {
            string a = (action ?? string.Empty).Trim();
            if (a.Length == 0)
                return "TRADE";

            string n = a.ToLowerInvariant();
            if (n == "buy")
                return "BOT";
            if (n == "sell")
                return "SLD";
            if (n == "buytocover")
                return "BTC";
            if (n == "sellshort")
                return "SSHORT";
            if (n == "long")
                return "LONG";
            if (n == "short")
                return "SHORT";
            return a.ToUpperInvariant();
        }

        private static string InferDirectionFromImportedAction(string action)
        {
            string a = (action ?? string.Empty).Trim().ToLowerInvariant();
            if (a.Contains("short") || a == "sellshort" || a == "sshort")
                return "Short";
            if (a == "short")
                return "Short";
            if (a.Contains("long") || a == "buy" || a == "bot")
                return "Long";
            if (a == "sell" || a == "sld")
                return "Long";
            if (a == "buytocover" || a == "btc")
                return "Short";
            return string.Empty;
        }

        private static double CalculateRealizedPnl(OpenLot lot, double exitPrice, int qty, double exitCommissionPerUnit)
        {
            double grossPerUnit;

            if (lot.Sign > 0)
                grossPerUnit = (exitPrice - lot.EntryPrice) * lot.PointValue;
            else
                grossPerUnit = (lot.EntryPrice - exitPrice) * lot.PointValue;

            double gross = grossPerUnit * qty;
            double commission = (lot.EntryCommissionPerUnit + exitCommissionPerUnit) * qty;
            return gross - commission;
        }

        private static double GetCommissionPerUnit(PersistedExecutionRecord record)
        {
            if (record == null || record.Quantity <= 0)
                return 0d;

            if (Math.Abs(record.Commission) > 0.000001)
                return record.Commission / record.Quantity;

            return FallbackCommissionPerContract;
        }

        private static int GetSignedQuantity(string action, int qty)
        {
            if (qty <= 0)
                return 0;

            switch ((action ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "buy":
                case "buytocover":
                case "bot":
                case "btc":
                    return qty;

                case "sell":
                case "sellshort":
                case "sld":
                case "sshort":
                    return -qty;
            }

            return 0;
        }

        private static string ToShortAction(string action)
        {
            switch ((action ?? string.Empty).Trim().ToLowerInvariant())
            {
                case "buy":
                    return "BOT";
                case "buytocover":
                    return "BTC";
                case "sell":
                    return "SLD";
                case "sellshort":
                    return "SSHORT";
                default:
                    return (action ?? string.Empty).ToUpperInvariant();
            }
        }

        private static double Round2(double value)
        {
            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }
    }

    public static class CurrentTradeListCsvExporter
    {
        public static void ExportMatchedTrades(string filePath, IEnumerable<MatchedTradeRow> rows)
        {
            EnsureOutputDirectory(filePath);

            using (StreamWriter writer = new StreamWriter(filePath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("EntryTime,ExitTime,Symbol,Side,Qty,EntryPrice,ExitPrice,Commission,PnL,Source,AccountName,EntryOrderId,ExitOrderId,EntryExecutionIds,ExitExecutionIds,MatchMethod");

                foreach (MatchedTradeRow row in rows ?? Enumerable.Empty<MatchedTradeRow>())
                {
                    if (row == null)
                        continue;

                    writer.WriteLine(string.Join(",",
                        CsvEscape(FormatDateTime(row.EntryTime)),
                        CsvEscape(FormatDateTime(row.ExitTime)),
                        CsvEscape(row.Symbol),
                        CsvEscape(row.Direction),
                        row.Qty.ToString(CultureInfo.InvariantCulture),
                        FormatNumber(row.EntryPrice),
                        FormatNumber(row.ExitPrice),
                        FormatNumber(row.Commission),
                        FormatNumber(row.Pnl),
                        CsvEscape(row.Source),
                        CsvEscape(row.AccountName),
                        CsvEscape(row.EntryOrderId),
                        CsvEscape(row.ExitOrderId),
                        CsvEscape(row.EntryExecutionIds),
                        CsvEscape(row.ExitExecutionIds),
                        CsvEscape(row.MatchMethod)));
                }
            }
        }

        public static void ExportExecutionLegs(string filePath, IEnumerable<ExecutionBreakdownRow> rows)
        {
            EnsureOutputDirectory(filePath);

            using (StreamWriter writer = new StreamWriter(filePath, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("Time,OpenedOn,ClosedOn,Symbol,Direction,Action,Qty,Price,Commission,PnL,Source");

                foreach (ExecutionBreakdownRow row in rows ?? Enumerable.Empty<ExecutionBreakdownRow>())
                {
                    if (row == null)
                        continue;

                    writer.WriteLine(string.Join(",",
                        CsvEscape(FormatDateTime(row.Time)),
                        CsvEscape(FormatDateTime(row.OpenedOn)),
                        CsvEscape(row.ClosedOn.HasValue ? FormatDateTime(row.ClosedOn.Value) : string.Empty),
                        CsvEscape(row.Symbol),
                        CsvEscape(row.Direction),
                        CsvEscape(row.Action),
                        row.Qty.ToString(CultureInfo.InvariantCulture),
                        FormatNumber(row.Price),
                        FormatNumber(row.Commission),
                        row.Pnl.HasValue ? FormatNumber(row.Pnl.Value) : string.Empty,
                        CsvEscape(row.Source)));
                }
            }
        }

        public static string SanitizeFileNamePart(string value)
        {
            string safe = string.IsNullOrWhiteSpace(value) ? "Unknown" : value.Trim();
            HashSet<char> invalid = new HashSet<char>(Path.GetInvalidFileNameChars());
            StringBuilder sanitized = new StringBuilder(safe.Length);

            foreach (char c in safe)
                sanitized.Append(invalid.Contains(c) || char.IsWhiteSpace(c) ? '_' : c);

            string result = sanitized.ToString().Trim('_', '.');
            return result.Length == 0 ? "Unknown" : result;
        }

        public static string CsvEscape(string value)
        {
            string safe = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ");
            if (safe.Contains(",") || safe.Contains("\""))
                return "\"" + safe.Replace("\"", "\"\"") + "\"";
            return safe;
        }

        private static void EnsureOutputDirectory(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("No output file was selected.", "filePath");

            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);
        }

        private static string FormatDateTime(DateTime value)
        {
            return value == DateTime.MinValue
                ? string.Empty
                : value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private static string FormatNumber(double value)
        {
            return value.ToString("R", CultureInfo.InvariantCulture);
        }
    }


    public static class TradeSessionCalendar_v3
    {
        // Futures-style trading day rollover. Times at or after 18:00 belong to the next trading day.
        // This matches the common CME evening session behavior much more closely than simple wall-clock dates.
        public static readonly TimeSpan TradingDayRollover = new TimeSpan(18, 0, 0);

        public static DateTime ToTradingDate(DateTime timestamp)
        {
            DateTime local = timestamp;
            DateTime date = local.Date;
            if (local.TimeOfDay >= TradingDayRollover)
                date = date.AddDays(1);
            return date;
        }
    }

    public static class CsvTradeImportParser
    {
        public static CsvParseResult ParseFile(string path)
        {
            CsvParseResult result = new CsvParseResult();

            try
            {
                if (!File.Exists(path))
                {
                    result.Errors.Add("File not found: " + path);
                    return result;
                }

                string[] lines = File.ReadAllLines(path);
                if (lines == null || lines.Length == 0)
                {
                    result.Errors.Add("File is empty: " + Path.GetFileName(path));
                    return result;
                }

                int headerIndex = FindHeaderIndex(lines);
                if (headerIndex < 0)
                {
                    result.Errors.Add("Could not find a usable header row in: " + Path.GetFileName(path));
                    return result;
                }

                List<string> rawHeaders = ParseCsvLine(lines[headerIndex]);
                List<string> headers = rawHeaders.Select(h => NormalizeHeader(h)).ToList();

                for (int i = headerIndex + 1; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    List<string> cells = ParseCsvLine(line);
                    if (cells.Count == 0)
                        continue;

                    result.RowsRead++;

                    ImportedTradeRecord trade;
                    string skipReason;
                    if (TryParseTradeRow(path, i + 1, headers, cells, out trade, out skipReason))
                    {
                        result.Trades.Add(trade);
                    }
                    else
                    {
                        result.RowsSkipped++;
                        if (!string.IsNullOrWhiteSpace(skipReason))
                            result.Errors.Add(skipReason);
                    }
                }
            }
            catch (Exception ex)
            {
                result.Errors.Add("Failed to parse " + Path.GetFileName(path) + ": " + ex.Message);
            }

            return result;
        }

        private static int FindHeaderIndex(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i]))
                    continue;

                List<string> cells = ParseCsvLine(lines[i]);
                List<string> normalized = cells.Select(c => NormalizeHeader(c)).ToList();

                bool hasSymbol = normalized.Contains("symbol") || normalized.Contains("instrument") || normalized.Contains("masterinstrument") || normalized.Contains("contract");
                bool hasPnl = normalized.Contains("pl") || normalized.Contains("pnl") || normalized.Contains("profit") || normalized.Contains("profitloss") || normalized.Contains("realizedpnl") || normalized.Contains("netprofit");
                bool hasTime = normalized.Contains("time") || normalized.Contains("openedon") || normalized.Contains("entrytime") || normalized.Contains("closedon") || normalized.Contains("exittime") || normalized.Contains("date");

                if (hasSymbol && hasPnl && hasTime)
                    return i;
            }

            return -1;
        }

        private static bool TryParseTradeRow(string path, int lineNumber, List<string> headers, List<string> cells, out ImportedTradeRecord trade, out string skipReason)
        {
            trade = null;
            skipReason = null;

            Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < headers.Count; i++)
            {
                string value = i < cells.Count ? cells[i] : string.Empty;
                row[headers[i]] = value;
            }

            string symbol = GetValue(row, "symbol", "instrument", "masterinstrument", "instrumentname", "contract");
            if (string.IsNullOrWhiteSpace(symbol))
            {
                skipReason = FormatSkip(path, lineNumber, "missing symbol/instrument");
                return false;
            }

            DateTime openedOn;
            DateTime closedOn;

            bool hasOpen = TryParseDate(GetValue(row, "openedon", "entrytime", "entrydate", "opentime", "starttime", "dateopen"), out openedOn);
            bool hasClose = TryParseDate(GetValue(row, "closedon", "exittime", "exitdate", "closetime", "time", "datetime", "dateclose", "date"), out closedOn);

            if (!hasOpen && hasClose)
                openedOn = closedOn;
            if (!hasClose && hasOpen)
                closedOn = openedOn;

            if (!hasOpen && !hasClose)
            {
                skipReason = FormatSkip(path, lineNumber, "missing opened/closed time columns");
                return false;
            }

            double pnl;
            if (!TryParseMoney(GetValue(row, "pl", "pnl", "profit", "profitloss", "netprofit", "realizedpnl", "realizedpl", "realizedprofitloss", "profitcurrency", "netpnl"), out pnl))
            {
                skipReason = FormatSkip(path, lineNumber, "missing PnL value");
                return false;
            }

            int qty;
            if (!TryParseInt(GetValue(row, "qty", "quantity", "contracts", "size", "shares"), out qty) || qty <= 0)
                qty = 1;

            double commission;
            if (!TryParseMoney(GetValue(row, "commission", "commissions", "fees", "commissioncurrency"), out commission))
                commission = 0d;

            double price;
            if (!TryParseMoney(GetValue(row, "price", "avgexitprice", "exitprice", "avgfillprice", "fillprice", "avgprice", "closeprice", "tradeprice", "entryprice", "avgentryprice"), out price))
                price = 0d;

            string action = GetValue(row, "action", "side", "marketposition", "position", "direction");
            string account = GetValue(row, "account", "acct", "accountname", "accountid");

            trade = new ImportedTradeRecord
            {
                AccountName = account,
                Symbol = symbol.Trim(),
                Action = string.IsNullOrWhiteSpace(action) ? "TRADE" : action.Trim(),
                Quantity = qty,
                OpenedOn = openedOn,
                ClosedOn = closedOn,
                Price = price,
                Commission = commission,
                Pnl = pnl,
                SourceFile = Path.GetFileName(path)
            };

            return true;
        }

        private static string GetValue(Dictionary<string, string> row, params string[] keys)
        {
            foreach (string key in keys)
            {
                string value;
                if (row.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }

            return string.Empty;
        }

        private static string NormalizeHeader(string header)
        {
            if (string.IsNullOrWhiteSpace(header))
                return string.Empty;

            StringBuilder sb = new StringBuilder();
            foreach (char c in header)
            {
                if (char.IsLetterOrDigit(c))
                    sb.Append(char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

        private static List<string> ParseCsvLine(string line)
        {
            List<string> values = new List<string>();
            if (line == null)
                return values;

            StringBuilder sb = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        sb.Append('"');
                        i++;
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    values.Add(sb.ToString().Trim());
                    sb.Length = 0;
                }
                else
                {
                    sb.Append(c);
                }
            }

            values.Add(sb.ToString().Trim());
            return values;
        }

        private static bool TryParseDate(string text, out DateTime value)
        {
            value = DateTime.MinValue;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string cleaned = text.Trim();

            DateTime parsed;
            if (DateTime.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out parsed))
            {
                value = parsed;
                return true;
            }

            if (DateTime.TryParse(cleaned, CultureInfo.CurrentCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out parsed))
            {
                value = parsed;
                return true;
            }

            string[] formats =
            {
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd H:mm:ss",
                "yyyy/MM/dd HH:mm:ss",
                "M/d/yyyy h:mm:ss tt",
                "M/d/yyyy H:mm",
                "M/d/yyyy H:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "MM/dd/yyyy hh:mm:ss tt",
                "dd/MM/yyyy HH:mm:ss"
            };

            if (DateTime.TryParseExact(cleaned, formats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal, out parsed))
            {
                value = parsed;
                return true;
            }

            return false;
        }

        private static bool TryParseMoney(string text, out double value)
        {
            value = 0d;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string cleaned = text.Trim();
            bool negative = false;

            if (cleaned.StartsWith("(") && cleaned.EndsWith(")"))
            {
                negative = true;
                cleaned = cleaned.Substring(1, cleaned.Length - 2);
            }

            cleaned = cleaned.Replace("$", string.Empty)
                             .Replace(",", string.Empty)
                             .Replace(" ", string.Empty)
                             .Replace("USD", string.Empty)
                             .Replace("CAD", string.Empty);

            double parsed;
            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out parsed) ||
                double.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out parsed))
            {
                value = negative ? -parsed : parsed;
                return true;
            }

            return false;
        }

        private static bool TryParseInt(string text, out int value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string cleaned = text.Trim().Replace(",", string.Empty);
            int parsedInt;
            if (int.TryParse(cleaned, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedInt))
            {
                value = Math.Abs(parsedInt);
                return true;
            }

            double parsedDouble;
            if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out parsedDouble) ||
                double.TryParse(cleaned, NumberStyles.Any, CultureInfo.CurrentCulture, out parsedDouble))
            {
                value = (int)Math.Abs(Math.Round(parsedDouble, MidpointRounding.AwayFromZero));
                return true;
            }

            return false;
        }

        private static string FormatSkip(string path, int lineNumber, string reason)
        {
            return Path.GetFileName(path) + " line " + lineNumber.ToString(CultureInfo.InvariantCulture) + ": " + reason;
        }
    }

    [Serializable]
    public class PersistedExecutionRecord
    {
        public string Key { get; set; }
        public string AccountName { get; set; }
        public string Instrument { get; set; }
        public string Action { get; set; }
        public int Quantity { get; set; }
        public double Price { get; set; }
        public double Commission { get; set; }
        public DateTime Time { get; set; }
        public string OrderId { get; set; }
        public string ExecutionId { get; set; }
        public string MarketPosition { get; set; }
        public string OrderName { get; set; }
        public double PointValue { get; set; }
    }

    [Serializable]
    public class ImportedTradeRecord
    {
        public string Key { get; set; }
        public string AccountName { get; set; }
        public string Symbol { get; set; }
        public string Action { get; set; }
        public int Quantity { get; set; }
        public DateTime OpenedOn { get; set; }
        public DateTime ClosedOn { get; set; }
        public double Price { get; set; }
        public double Commission { get; set; }
        public double Pnl { get; set; }
        public string SourceFile { get; set; }
    }

    public class OpenLot
    {
        public DateTime EntryTime { get; set; }
        public double EntryPrice { get; set; }
        public int RemainingQty { get; set; }
        public int Sign { get; set; }
        public double EntryCommissionPerUnit { get; set; }
        public double PointValue { get; set; }
        public string Symbol { get; set; }
    }

    public class MatchedOpenLot
    {
        public DateTime EntryTime { get; set; }
        public double EntryPrice { get; set; }
        public int RemainingQty { get; set; }
        public int Sign { get; set; }
        public double EntryCommissionPerUnit { get; set; }
        public double PointValue { get; set; }
        public string AccountName { get; set; }
        public string Symbol { get; set; }
        public string OrderId { get; set; }
        public string ExecutionId { get; set; }
        public string OrderName { get; set; }
    }

    public class MatchedTradeRow
    {
        public DateTime EntryTime { get; set; }
        public DateTime ExitTime { get; set; }
        public string Symbol { get; set; }
        public string Direction { get; set; }
        public int Qty { get; set; }
        public double EntryPrice { get; set; }
        public double ExitPrice { get; set; }
        public double Commission { get; set; }
        public double Pnl { get; set; }
        public string Source { get; set; }
        public string AccountName { get; set; }
        public string EntryOrderId { get; set; }
        public string ExitOrderId { get; set; }
        public string EntryExecutionIds { get; set; }
        public string ExitExecutionIds { get; set; }
        public string MatchMethod { get; set; }

        [XmlIgnore]
        public string EntryTimeDisplay
        {
            get { return EntryTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string ExitTimeDisplay
        {
            get { return ExitTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string PnlDisplay
        {
            get { return Pnl.ToString("N2", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string EntryPriceDisplay
        {
            get { return EntryPrice.ToString("N2", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string ExitPriceDisplay
        {
            get { return ExitPrice.ToString("N2", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string CommissionDisplay
        {
            get { return Commission.ToString("N2", CultureInfo.InvariantCulture); }
            set { }
        }
    }

    public class MatchedTradeBuildResult
    {
        public List<MatchedTradeRow> Rows { get; private set; }
        public int IdMatchRows { get; set; }
        public int FallbackMatchRows { get; set; }
        public int UnmatchedExecutions { get; set; }

        public MatchedTradeBuildResult()
        {
            Rows = new List<MatchedTradeRow>();
        }
    }

    public class ExecutionBreakdownRow
    {
        public DateTime Time { get; set; }
        public DateTime OpenedOn { get; set; }
        public DateTime? ClosedOn { get; set; }
        public string Symbol { get; set; }
        public string Direction { get; set; }
        public string Action { get; set; }
        public int Qty { get; set; }
        public double Price { get; set; }
        public double Commission { get; set; }
        public double? Pnl { get; set; }
        public string Source { get; set; }

        [XmlIgnore]
        public string OpenedOnDisplay
        {
            get { return OpenedOn.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture); }
            set { }
        }

        [XmlIgnore]
        public string ClosedOnDisplay
        {
            get { return ClosedOn.HasValue ? ClosedOn.Value.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) : "open"; }
            set { }
        }

        [XmlIgnore]
        public string PnlDisplay
        {
            get { return Pnl.HasValue ? string.Format(CultureInfo.InvariantCulture, "{0:N2}", Pnl.Value) : string.Empty; }
            set { }
        }

        public ExecutionBreakdownRow Clone()
        {
            return new ExecutionBreakdownRow
            {
                Time = Time,
                OpenedOn = OpenedOn,
                ClosedOn = ClosedOn,
                Symbol = Symbol,
                Direction = Direction,
                Action = Action,
                Qty = Qty,
                Price = Price,
                Commission = Commission,
                Pnl = Pnl,
                Source = Source
            };
        }
    }

    public class DayBook
    {
        public DateTime Date { get; set; }
        public double TotalPnl { get; set; }
        public Dictionary<string, double> InstrumentPnl { get; private set; }
        public List<ExecutionBreakdownRow> Rows { get; private set; }
        public bool HasActivity { get; set; }

        public DayBook()
        {
            InstrumentPnl = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            Rows = new List<ExecutionBreakdownRow>();
        }
    }


    public class TradeAccountBook
    {
        private readonly Dictionary<DateTime, DayBook> days = new Dictionary<DateTime, DayBook>();

        public IEnumerable<KeyValuePair<DateTime, DayBook>> Days
        {
            get { return days; }
        }

        public void AddActivityDate(DateTime date)
        {
            GetOrCreate(date).HasActivity = true;
        }

        public void AddRealized(DateTime date, string symbol, double pnl)
        {
            DayBook day = GetOrCreate(date);
            day.TotalPnl += pnl;
            day.HasActivity = true;

            double current;
            if (!day.InstrumentPnl.TryGetValue(symbol, out current))
                current = 0d;

            day.InstrumentPnl[symbol] = current + pnl;
        }

        public void AddRow(ExecutionBreakdownRow row)
        {
            if (row == null)
                return;

            GetOrCreate(TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn)).HasActivity = true;

            if (row.ClosedOn.HasValue)
            {
                DateTime closeTradeDate = TradeSessionCalendar_v3.ToTradingDate(row.ClosedOn.Value);
                DayBook closeDay = GetOrCreate(closeTradeDate);
                closeDay.Rows.Add(row);
                closeDay.HasActivity = true;

                DateTime openTradeDate = TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn);
                if (openTradeDate != closeTradeDate)
                    GetOrCreate(openTradeDate).HasActivity = true;
            }
            else
            {
                DayBook openDay = GetOrCreate(TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn));
                openDay.Rows.Add(row);
                openDay.HasActivity = true;
            }
        }

        public void AddFilteredRow(ExecutionBreakdownRow row, bool includeOpenDateActivity)
        {
            if (row == null)
                return;

            DateTime bucketDate = row.ClosedOn.HasValue ? TradeSessionCalendar_v3.ToTradingDate(row.ClosedOn.Value) : TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn);
            DayBook bucket = GetOrCreate(bucketDate);
            bucket.Rows.Add(row);
            bucket.HasActivity = true;

            DateTime openTradeDate = TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn);
            if (includeOpenDateActivity && openTradeDate != bucketDate)
                GetOrCreate(openTradeDate).HasActivity = true;
        }

        public TradeAccountBook ApplyFilters(TradeQueryFilters filters)
        {
            if (filters == null || !filters.HasAnyRowLevelFilter)
                return this;

            TradeAccountBook filtered = new TradeAccountBook();

            foreach (KeyValuePair<DateTime, DayBook> kv in days.OrderBy(k => k.Key))
            {
                foreach (ExecutionBreakdownRow originalRow in kv.Value.Rows)
                {
                    if (originalRow == null || !filters.MatchesRow(originalRow))
                        continue;

                    ExecutionBreakdownRow row = originalRow.Clone();
                    bool includeOpenDateActivity = !row.ClosedOn.HasValue || filters.MatchesDate(TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn));
                    filtered.AddFilteredRow(row, includeOpenDateActivity);

                    if (row.Pnl.HasValue)
                    {
                        DateTime pnlDate = row.ClosedOn.HasValue ? TradeSessionCalendar_v3.ToTradingDate(row.ClosedOn.Value) : TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn);
                        filtered.AddRealized(pnlDate, row.Symbol, row.Pnl.Value);
                    }
                }
            }

            return filtered;
        }

        public DayBook GetDayBook(DateTime date)
        {
            DayBook day;
            if (days.TryGetValue(date.Date, out day))
            {
                DayBook copy = new DayBook
                {
                    Date = day.Date,
                    TotalPnl = Math.Round(day.TotalPnl, 2, MidpointRounding.AwayFromZero),
                    HasActivity = day.HasActivity
                };

                foreach (KeyValuePair<string, double> kv in day.InstrumentPnl)
                    copy.InstrumentPnl[kv.Key] = Math.Round(kv.Value, 2, MidpointRounding.AwayFromZero);

                copy.Rows.AddRange(day.Rows.Where(r => TradeSessionCalendar_v3.ToTradingDate(r.OpenedOn) == date.Date || (r.ClosedOn.HasValue && TradeSessionCalendar_v3.ToTradingDate(r.ClosedOn.Value) == date.Date)));
                return copy;
            }

            return new DayBook { Date = date.Date };
        }

        public bool DayHasAnyActivity(DateTime date)
        {
            DayBook day;
            return days.TryGetValue(date.Date, out day) && day.HasActivity;
        }

        public double GetDayTotal(DateTime date)
        {
            DayBook day;
            return days.TryGetValue(date.Date, out day) ? Math.Round(day.TotalPnl, 2, MidpointRounding.AwayFromZero) : 0d;
        }

        public double GetMonthTotal(DateTime firstOfMonth)
        {
            return days.Where(kv => kv.Key.Year == firstOfMonth.Year && kv.Key.Month == firstOfMonth.Month)
                       .Sum(kv => kv.Value.TotalPnl);
        }

        public double GetWeekTotalForSaturday(DateTime saturday)
        {
            DateTime start = saturday.Date.AddDays(-6);
            DateTime end = saturday.Date;
            return days.Where(kv => kv.Key >= start && kv.Key <= end)
                       .Sum(kv => kv.Value.TotalPnl);
        }

        private DayBook GetOrCreate(DateTime date)
        {
            DayBook day;
            if (!days.TryGetValue(date.Date, out day))
            {
                day = new DayBook { Date = date.Date };
                days[date.Date] = day;
            }
            return day;
        }
    }


    [Serializable]
    public class TradeCalendarUiState
    {
        public string AccountName { get; set; }
        public bool AllAccounts { get; set; }
        public string LastSelectedAccountName { get; set; }
        public int? LastSelectedAccountIndex { get; set; }
        public string InstrumentFilter { get; set; }
        public string SourceFilter { get; set; }
        public string SideFilter { get; set; }
        public string ResultFilter { get; set; }
        public DateTime? FromDate { get; set; }
        public DateTime? ToDate { get; set; }
        public DateTime CurrentMonth { get; set; }
        public DateTime SelectedDate { get; set; }
        public string TradeListViewMode { get; set; }

        [XmlIgnore]
        private static string SettingsPath
        {
            get { return Path.Combine(NinjaTrader.Core.Globals.UserDataDir, "TradeCalendarUiState.xml"); }
        }

        public void Save()
        {
            string directory = Path.GetDirectoryName(SettingsPath);
            if (!Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            XmlSerializer serializer = new XmlSerializer(typeof(TradeCalendarUiState));
            using (FileStream stream = File.Create(SettingsPath))
                serializer.Serialize(stream, this);
        }

        public static TradeCalendarUiState Load()
        {
            try
            {
                if (!File.Exists(SettingsPath))
                    return new TradeCalendarUiState();

                XmlSerializer serializer = new XmlSerializer(typeof(TradeCalendarUiState));
                using (FileStream stream = File.OpenRead(SettingsPath))
                    return serializer.Deserialize(stream) as TradeCalendarUiState ?? new TradeCalendarUiState();
            }
            catch
            {
                return new TradeCalendarUiState();
            }
        }
    }

    public class TradeQueryFilters
    {
        public string AccountName { get; set; }
        public string Instrument { get; set; }
        public string Source { get; set; }
        public string Side { get; set; }
        public string Result { get; set; }
        public DateTime? FromDate { get; set; }
        public DateTime? ToDate { get; set; }

        [XmlIgnore]
        public bool HasInstrumentFilter
        {
            get { return !string.IsNullOrWhiteSpace(Instrument) && !string.Equals(Instrument, "All instruments", StringComparison.OrdinalIgnoreCase); }
        }

        [XmlIgnore]
        public bool HasSourceFilter
        {
            get { return !string.IsNullOrWhiteSpace(Source) && !string.Equals(Source, "All sources", StringComparison.OrdinalIgnoreCase); }
        }

        [XmlIgnore]
        public bool HasSideFilter
        {
            get { return !string.IsNullOrWhiteSpace(Side) && !string.Equals(Side, "All sides", StringComparison.OrdinalIgnoreCase); }
        }

        [XmlIgnore]
        public bool HasResultFilter
        {
            get { return !string.IsNullOrWhiteSpace(Result) && !string.Equals(Result, "All results", StringComparison.OrdinalIgnoreCase); }
        }

        [XmlIgnore]
        public bool HasDateFilter
        {
            get { return FromDate.HasValue || ToDate.HasValue; }
        }

        [XmlIgnore]
        public bool HasAnyRowLevelFilter
        {
            get { return HasInstrumentFilter || HasSourceFilter || HasSideFilter || HasResultFilter || HasDateFilter; }
        }

        public bool MatchesRow(ExecutionBreakdownRow row)
        {
            if (row == null)
                return false;

            if (HasInstrumentFilter && !string.Equals(row.Symbol, Instrument, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasSourceFilter && !string.Equals(row.Source, Source, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasSideFilter && !string.Equals(row.Direction, Side, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasResultFilter)
            {
                string normalized = (Result ?? string.Empty).Trim().ToLowerInvariant();
                if (normalized == "winning" && (!row.Pnl.HasValue || row.Pnl.Value <= 0))
                    return false;
                if (normalized == "losing" && (!row.Pnl.HasValue || row.Pnl.Value >= 0))
                    return false;
                if (normalized == "breakeven" && (!row.Pnl.HasValue || Math.Abs(row.Pnl.Value) > 0.000001))
                    return false;
                if (normalized == "open" && row.Pnl.HasValue)
                    return false;
            }

            DateTime effectiveDate = row.ClosedOn.HasValue ? TradeSessionCalendar_v3.ToTradingDate(row.ClosedOn.Value) : TradeSessionCalendar_v3.ToTradingDate(row.OpenedOn);
            if (!MatchesDate(effectiveDate))
                return false;

            return true;
        }

        public bool MatchesDate(DateTime date)
        {
            DateTime d = date.Date;
            if (FromDate.HasValue && d < FromDate.Value.Date)
                return false;
            if (ToDate.HasValue && d > ToDate.Value.Date)
                return false;
            return true;
        }

        public bool MatchesMatchedTrade(MatchedTradeRow row)
        {
            if (row == null)
                return false;

            if (!string.IsNullOrWhiteSpace(AccountName)
                && !string.Equals(row.AccountName, AccountName, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasInstrumentFilter && !string.Equals(row.Symbol, Instrument, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasSourceFilter && !string.Equals(row.Source, Source, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasSideFilter && !string.Equals(row.Direction, Side, StringComparison.OrdinalIgnoreCase))
                return false;

            if (HasResultFilter)
            {
                string normalized = (Result ?? string.Empty).Trim().ToLowerInvariant();
                if (normalized == "winning" && row.Pnl <= 0)
                    return false;
                if (normalized == "losing" && row.Pnl >= 0)
                    return false;
                if (normalized == "breakeven" && Math.Abs(row.Pnl) > 0.000001)
                    return false;
                if (normalized == "open")
                    return false;
            }

            return MatchesDate(TradeSessionCalendar_v3.ToTradingDate(row.ExitTime));
        }
    }

    public class CsvParseResult
    {
        public int RowsRead { get; set; }
        public int RowsSkipped { get; set; }
        public List<ImportedTradeRecord> Trades { get; private set; }
        public List<string> Errors { get; private set; }

        public CsvParseResult()
        {
            Trades = new List<ImportedTradeRecord>();
            Errors = new List<string>();
        }
    }


    public class LedgerExportResult
    {
        public string OutputPath { get; set; }
        public int LiveRowsExported { get; set; }
        public int ImportedRowsExported { get; set; }
        public List<string> Errors { get; private set; }

        public LedgerExportResult()
        {
            Errors = new List<string>();
        }

        public string ToDisplayText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Export complete.");
            sb.AppendLine();
            sb.AppendLine("Live ledger rows: " + LiveRowsExported.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Imported ledger rows: " + ImportedRowsExported.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Total rows: " + (LiveRowsExported + ImportedRowsExported).ToString(CultureInfo.InvariantCulture));

            if (!string.IsNullOrWhiteSpace(OutputPath))
            {
                sb.AppendLine();
                sb.AppendLine("Saved to:");
                sb.AppendLine(OutputPath);
            }

            if (Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Errors:");
                foreach (string error in Errors.Take(20))
                    sb.AppendLine("- " + error);
            }

            return sb.ToString().TrimEnd();
        }
    }

    public class CsvImportResult
    {
        public int FilesProcessed { get; set; }
        public int RowsRead { get; set; }
        public int RowsImported { get; set; }
        public int RowsDuplicate { get; set; }
        public int RowsSkipped { get; set; }
        public List<string> Errors { get; private set; }

        public CsvImportResult()
        {
            Errors = new List<string>();
        }

        public string ToDisplayText()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Import complete.");
            sb.AppendLine();
            sb.AppendLine("Files processed: " + FilesProcessed.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Rows read: " + RowsRead.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Rows imported: " + RowsImported.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Duplicates skipped: " + RowsDuplicate.ToString(CultureInfo.InvariantCulture));
            sb.AppendLine("Rows skipped: " + RowsSkipped.ToString(CultureInfo.InvariantCulture));

            if (Errors.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Notes / skipped-row reasons:");
                foreach (string err in Errors.Take(12))
                    sb.AppendLine("- " + err);

                if (Errors.Count > 12)
                    sb.AppendLine("- ... " + (Errors.Count - 12).ToString(CultureInfo.InvariantCulture) + " more");
            }

            return sb.ToString();
        }
    }
}
