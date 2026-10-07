using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using TankManager.Core.Services;

namespace TankManager.Views
{
    /// <summary>
    /// Журнал изменений с фильтрами по изделию, сотруднику, действию и месяцу
    /// </summary>
    public partial class AuditLogDialog : Window
    {
        private const string AllKey = "";

        private readonly Func<List<AuditEntry>> _loader;
        private List<AuditEntry> _entries = new List<AuditEntry>();
        private bool _updatingFilters;

        /// <param name="loader">Чтение журнала (вызывается вне UI-потока)</param>
        /// <param name="productFilter">Начальный фильтр по изделию (null — все)</param>
        public AuditLogDialog(Func<List<AuditEntry>> loader, string productFilter = null)
        {
            InitializeComponent();
            _loader = loader;

            _updatingFilters = true;
            ProductFilterBox.Text = productFilter ?? string.Empty;
            ActionFilter.ItemsSource = new[] { new KeyValuePair<string, string>(AllKey, "Все") }
                .Concat(Enum.GetValues(typeof(AuditAction)).Cast<AuditAction>()
                    .Select(a => new KeyValuePair<string, string>(a.ToString(), AuditEntry.GetActionTitle(a))))
                .ToList();
            ActionFilter.SelectedIndex = 0;
            _updatingFilters = false;

            Loaded += async (s, e) => await LoadAsync();
        }

        private async Task LoadAsync()
        {
            StatusText.Text = "Загрузка журнала...";
            StatusText.Visibility = Visibility.Visible;
            EntriesGrid.ItemsSource = null;

            try
            {
                _entries = await Task.Run(_loader) ?? new List<AuditEntry>();
            }
            catch (Exception ex)
            {
                _entries = new List<AuditEntry>();
                StatusText.Text = $"Не удалось прочитать журнал: {ex.Message}";
                return;
            }

            FillFilterOptions();
            ApplyFilter();
        }

        private void FillFilterOptions()
        {
            _updatingFilters = true;

            string selectedUser = UserFilter.SelectedValue as string;
            UserFilter.ItemsSource = new[] { new KeyValuePair<string, string>(AllKey, "Все") }
                .Concat(_entries
                    .GroupBy(e => e.Login ?? "", StringComparer.OrdinalIgnoreCase)
                    .Select(g => new KeyValuePair<string, string>(g.Key, g.First().UserDisplayName))
                    .OrderBy(p => p.Value, StringComparer.CurrentCultureIgnoreCase))
                .ToList();
            UserFilter.SelectedValue = selectedUser ?? AllKey;
            if (UserFilter.SelectedIndex < 0) UserFilter.SelectedIndex = 0;

            string selectedPeriod = PeriodFilter.SelectedValue as string;
            PeriodFilter.ItemsSource = new[] { new KeyValuePair<string, string>(AllKey, "Всё время") }
                .Concat(_entries
                    .Select(e => e.LocalTime.ToString("yyyy-MM"))
                    .Distinct()
                    .OrderByDescending(m => m)
                    .Select(m => new KeyValuePair<string, string>(m, DateTime.ParseExact(m, "yyyy-MM", null).ToString("MMMM yyyy"))))
                .ToList();
            PeriodFilter.SelectedValue = selectedPeriod ?? AllKey;
            if (PeriodFilter.SelectedIndex < 0) PeriodFilter.SelectedIndex = 0;

            _updatingFilters = false;
        }

        private void ApplyFilter()
        {
            string product = ProductFilterBox.Text?.Trim();
            string user = UserFilter.SelectedValue as string;
            string action = ActionFilter.SelectedValue as string;
            string period = PeriodFilter.SelectedValue as string;

            var filtered = _entries.Where(e =>
                    (string.IsNullOrEmpty(product) ||
                     (e.ProductName?.IndexOf(product, StringComparison.CurrentCultureIgnoreCase) >= 0) ||
                     (e.Marking?.IndexOf(product, StringComparison.CurrentCultureIgnoreCase) >= 0)) &&
                    (string.IsNullOrEmpty(user) || string.Equals(e.Login, user, StringComparison.OrdinalIgnoreCase)) &&
                    (string.IsNullOrEmpty(action) || e.Action.ToString() == action) &&
                    (string.IsNullOrEmpty(period) || e.LocalTime.ToString("yyyy-MM") == period))
                .ToList();

            EntriesGrid.ItemsSource = filtered;
            CountText.Text = filtered.Count == _entries.Count
                ? $"Записей: {_entries.Count}"
                : $"Показано {filtered.Count} из {_entries.Count}";

            if (_entries.Count == 0)
            {
                StatusText.Text = "Журнал пуст";
                StatusText.Visibility = Visibility.Visible;
            }
            else if (filtered.Count == 0)
            {
                StatusText.Text = "Нет записей по выбранным фильтрам";
                StatusText.Visibility = Visibility.Visible;
            }
            else
            {
                StatusText.Visibility = Visibility.Collapsed;
            }
        }

        private void Filter_Changed(object sender, RoutedEventArgs e)
        {
            if (!_updatingFilters && IsLoaded)
                ApplyFilter();
        }

        private void EntriesGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!(EntriesGrid.SelectedItem is AuditEntry entry))
            {
                DetailsText.Text = "Выберите запись, чтобы увидеть, что изменилось";
                return;
            }

            var header = $"{entry.LocalTime:dd.MM.yyyy HH:mm} — {entry.UserDisplayName} ({entry.Login}) — {entry.ActionTitle}";
            DetailsText.Text = string.IsNullOrEmpty(entry.Details)
                ? header
                : header + Environment.NewLine + Environment.NewLine + entry.Details.Replace("\n", Environment.NewLine);
        }

        private async void RefreshButton_Click(object sender, RoutedEventArgs e) => await LoadAsync();

        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();
    }
}
