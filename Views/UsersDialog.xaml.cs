using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using TankManager.Core.Services;

namespace TankManager.Views
{
    /// <summary>
    /// Список сотрудников: роли, администраторы, отключённые (для администратора)
    /// </summary>
    public partial class UsersDialog : Window
    {
        private readonly Dictionary<string, UserAccount> _original;
        private readonly ICollection<string> _bootstrapAdmins;

        public ObservableCollection<UserAccount> Users { get; }

        public IReadOnlyList<KeyValuePair<UserRole, string>> RoleOptions { get; } = new[]
        {
            new KeyValuePair<UserRole, string>(UserRole.Engineer, UserAccount.RoleTitle(UserRole.Engineer)),
            new KeyValuePair<UserRole, string>(UserRole.Technologist, UserAccount.RoleTitle(UserRole.Technologist)),
            new KeyValuePair<UserRole, string>(UserRole.Viewer, UserAccount.RoleTitle(UserRole.Viewer))
        };

        /// <summary>
        /// Изменённые записи (после «Сохранить»)
        /// </summary>
        public List<UserAccount> ChangedUsers { get; private set; } = new List<UserAccount>();

        /// <summary>
        /// Описание изменений для журнала
        /// </summary>
        public string ChangesDescription { get; private set; }

        public UsersDialog(IEnumerable<UserAccount> users, ICollection<string> bootstrapAdmins)
        {
            InitializeComponent();

            var list = users.ToList();
            _original = list.ToDictionary(u => u.Login, u => u.Clone(), StringComparer.OrdinalIgnoreCase);
            _bootstrapAdmins = bootstrapAdmins ?? new List<string>();

            Users = new ObservableCollection<UserAccount>(list
                .Select(u => u.Clone())
                .OrderBy(u => u.IsDisabled)
                .ThenBy(u => u.DisplayName, StringComparer.CurrentCultureIgnoreCase));

            UsersGrid.ItemsSource = Users;
            CountText.Text = $"Сотрудников: {Users.Count}";
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            UsersGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            UsersGrid.CommitEdit(DataGridEditingUnit.Row, true);

            bool hasAdmin = _bootstrapAdmins.Count > 0 || Users.Any(u => u.IsAdmin && !u.IsDisabled);
            if (!hasAdmin)
            {
                MessageBox.Show(this, "Должен остаться хотя бы один администратор.", "Сотрудники",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var lines = new List<string>();
            foreach (var user in Users)
            {
                user.Name = user.Name?.Trim();
                if (!_original.TryGetValue(user.Login, out var was))
                    continue;

                var diffs = new List<string>();
                if (!string.Equals(was.Name ?? "", user.Name ?? "", StringComparison.Ordinal))
                    diffs.Add($"ФИО «{was.Name}» → «{user.Name}»");
                if (was.Role != user.Role)
                    diffs.Add($"роль {UserAccount.RoleTitle(was.Role)} → {UserAccount.RoleTitle(user.Role)}");
                if (was.IsAdmin != user.IsAdmin)
                    diffs.Add(user.IsAdmin ? "назначен администратором" : "снят с администраторов");
                if (was.IsDisabled != user.IsDisabled)
                    diffs.Add(user.IsDisabled ? "отключён" : "включён");

                if (diffs.Count == 0)
                    continue;

                ChangedUsers.Add(user);
                lines.Add($"{user.Login} ({user.DisplayName}): {string.Join(", ", diffs)}");
            }

            ChangesDescription = lines.Count == 0 ? null : string.Join("\n", lines);
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
