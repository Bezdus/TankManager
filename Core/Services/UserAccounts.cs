using System;
using System.Collections.Generic;
using System.DirectoryServices.AccountManagement;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Threading.Tasks;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Роль сотрудника; определяет режим работы программы
    /// </summary>
    public enum UserRole
    {
        /// <summary>Просмотр: только чтение изделий с сервера</summary>
        Viewer,
        /// <summary>Технолог: просмотр + правка операций изготовления</summary>
        Technologist,
        /// <summary>Конструктор: загрузка из КОМПАС, сохранение, удаление, расценки</summary>
        Engineer
    }

    /// <summary>
    /// Сотрудник из общего списка (_users.json в серверной папке)
    /// </summary>
    [DataContract]
    public class UserAccount
    {
        /// <summary>
        /// Логин Windows без домена
        /// </summary>
        [DataMember]
        public string Login { get; set; }

        /// <summary>
        /// ФИО (из домена при регистрации, может править администратор)
        /// </summary>
        [DataMember]
        public string Name { get; set; }

        [DataMember(Name = "Role")]
        private string RoleText
        {
            get => Role.ToString();
            set => Role = Enum.TryParse(value, true, out UserRole role) ? role : UserRole.Viewer;
        }

        public UserRole Role { get; set; }

        /// <summary>
        /// Администратор: ведёт список сотрудников
        /// </summary>
        [DataMember]
        public bool IsAdmin { get; set; }

        /// <summary>
        /// Отключён (уволен): работает только просмотр
        /// </summary>
        [DataMember]
        public bool IsDisabled { get; set; }

        [DataMember]
        public long AddedUtcTicks { get; set; }

        public string DisplayName => string.IsNullOrWhiteSpace(Name) ? Login : Name;

        public DateTime AddedLocal => AddedUtcTicks > 0
            ? new DateTime(AddedUtcTicks, DateTimeKind.Utc).ToLocalTime()
            : DateTime.MinValue;

        public UserAccount Clone() => (UserAccount)MemberwiseClone();

        public static string RoleTitle(UserRole role)
        {
            switch (role)
            {
                case UserRole.Engineer: return "Конструктор";
                case UserRole.Technologist: return "Технолог";
                default: return "Просмотр";
            }
        }
    }

    [DataContract]
    public class UsersFile
    {
        [DataMember]
        public List<UserAccount> Users { get; set; }
    }

    /// <summary>
    /// Текущий сотрудник. Определяется один раз при запуске (<see cref="Initialize"/>), как и <see cref="AppMode"/>
    /// </summary>
    public static class CurrentUser
    {
        /// <summary>
        /// Логин Windows (без домена); Windows уже проверила пароль при входе
        /// </summary>
        public static string Login { get; } = Environment.UserName;

        /// <summary>
        /// ФИО из списка сотрудников (null — ещё не известно)
        /// </summary>
        public static string Name { get; private set; }

        public static string DisplayName => string.IsNullOrWhiteSpace(Name) ? Login : Name;

        public static UserRole Role { get; private set; } = UserRole.Viewer;

        public static bool IsAdmin { get; private set; }

        /// <summary>
        /// Список сотрудников есть (на сервере или в кэше): роль берётся из него, а не из переключателя режима
        /// </summary>
        public static bool AccountsEnabled { get; private set; }

        /// <summary>
        /// Сотрудника нет в списке: его нужно зарегистрировать
        /// </summary>
        public static bool NeedsRegistration { get; private set; }

        /// <summary>
        /// Определяет сотрудника и настройку режима, с которой нужно инициализировать <see cref="AppMode"/>
        /// </summary>
        /// <param name="users">Список сотрудников (null — списка нет, работаем по-старому)</param>
        /// <param name="admins">Логины, которые всегда администраторы (из storage_settings)</param>
        /// <param name="fallback">Настройка режима для работы без списка и для ещё не зарегистрированных</param>
        public static AppModeSetting Initialize(List<UserAccount> users, IEnumerable<string> admins, AppModeSetting fallback)
        {
            bool isBootstrapAdmin = admins != null && admins.Any(a => string.Equals(a?.Trim(), Login, StringComparison.OrdinalIgnoreCase));

            AccountsEnabled = users != null;
            NeedsRegistration = false;
            IsAdmin = isBootstrapAdmin;
            Name = null;

            if (users == null)
                return fallback;

            var account = UserDirectoryService.Find(users, Login);
            if (account == null)
            {
                // Новый сотрудник работает как раньше (по переключателю режима), пока администратор не назначит роль
                NeedsRegistration = true;
                if (!isBootstrapAdmin)
                    return fallback;

                Role = UserRole.Engineer;
                return AppModeSetting.Engineer;
            }

            Name = account.Name;
            IsAdmin = isBootstrapAdmin || (account.IsAdmin && !account.IsDisabled);
            Role = isBootstrapAdmin ? UserRole.Engineer
                : account.IsDisabled ? UserRole.Viewer
                : account.Role;

            return ToModeSetting(Role);
        }

        /// <summary>
        /// После регистрации (ФИО из домена приходит в фоне)
        /// </summary>
        public static void SetRegistered(UserAccount account)
        {
            if (account == null) return;
            Name = account.Name;
            NeedsRegistration = false;
        }

        /// <summary>
        /// Роль, с которой регистрируется новый сотрудник: режим, в котором он сейчас работает
        /// </summary>
        public static UserRole RoleFromCurrentMode()
        {
            if (AppMode.IsTechnologist) return UserRole.Technologist;
            return AppMode.IsViewer ? UserRole.Viewer : UserRole.Engineer;
        }

        public static AppModeSetting ToModeSetting(UserRole role)
        {
            switch (role)
            {
                case UserRole.Engineer: return AppModeSetting.Engineer;
                case UserRole.Technologist: return AppModeSetting.Technologist;
                default: return AppModeSetting.Viewer;
            }
        }
    }

    /// <summary>
    /// Общий список сотрудников: &lt;серверная папка&gt;\_users.json, локальная копия users_cache.json рядом с exe
    /// </summary>
    public class UserDirectoryService
    {
        public const string ServerFileName = "_users.json";

        private static readonly string CachePath =
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "users_cache.json");

        private static readonly TimeSpan StartupReadTimeout = TimeSpan.FromSeconds(3);
        private static readonly TimeSpan DomainLookupTimeout = TimeSpan.FromSeconds(5);
        private static readonly object FileLock = new object();

        private readonly ILogger _logger = new FileLogger();
        private readonly string _serverFolder;

        public UserDirectoryService(string serverFolder)
        {
            _serverFolder = serverFolder;
        }

        private string ServerFilePath =>
            string.IsNullOrEmpty(_serverFolder) ? null : Path.Combine(_serverFolder, ServerFileName);

        public static UserAccount Find(IEnumerable<UserAccount> users, string login)
        {
            return users?.FirstOrDefault(u => string.Equals(u?.Login, login, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Список для определения роли при запуске: с сервера (не дольше нескольких секунд), иначе из кэша.
        /// null — списка нет: на сервере он ещё не создан и кэша нет
        /// </summary>
        public List<UserAccount> LoadForStartup()
        {
            string serverPath = ServerFilePath;
            if (serverPath != null)
            {
                try
                {
                    var task = Task.Run(() =>
                    {
                        if (!Directory.Exists(_serverFolder))
                            return Tuple.Create(false, (List<UserAccount>)null);

                        // Папка доступна, файла нет — списка ещё нет (кэш тоже устарел)
                        if (!File.Exists(serverPath))
                            return Tuple.Create(true, (List<UserAccount>)null);

                        return Tuple.Create(true, ReadFile(serverPath));
                    });

                    if (task.Wait(StartupReadTimeout))
                    {
                        bool reached = task.Result.Item1;
                        var users = task.Result.Item2;
                        if (reached && users != null)
                        {
                            WriteCache(users);
                            return users;
                        }

                        if (reached)
                            return null;
                    }
                    else
                    {
                        _logger.LogWarning("Список сотрудников: сервер не ответил вовремя, используется локальная копия");
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning($"Не удалось прочитать список сотрудников с сервера: {ex.GetBaseException().Message}");
                }
            }

            try
            {
                return File.Exists(CachePath) ? ReadFile(CachePath) : null;
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось прочитать локальную копию списка сотрудников: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Свежий список с сервера (для окна «Сотрудники»). Нет файла — пустой список
        /// </summary>
        public List<UserAccount> ReadServer()
        {
            string serverPath = ServerFilePath;
            if (serverPath == null)
                throw new InvalidOperationException("Серверная папка не задана");

            lock (FileLock)
            {
                var users = File.Exists(serverPath) ? ReadFile(serverPath) : new List<UserAccount>();
                WriteCache(users);
                return users;
            }
        }

        /// <summary>
        /// Регистрирует текущего сотрудника, если его ещё нет в списке на сервере. ФИО берётся из домена.
        /// Возвращает запись сотрудника (null — сервер недоступен); created — запись добавлена сейчас
        /// </summary>
        public UserAccount Register(string login, UserRole role, bool isAdmin, out bool created)
        {
            created = false;
            string serverPath = ServerFilePath;
            if (serverPath == null || !Directory.Exists(_serverFolder))
                return null;

            // Запрос к домену — до блокировки файла: он может быть медленным
            string name = LookupDomainDisplayName();

            lock (FileLock)
            {
                var users = File.Exists(serverPath) ? ReadFile(serverPath) : new List<UserAccount>();
                var existing = Find(users, login);
                if (existing != null)
                {
                    WriteCache(users);
                    return existing;
                }

                var account = new UserAccount
                {
                    Login = login,
                    Name = name,
                    Role = role,
                    IsAdmin = isAdmin,
                    AddedUtcTicks = DateTime.UtcNow.Ticks
                };
                users.Add(account);
                WriteFile(serverPath, users);
                WriteCache(users);
                created = true;
                return account;
            }
        }

        /// <summary>
        /// Сохраняет изменения администратора: перечитывает файл и накладывает правки по логину,
        /// чтобы не потерять сотрудников, зарегистрировавшихся, пока окно было открыто
        /// </summary>
        public List<UserAccount> SaveChanges(IEnumerable<UserAccount> changed)
        {
            string serverPath = ServerFilePath;
            if (serverPath == null)
                throw new InvalidOperationException("Серверная папка не задана");

            lock (FileLock)
            {
                var users = File.Exists(serverPath) ? ReadFile(serverPath) : new List<UserAccount>();
                foreach (var edit in changed)
                {
                    var existing = Find(users, edit.Login);
                    if (existing == null)
                    {
                        users.Add(edit.Clone());
                        continue;
                    }

                    existing.Name = edit.Name;
                    existing.Role = edit.Role;
                    existing.IsAdmin = edit.IsAdmin;
                    existing.IsDisabled = edit.IsDisabled;
                }

                WriteFile(serverPath, users);
                WriteCache(users);
                return users;
            }
        }

        /// <summary>
        /// ФИО из Active Directory (только чтение, от имени текущего пользователя).
        /// Если домен недоступен или не ответил вовремя — null
        /// </summary>
        public string LookupDomainDisplayName()
        {
            try
            {
                var task = Task.Run(() =>
                {
                    using (var principal = UserPrincipal.Current)
                        return principal?.DisplayName;
                });

                if (task.Wait(DomainLookupTimeout))
                {
                    string name = task.Result?.Trim();
                    _logger.LogInfo($"ФИО из домена для {CurrentUser.Login}: {(string.IsNullOrEmpty(name) ? "не задано" : name)}");
                    return string.IsNullOrEmpty(name) ? null : name;
                }

                _logger.LogWarning("Домен не ответил вовремя, ФИО не получено");
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось получить ФИО из домена: {ex.GetBaseException().Message}");
            }

            return null;
        }

        private static List<UserAccount> ReadFile(string path)
        {
            var serializer = new DataContractJsonSerializer(typeof(UsersFile));
            using (var stream = new MemoryStream(File.ReadAllBytes(path)))
            {
                var file = (UsersFile)serializer.ReadObject(stream);
                return (file?.Users ?? new List<UserAccount>())
                    .Where(u => !string.IsNullOrWhiteSpace(u?.Login))
                    .ToList();
            }
        }

        private static void WriteFile(string path, List<UserAccount> users)
        {
            var serializer = new DataContractJsonSerializer(typeof(UsersFile));
            using (var stream = new MemoryStream())
            {
                var sorted = users.OrderBy(u => u.Login, StringComparer.OrdinalIgnoreCase).ToList();
                serializer.WriteObject(stream, new UsersFile { Users = sorted });
                AtomicFile.WriteAllBytes(path, stream.ToArray());
            }
        }

        private void WriteCache(List<UserAccount> users)
        {
            try
            {
                WriteFile(CachePath, users);
            }
            catch (Exception ex)
            {
                _logger.LogWarning($"Не удалось обновить локальную копию списка сотрудников: {ex.Message}");
            }
        }
    }
}
