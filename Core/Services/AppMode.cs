using System;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Настройка режима работы приложения
    /// </summary>
    public enum AppModeSetting
    {
        /// <summary>Определять по наличию КОМПАС-3D на компьютере</summary>
        Auto,
        /// <summary>Конструктор: загрузка сборок из КОМПАС, сохранение, удаление</summary>
        Engineer,
        /// <summary>Просмотр: только чтение изделий с сервера</summary>
        Viewer
    }

    /// <summary>
    /// Режим работы приложения. Определяется один раз при запуске (<see cref="Initialize"/>):
    /// смена настройки применяется после перезапуска.
    /// </summary>
    public static class AppMode
    {
        private const string KompasProgId = "KOMPAS.Application.7";

        /// <summary>
        /// Настройка, с которой запущено приложение
        /// </summary>
        public static AppModeSetting Setting { get; private set; } = AppModeSetting.Auto;

        /// <summary>
        /// Режим просмотра: без КОМПАС, без записи на сервер
        /// </summary>
        public static bool IsViewer { get; private set; }

        /// <summary>
        /// Зарегистрирован ли КОМПАС-3D в системе
        /// </summary>
        public static bool IsKompasInstalled { get; private set; }

        public static void Initialize(AppModeSetting setting)
        {
            Setting = setting;
            IsKompasInstalled = DetectKompas();

            switch (setting)
            {
                case AppModeSetting.Engineer:
                    IsViewer = false;
                    break;
                case AppModeSetting.Viewer:
                    IsViewer = true;
                    break;
                default:
                    IsViewer = !IsKompasInstalled;
                    break;
            }
        }

        private static bool DetectKompas()
        {
            try
            {
                return Type.GetTypeFromProgID(KompasProgId, false) != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
