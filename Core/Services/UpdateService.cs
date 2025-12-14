using System;
using System.Threading.Tasks;
using AutoUpdaterDotNET;

namespace TankManager.Core.Services
{
    /// <summary>
    /// Сервис для проверки и установки обновлений приложения
    /// </summary>
    public class UpdateService
    {
        private const string UPDATE_URL = "https://raw.githubusercontent.com/Bezdus/TankManager/master/update.xml";

        /// <summary>
        /// Проверяет наличие обновлений
        /// </summary>
        /// <param name="showNoUpdateMessage">Показывать ли сообщение, если обновлений нет</param>
        public static void CheckForUpdates(bool showNoUpdateMessage = false)
        {
            // Конфигурация AutoUpdater
            AutoUpdater.ShowSkipButton = false;
            AutoUpdater.ShowRemindLaterButton = false;
            AutoUpdater.Mandatory = false;
            AutoUpdater.RunUpdateAsAdmin = false;
            AutoUpdater.ReportErrors = showNoUpdateMessage;
            
            // Настройки UI
            AutoUpdater.ApplicationExitEvent += AutoUpdater_ApplicationExitEvent;
            
            // Запуск проверки обновлений
            AutoUpdater.Start(UPDATE_URL);
        }

        /// <summary>
        /// Асинхронная проверка обновлений
        /// </summary>
        public static async Task CheckForUpdatesAsync(bool showNoUpdateMessage = false)
        {
            await Task.Run(() => CheckForUpdates(showNoUpdateMessage));
        }

        private static void AutoUpdater_ApplicationExitEvent()
        {
            System.Windows.Application.Current?.Shutdown();
        }
    }
}
