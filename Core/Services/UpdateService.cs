using System;
using System.Windows;
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
            try
            {
                AutoUpdater.ShowSkipButton = false;
                AutoUpdater.ShowRemindLaterButton = false;
                AutoUpdater.Mandatory = false;
                AutoUpdater.RunUpdateAsAdmin = false;
                AutoUpdater.ReportErrors = showNoUpdateMessage;
                AutoUpdater.ApplicationExitEvent -= OnApplicationExit;
                AutoUpdater.ApplicationExitEvent += OnApplicationExit;

                AutoUpdater.Start(UPDATE_URL);
            }
            catch (Exception ex)
            {
                if (showNoUpdateMessage)
                {
                    MessageBox.Show(
                        $"Не удалось проверить обновления:\n{ex.Message}",
                        "Ошибка проверки обновлений",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
        }

        private static void OnApplicationExit()
        {
            System.Windows.Application.Current?.Shutdown();
        }
    }
}
