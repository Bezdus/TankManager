using System;
using System.Threading.Tasks;
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
                // Конфигурация AutoUpdater
                AutoUpdater.ShowSkipButton = false;
                AutoUpdater.ShowRemindLaterButton = false;
                AutoUpdater.Mandatory = false;
                AutoUpdater.RunUpdateAsAdmin = false;
                AutoUpdater.ReportErrors = showNoUpdateMessage;
                
                // Обработка ошибок
                AutoUpdater.CheckForUpdateEvent += AutoUpdaterOnCheckForUpdateEvent;
                
                // Настройки UI
                AutoUpdater.ApplicationExitEvent += AutoUpdater_ApplicationExitEvent;
                
                // Запуск проверки обновлений
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

        /// <summary>
        /// Асинхронная проверка обновлений
        /// </summary>
        public static async Task CheckForUpdatesAsync(bool showNoUpdateMessage = false)
        {
            await Task.Run(() => CheckForUpdates(showNoUpdateMessage));
        }

        private static void AutoUpdaterOnCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            if (args.Error == null)
            {
                if (args.IsUpdateAvailable)
                {
                    // Обновление доступно - AutoUpdater покажет диалог автоматически
                }
                else
                {
                    // Обновлений нет
                }
            }
            else
            {
                if (args.Error is System.Net.WebException)
                {
                    // Проблема с сетью - тихо игнорируем при автоматической проверке
                    if (AutoUpdater.ReportErrors)
                    {
                        MessageBox.Show(
                            "Не удалось проверить обновления.\nПроверьте подключение к интернету или повторите попытку позже.",
                            "Ошибка подключения",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
                else
                {
                    if (AutoUpdater.ReportErrors)
                    {
                        MessageBox.Show(
                            $"Ошибка при проверке обновлений:\n{args.Error.Message}",
                            "Ошибка",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    }
                }
            }
        }

        private static void AutoUpdater_ApplicationExitEvent()
        {
            System.Windows.Application.Current?.Shutdown();
        }
    }
}
