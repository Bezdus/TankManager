using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Markup;
using TankManager.Core.Services;

namespace TankManager
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static readonly ILogger Logger = new FileLogger();

        protected override void OnStartup(StartupEventArgs e)
        {
            // Привязки WPF форматируют числа по Language элемента (по умолчанию en-US).
            // Культуру потока не меняем: от неё зависят выгрузка в Excel и сохранение данных.
            FrameworkElement.LanguageProperty.OverrideMetadata(
                typeof(FrameworkElement),
                new FrameworkPropertyMetadata(XmlLanguage.GetLanguage("ru-RU")));

            base.OnStartup(e);
            
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += TaskScheduler_UnobservedTaskException;

            // Проверка обновлений при запуске приложения
            UpdateService.CheckForUpdates(showNoUpdateMessage: false);

            try
            {
                var mainWindow = new MainWindow();
                this.MainWindow = mainWindow;
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                Logger.LogError("Критическая ошибка при запуске приложения", ex);
                MessageBox.Show(
                    $"Критическая ошибка при запуске приложения:\n\n{ex.Message}",
                    "Ошибка запуска", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
                Shutdown();
            }
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Logger.LogError("Необработанное исключение в UI-потоке", e.Exception);
            MessageBox.Show(
                $"Необработанное исключение:\n\n{e.Exception.Message}",
                "Ошибка", 
                MessageBoxButton.OK, 
                MessageBoxImage.Error);
            
            e.Handled = true;
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            Logger.LogError($"Критическая ошибка (завершение: {e.IsTerminating})", exception);
            MessageBox.Show(
                $"Критическая ошибка:\n\n{exception?.Message}",
                "Критическая ошибка", 
                MessageBoxButton.OK, 
                MessageBoxImage.Error);
        }

        // Исключение в задаче, результат которой никто не ждал: без обработчика оно пропадает бесследно
        private void TaskScheduler_UnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            Logger.LogError("Необработанное исключение в фоновой задаче", e.Exception?.GetBaseException());
            e.SetObserved();
        }
    }
}
