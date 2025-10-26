using System;
using System.Windows;
using TankManager.Views;

namespace TankManager
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Временно устанавливаем OnExplicitShutdown, чтобы предотвратить 
            // преждевременное завершение приложения до инициализации главного окна
            this.ShutdownMode = ShutdownMode.OnExplicitShutdown;
            
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;

            try
            {
                var loadingScreen = new LoadingScreen();
                var dialogResult = loadingScreen.ShowDialog();
                
                if (dialogResult == true)
                {
                    var mainWindow = new MainWindow();
                    this.MainWindow = mainWindow;
                    
                    // После установки главного окна переключаемся на стандартный режим
                    this.ShutdownMode = ShutdownMode.OnMainWindowClose;
                    
                    // Передаем параметры загрузки
                    if (loadingScreen.LoadFromActiveDocument)
                    {
                        mainWindow.SetLoadingParameters(loadFromActiveDocument: true);
                    }
                    else if (!string.IsNullOrEmpty(loadingScreen.SelectedFilePath))
                    {
                        mainWindow.SetLoadingParameters(filePath: loadingScreen.SelectedFilePath);
                    }

                    mainWindow.Show();
                }
                else
                {
                    Shutdown();
                }
            }
            catch (Exception ex)
            {
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
            MessageBox.Show(
                $"Критическая ошибка:\n\n{exception?.Message}",
                "Критическая ошибка", 
                MessageBoxButton.OK, 
                MessageBoxImage.Error);
        }
    }
}
