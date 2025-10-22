using System.Configuration;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using EmailCompleteApp.Services;
using EmailCompleteApp.Windows;

namespace EmailCompleteApp
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            base.OnStartup(e);

            ShutdownMode = ShutdownMode.OnExplicitShutdown;

            var loadingWindow = new LoadingWindow();
            loadingWindow.Show();

            try
            {
                var searchService = SearchService.Instance;
                searchService.ProgressChanged += message => loadingWindow.UpdateProgress(message);
                searchService.DetailChanged += detail => loadingWindow.UpdateDetail(detail);

                await searchService.LoadAllDataAsync();
                await Task.Delay(300); 

                var mainWindow = new MainWindow();
                MainWindow = mainWindow;
                mainWindow.Show();

                loadingWindow.Close();

                ShutdownMode = ShutdownMode.OnLastWindowClose;
            }
            catch (System.Exception ex)
            {
                loadingWindow.Close();
                MessageBox.Show($"Failed to initialize application: {ex.Message}",
                               "Startup Error",
                               MessageBoxButton.OK,
                               MessageBoxImage.Error);
                Shutdown();
            }
        }
    }
}
