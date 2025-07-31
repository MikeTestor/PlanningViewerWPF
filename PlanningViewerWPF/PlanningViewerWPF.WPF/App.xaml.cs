using System.Configuration;
using System.Data;
using System.Windows;

namespace PlanningViewerWPF.WPF
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            var mainWindow = new MainWindow();
            mainWindow.DataContext = new ViewModels.MainViewModel();
            mainWindow.Show();
        }
    }
}
