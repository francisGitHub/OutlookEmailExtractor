using System.Windows;
using Ninject;
using OutlookEmailExtractor.Services;
using OutlookEmailExtractor.Services.Impl;

namespace OutlookEmailExtractor
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private IKernel _container;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            ConfigureContainer();
            ComposeObjects();

            Current.MainWindow.Show();
        }

        private void ConfigureContainer()
        {
            _container = new StandardKernel();

            _container.Bind<IExtractEmails>().To<EmailExtractionService>();
        }

        private void ComposeObjects()
        {
            Current.MainWindow = this._container.Get<MainWindow>();
        }
    }
}