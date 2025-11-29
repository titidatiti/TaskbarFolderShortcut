using System.Windows;
using TrayFolder.Models;
using TrayFolder.Services;

namespace TrayFolder
{
    public partial class MainWindow : Window
    {
        private System.Windows.Controls.ContextMenu _rootContextMenu;

        public MainWindow()
        {
            InitializeComponent();
            _rootContextMenu = (System.Windows.Controls.ContextMenu)FindResource("RootContextMenu");
        }

        public void LoadRootFolder()
        {
            var settings = SettingsService.Load();
            var rootItem = new FileSystemItem(settings.RootFolderPath);
            rootItem.LoadChildren(); // Load first level
            _rootContextMenu.ItemsSource = rootItem.Children;
        }

        public void ShowFolderMenu()
        {
            this.Show();
            this.Activate();
            _rootContextMenu.IsOpen = true;
        }

        private void RootContextMenu_Closed(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }
    }
}