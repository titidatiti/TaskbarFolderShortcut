using System.Windows;
using TrayFolder.Services;
using System.Windows.Forms; // For FolderBrowserDialog

namespace TrayFolder.Views
{
    public partial class SettingsWindow : Window
    {
        public string SelectedPath { get; private set; }

        public SettingsWindow(string currentPath)
        {
            InitializeComponent();
            SelectedPath = currentPath;
            PathTextBox.Text = SelectedPath;
        }

        private void Browse_Click(object sender, RoutedEventArgs e)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.SelectedPath = SelectedPath;
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SelectedPath = dialog.SelectedPath;
                    PathTextBox.Text = SelectedPath;
                }
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
