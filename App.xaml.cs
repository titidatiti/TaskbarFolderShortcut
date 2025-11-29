using System;
using System.Windows;
using System.Windows.Forms; // For NotifyIcon
using TaskbarFolderShortcut.Services;
using TaskbarFolderShortcut.Views;
using Application = System.Windows.Application;

namespace TaskbarFolderShortcut
{
    public partial class App : Application
    {
        private NotifyIcon _notifyIcon;
        private MainWindow _mainWindow;

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            try
            {
                // Initialize MainWindow (hidden)
                _mainWindow = new MainWindow();
                _mainWindow.LoadRootFolder();
                
                // Initialize Tray Icon
                _notifyIcon = new NotifyIcon
                {
                    Text = "Taskbar Folder Shortcut"
                };

                try
                {
                    using (var stream = Application.GetResourceStream(new Uri("pack://application:,,,/app.ico")).Stream)
                    {
                        _notifyIcon.Icon = new System.Drawing.Icon(stream);
                    }
                }
                catch
                {
                    _notifyIcon.Icon = System.Drawing.SystemIcons.Application;
                }

                _notifyIcon.Visible = true;

                _notifyIcon.MouseClick += NotifyIcon_MouseClick;

                // Context Menu for Tray Icon (Right Click)
                var contextMenu = new ContextMenuStrip();
                
                // Add App Name Header
                var header = new ToolStripMenuItem("Taskbar Folder Shortcut");
                header.Enabled = false; 
                contextMenu.Items.Add(header);
                contextMenu.Items.Add("-"); // Separator

                // Settings
                var settingsItem = new ToolStripMenuItem("Settings", Helpers.FontIconHelper.CreateIcon("\uE713"), (s, args) => OpenSettings());
                contextMenu.Items.Add(settingsItem);

                // Refresh
                var refreshItem = new ToolStripMenuItem("Refresh", Helpers.FontIconHelper.CreateIcon("\uE72C"), (s, args) => RefreshFolders());
                contextMenu.Items.Add(refreshItem);

                contextMenu.Items.Add("-"); // Separator

                // Exit
                var exitItem = new ToolStripMenuItem("Exit", Helpers.FontIconHelper.CreateIcon("\uE7E8"), (s, args) => Shutdown());
                contextMenu.Items.Add(exitItem);

                _notifyIcon.ContextMenuStrip = contextMenu;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Startup Error: {ex.Message}\n{ex.StackTrace}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown();
            }
        }

        private void NotifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                _mainWindow.ShowFolderMenu();
            }
        }

        private void OpenSettings()
        {
            var currentSettings = SettingsService.Load();
            var settingsWindow = new SettingsWindow(currentSettings.RootFolderPath);
            if (settingsWindow.ShowDialog() == true)
            {
                currentSettings.RootFolderPath = settingsWindow.SelectedPath;
                SettingsService.Save(currentSettings);
                _mainWindow.LoadRootFolder(); // Reload the menu
            }
        }

        private void RefreshFolders()
        {
            _mainWindow.LoadRootFolder();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            if (_notifyIcon != null) _notifyIcon.Dispose();
            base.OnExit(e);
        }
    }
}
