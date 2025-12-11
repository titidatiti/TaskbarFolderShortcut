using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;

using System.Diagnostics;
using System.Windows.Input;
using TrayFolder.Helpers;
using TrayFolder.Services;
using System.Windows;

namespace TrayFolder.Models
{
    public class FileSystemItem
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public bool IsFolder { get; set; }
        public ImageSource SysIcon { get; set; }
        public ObservableCollection<FileSystemItem> Children { get; set; }
        public ICommand OpenCommand { get; set; }
        public ICommand OpenInExplorerCommand { get; set; }
        public ICommand PropertiesCommand { get; set; }
        public ICommand CopyPathCommand { get; set; }

        public string TargetFolderPath { get; set; }

        private bool _isExpanded;
        public bool IsExpanded
        {
            get { return _isExpanded; }
            set
            {
                _isExpanded = value;
                if (_isExpanded && IsFolder && Children.Count == 1 && Children[0].Name == "Loading...")
                {
                    LoadChildren();
                }
            }
        }

        public FileSystemItem()
        {
            Children = new ObservableCollection<FileSystemItem>();
            InitializeCommands();
        }

        public FileSystemItem(string path)
        {
            Path = path;
            Name = System.IO.Path.GetFileName(path);
            if (string.IsNullOrEmpty(Name)) Name = path; // Handle root drives

            // Check if it's a shortcut to a folder
            if (System.IO.Path.GetExtension(path).ToLower() == ".lnk")
            {
                // Always hide extension for shortcuts
                Name = System.IO.Path.GetFileNameWithoutExtension(path);

                string target = ShortcutHelper.ResolveShortcut(path);
                if (!string.IsNullOrEmpty(target) && Directory.Exists(target))
                {
                    TargetFolderPath = target;
                    IsFolder = true;
                }
                else
                {
                    IsFolder = false;
                }
            }
            else
            {
                IsFolder = Directory.Exists(path);
            }

            try
            {
                SysIcon = IconHelper.GetIcon(path, IsFolder);
            }
            catch { }

            Children = new ObservableCollection<FileSystemItem>();
            InitializeCommands();

            if (IsFolder)
            {
                // Add a dummy item to enable expansion if it's a folder
                Children.Add(new FileSystemItem { Name = "Loading..." });
            }
        }

        public ICommand ShellContextMenuCommand { get; set; }

        private void InitializeCommands()
        {
            OpenCommand = new RelayCommand(ExecuteOpen);
            OpenInExplorerCommand = new RelayCommand(ExecuteOpenInExplorer);
            PropertiesCommand = new RelayCommand(ExecuteProperties);
            CopyPathCommand = new RelayCommand(ExecuteCopyPath);
            ContextOpenCommand = new RelayCommand(ExecuteContextOpen);
            ShellContextMenuCommand = new RelayCommand(ExecuteShellContextMenu);
        }

        public ICommand ContextOpenCommand { get; set; }

        private void ExecuteContextOpen(object obj)
        {
            if (IsFolder)
            {
                ExecuteOpenInExplorer(obj);
            }
            else
            {
                ExecuteOpen(obj);
            }
        }

        private void ExecuteShellContextMenu(object obj)
        {
            try
            {
                // We need screen coordinates for the menu
                // Since this is a ViewModel, we have to refer to the View or pass parameters.
                // However, the cleanest way is to use GetCursorPos or get it from the mouse event if passed.
                // But for now, let's assume we get the cursor position using Win32 API in helper if needed?
                // Actually ShellContextMenuHelper.ShowContextMenu takes a Point.

                // Let's get the current cursor position
                var point = GetCursorPosition();

                // We need a window handle. Application.Current.MainWindow might be null or hidden.
                // But we need a valid window handle to receive messages.
                // We can use the dummy MainWindow we created.

                System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                {
                    var mainWindow = System.Windows.Application.Current.MainWindow as MainWindow;
                    if (mainWindow != null) mainWindow.SetStaysOpen(true);

                    try
                    {
                        NativeContextMenuHelper.ShowContextMenu(Path, (int)point.X, (int)point.Y, mainWindow);
                    }
                    finally
                    {
                        // Regain focus to the WPF window so it doesn't think it lost activation
                        if (mainWindow != null) mainWindow.Activate();

                        // Delay resetting StaysOpen to allow any pending focus/activation messages to settle
                        // while StaysOpen is still true.
                        System.Windows.Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                        {
                            if (mainWindow != null) mainWindow.SetStaysOpen(false);
                        }), System.Windows.Threading.DispatcherPriority.ApplicationIdle);
                    }
                }), System.Windows.Threading.DispatcherPriority.Input);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing context menu: {ex.Message}");
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        private System.Windows.Point GetCursorPosition()
        {
            POINT lpPoint;
            GetCursorPos(out lpPoint);
            return new System.Windows.Point(lpPoint.X, lpPoint.Y);
        }

        private void ExecuteOpen(object obj)
        {
            if (IsFolder) return; // Folders expand, files open
            try
            {
                Process.Start(new ProcessStartInfo(Path) { UseShellExecute = true });
                CloseMainWindow();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error opening file: {ex.Message}");
            }
        }

        private void ExecuteOpenInExplorer(object obj)
        {
            if (!IsFolder)
            {
                ExecuteOpen(obj);
                return;
            }

            try
            {
                string target = !string.IsNullOrEmpty(TargetFolderPath) ? TargetFolderPath : Path;
                Process.Start(new ProcessStartInfo(target) { UseShellExecute = true });
                CloseMainWindow();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error opening folder: {ex.Message}");
            }
        }

        private void CloseMainWindow()
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var mainWindow = System.Windows.Application.Current.MainWindow;
                if (mainWindow != null) mainWindow.Hide();
            });
        }

        private void ExecuteProperties(object obj)
        {
            try
            {
                ShellHelper.ShowFileProperties(Path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing properties: {ex.Message}");
            }
        }

        private void ExecuteCopyPath(object obj)
        {
            try
            {
                System.Windows.Clipboard.SetText(Path);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error copying path: {ex.Message}");
            }
        }

        public void LoadChildren()
        {
            if (!IsFolder) return;

            try
            {
                Children.Clear();
                string path = !string.IsNullOrEmpty(TargetFolderPath) ? TargetFolderPath : Path;
                var dirInfo = new DirectoryInfo(path);

                foreach (var directory in dirInfo.GetDirectories())
                {
                    Children.Add(new FileSystemItem(directory.FullName));
                }

                foreach (var file in dirInfo.GetFiles())
                {
                    Children.Add(new FileSystemItem(file.FullName));
                }
            }
            catch (Exception)
            {
                // Handle access exceptions or others
                Children.Add(new FileSystemItem { Name = "Access Denied" });
            }
        }
    }
}
