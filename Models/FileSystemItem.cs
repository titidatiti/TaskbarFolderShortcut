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
        }

        public FileSystemItem(string path)
        {
            Path = path;
            Name = System.IO.Path.GetFileName(path);
            if (string.IsNullOrEmpty(Name)) Name = path; // Handle root drives
            IsFolder = Directory.Exists(path);
            
            try 
            {
                SysIcon = IconHelper.GetIcon(path, IsFolder);
            }
            catch { }

            Children = new ObservableCollection<FileSystemItem>();
            OpenCommand = new RelayCommand(ExecuteOpen);

            if (IsFolder)
            {
                // Add a dummy item to enable expansion if it's a folder
                Children.Add(new FileSystemItem { Name = "Loading..." });
            }
        }

        private void ExecuteOpen(object obj)
        {
            if (IsFolder) return; // Folders expand, files open
            try
            {
                Process.Start(new ProcessStartInfo(Path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                // Show error?
                Debug.WriteLine($"Error opening file: {ex.Message}");
            }
        }

        public void LoadChildren()
        {
            if (!IsFolder) return;

            try
            {
                Children.Clear();
                var dirInfo = new DirectoryInfo(Path);
                
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
