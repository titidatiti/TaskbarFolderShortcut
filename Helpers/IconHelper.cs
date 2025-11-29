using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace TrayFolder.Helpers
{
    public static class IconHelper
    {
        public static ImageSource GetIcon(string path, bool isFolder, bool isExpanded = false)
        {
            // Simple approach: Use SHGetFileInfo for robust icon retrieval
            // But for simplicity in this iteration, let's try System.Drawing.Icon first where possible,
            // or a basic P/Invoke if needed.
            
            // Actually, for folders, ExtractAssociatedIcon doesn't work.
            // So we need SHGetFileInfo.
            
            return GetIconFromShell(path, isFolder, isExpanded);
        }

        private static ImageSource GetIconFromShell(string path, bool isFolder, bool isExpanded)
        {
            var shinfo = new SHFILEINFO();
            uint flags = SHGFI_ICON | SHGFI_SMALLICON;
            
            // Use basic attributes if path doesn't exist or just to be safe
            uint attributes = FILE_ATTRIBUTE_NORMAL;
            if (isFolder)
            {
                attributes |= FILE_ATTRIBUTE_DIRECTORY;
                flags |= SHGFI_USEFILEATTRIBUTES; // Use attributes instead of accessing file system (faster)
            }
            
            if (isFolder && isExpanded)
            {
                flags |= SHGFI_OPENICON;
            }

            SHGetFileInfo(path, attributes, ref shinfo, (uint)Marshal.SizeOf(shinfo), flags);

            if (shinfo.hIcon == IntPtr.Zero) return null;

            var icon = System.Drawing.Icon.FromHandle(shinfo.hIcon);
            var imageSource = Imaging.CreateBitmapSourceFromHIcon(
                icon.Handle,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            DestroyIcon(shinfo.hIcon);
            return imageSource;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            public uint dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        };

        private const uint SHGFI_ICON = 0x100;
        private const uint SHGFI_SMALLICON = 0x1;
        private const uint SHGFI_OPENICON = 0x2;
        private const uint SHGFI_USEFILEATTRIBUTES = 0x10;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x80;
        private const uint FILE_ATTRIBUTE_DIRECTORY = 0x10;

        [DllImport("shell32.dll")]
        private static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DestroyIcon(IntPtr hIcon);
    }
}
