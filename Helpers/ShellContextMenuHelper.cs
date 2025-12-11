using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;

namespace TrayFolder.Helpers
{
    public static class NativeContextMenuHelper
    {
        #region COM Interfaces

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E4-0000-0000-C000-000000000046")]
        public interface IContextMenu
        {
            [PreserveSig]
            int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
            [PreserveSig]
            int InvokeCommand(ref CMINVOKECOMMANDINFO pici);
            [PreserveSig]
            int GetCommandString(UIntPtr idCmd, uint uType, IntPtr pReserved, StringBuilder pszName, uint cchMax);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214f4-0000-0000-c000-000000000046")]
        public interface IContextMenu2 : IContextMenu
        {
            [PreserveSig]
            new int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
            [PreserveSig]
            new int InvokeCommand(ref CMINVOKECOMMANDINFO pici);
            [PreserveSig]
            new int GetCommandString(UIntPtr idCmd, uint uType, IntPtr pReserved, StringBuilder pszName, uint cchMax);
            [PreserveSig]
            int HandleMenuMsg(uint uMsg, IntPtr wParam, IntPtr lParam);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("BCFCE0A0-EC17-11d0-8D10-00A0C90F2719")]
        public interface IContextMenu3 : IContextMenu2
        {
            [PreserveSig]
            new int QueryContextMenu(IntPtr hMenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);
            [PreserveSig]
            new int InvokeCommand(ref CMINVOKECOMMANDINFO pici);
            [PreserveSig]
            new int GetCommandString(UIntPtr idCmd, uint uType, IntPtr pReserved, StringBuilder pszName, uint cchMax);
            [PreserveSig]
            new int HandleMenuMsg(uint uMsg, IntPtr wParam, IntPtr lParam);
            [PreserveSig]
            int HandleMenuMsg2(uint uMsg, IntPtr wParam, IntPtr lParam, IntPtr plResult);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214E6-0000-0000-C000-000000000046")]
        public interface IShellFolder
        {
            [PreserveSig]
            int ParseDisplayName(IntPtr hwnd, IntPtr pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, out uint pchEaten, out IntPtr ppidl, ref uint pdwAttributes);
            [PreserveSig]
            int EnumObjects(IntPtr hwnd, int grfFlags, out IntPtr ppenumIDList);
            [PreserveSig]
            int BindToObject(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);
            [PreserveSig]
            int BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);
            [PreserveSig]
            int CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2);
            [PreserveSig]
            int CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv);
            [PreserveSig]
            int GetAttributesOf(uint cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, ref uint rgfInOut);
            [PreserveSig]
            int GetUIObjectOf(IntPtr hwndOwner, uint cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, [In] ref Guid riid, IntPtr rgfReserved, out IntPtr ppv);
            [PreserveSig]
            int GetDisplayNameOf(IntPtr pidl, uint uFlags, IntPtr pName);
            [PreserveSig]
            int SetNameOf(IntPtr hwnd, IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, uint uFlags, out IntPtr ppidlOut);
        }

        #endregion

        #region Structs and Constants

        [StructLayout(LayoutKind.Sequential)]
        public struct CMINVOKECOMMANDINFO
        {
            public int cbSize;
            public int fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            public IntPtr lpParameters;
            public IntPtr lpDirectory;
            public int nShow;
            public int dwHotKey;
            public IntPtr hIcon;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern int SHGetDesktopFolder(out IShellFolder ppshf);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern bool DestroyMenu(IntPtr hMenu);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        public static extern void SHFree(IntPtr pv);

        [DllImport("shell32.dll")]
        public static extern IntPtr ILFindLastID(IntPtr pidl);

        [DllImport("shell32.dll")]
        public static extern bool ILRemoveLastID(IntPtr pidl);

        [DllImport("shell32.dll")]
        public static extern void ILFree(IntPtr pidl);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

        public const uint CMF_NORMAL = 0x00000000;
        public const uint CMF_EXPLORE = 0x00000004;

        public const uint TPM_ReturnCmd = 0x0100;
        public const uint TPM_LeftButton = 0x0000;
        public const uint TPM_RightButton = 0x0002;

        public static Guid IID_IContextMenu = new Guid("000214E4-0000-0000-C000-000000000046");
        public static Guid IID_IShellFolder = new Guid("000214E6-0000-0000-C000-000000000046");

        #endregion

        public static void ShowContextMenu(string path, int x, int y, Window ownerWindow)
        {
            if (string.IsNullOrEmpty(path)) return;

            IntPtr pidlFull = IntPtr.Zero;
            IntPtr pidlChild = IntPtr.Zero;
            IShellFolder desktopFolder = null;
            IShellFolder parentFolder = null;
            IntPtr hMenu = IntPtr.Zero;
            IntPtr iContextMenuPtr = IntPtr.Zero;

            try
            {
                // 1. Get Desktop Folder
                if (SHGetDesktopFolder(out desktopFolder) != 0) return;

                // 2. Parse Path to PIDL
                uint attributes = 0;
                if (SHParseDisplayName(path, IntPtr.Zero, out pidlFull, 0, out attributes) != 0) return;

                // 3. Split PIDL into Parent and Child
                // We need the parent IShellFolder and the child PIDL (relative to parent)

                // Get the last ID (child) and remove it from full to get parent PIDL?
                // Actually SHParseDisplayName returns an absolute PIDL (relative to Desktop).
                // Easier way: Bind to parent folder. But we need to separate the path or the PIDL.

                // Alternative: Use SHBindToParent (Standard API)
                // Let's implement SHBindToParent P/Invoke as it's cleaner.
                // But since I didn't define it, I'll do it manually:

                // Manually getting parent IShellFolder is complex with raw PIDLs without SHBindToParent.
                // Let's use the file path string manipulation which is safer for this simple case? 
                // No, PIDLs are better for special folders. But let's assume standard file system paths here.

                // Let's use SHBindToParent. I'll add the P/Invoke dynamically or just add it below.

                IntPtr pidlParent;
                if (SHBindToParent(pidlFull, ref IID_IShellFolder, out parentFolder, out pidlChild) != 0)
                {
                    // Fallback to Desktop if it's a root or something
                    return;
                }

                // 4. Get IContextMenu
                IntPtr[] apidl = new IntPtr[] { pidlChild };
                if (parentFolder.GetUIObjectOf(IntPtr.Zero, 1, apidl, ref IID_IContextMenu, IntPtr.Zero, out iContextMenuPtr) != 0) return;

                var contextMenu = (IContextMenu)Marshal.GetObjectForIUnknown(iContextMenuPtr);

                // 5. Create Menu
                hMenu = CreatePopupMenu();

                // 6. Query Context Menu
                contextMenu.QueryContextMenu(hMenu, 0, 1, 0x7FFF, CMF_NORMAL);

                // 7. Track Menu
                IntPtr hwnd = (new WindowInteropHelper(ownerWindow)).Handle;
                int command = TrackPopupMenuEx(hMenu, TPM_ReturnCmd | TPM_LeftButton, x, y, hwnd, IntPtr.Zero);

                // 8. Invoke Command
                if (command > 0)
                {
                    CMINVOKECOMMANDINFO invoke = new CMINVOKECOMMANDINFO();
                    invoke.cbSize = Marshal.SizeOf(invoke);
                    invoke.hwnd = hwnd;
                    invoke.lpVerb = (IntPtr)(command - 1);
                    invoke.nShow = 1; // SW_SHOWNORMAL

                    contextMenu.InvokeCommand(ref invoke);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex);
            }
            finally
            {
                if (hMenu != IntPtr.Zero) DestroyMenu(hMenu);
                if (iContextMenuPtr != IntPtr.Zero) Marshal.Release(iContextMenuPtr);
                if (parentFolder != null) Marshal.ReleaseComObject(parentFolder);
                if (desktopFolder != null) Marshal.ReleaseComObject(desktopFolder);
                if (pidlFull != IntPtr.Zero) ILFree(pidlFull);
                // pidlChild is a pointer into pidlFull (or separate depending on SHBindToParent impl), 
                // but usually SHBindToParent returns a pointer to the last ID in the PIDL list, not a new allocation, so valid as long as pidlFull is valid?
                // Actually SHBindToParent DOES NOT allocate a new PIDL for child, it points to the last ID in the absolute PIDL.
                // So we only free pidlFull.
            }
        }

        [DllImport("shell32.dll", ExactSpelling = true, PreserveSig = true)]
        public static extern int SHBindToParent(IntPtr pidl, [In] ref Guid riid, out IShellFolder ppv, out IntPtr ppidlLast);
    }
}
