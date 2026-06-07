using System;
using System.Runtime.InteropServices;

namespace KoyoBrowse
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                // Instantiate the native FileOpenDialog COM class
                Type dialogType = Type.GetTypeFromCLSID(new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7"));
                var dialog = (IFileOpenDialog)Activator.CreateInstance(dialogType);

                if (dialog != null)
                {
                    // Set options: FOS_PICKFOLDERS (0x20) | FOS_FORCEFILESYSTEM (0x40)
                    dialog.SetOptions(0x20 | 0x40);
                    
                    string title = args.Length > 0 ? args[0] : "Select Destination Folder";
                    dialog.SetTitle(title);

                    // Get current active window (Outlook) to set it as the owner, which forces the dialog to foreground
                    IntPtr owner = GetForegroundWindow();
                    int hr = dialog.Show(owner);
                    
                    if (hr == 0) // S_OK
                    {
                        IShellItem result;
                        dialog.GetResult(out result);
                        if (result != null)
                        {
                            string path;
                            // SIGDN_FILESYSPATH = 0x80058000
                            result.GetDisplayName(0x80058000, out path);
                            Console.Write(path);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("Error: " + ex.Message);
                Environment.Exit(1);
            }
        }
    }

    [ComImport]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")] // IFileOpenDialog Guid
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, ref uint rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(uint fos);
        void GetOptions(out uint fos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close([MarshalAs(UnmanagedType.Error)] int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")] // IShellItem Guid
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(uint sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }
}
