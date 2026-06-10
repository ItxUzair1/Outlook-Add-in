using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace KoyoFile
{
    class Program
    {
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        public class WindowWrapper : IWin32Window
        {
            public IntPtr Handle { get; private set; }
            public WindowWrapper(IntPtr handle) { Handle = handle; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            try 
            {
                Application.EnableVisualStyles();
                
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.Filter = "Collection Files (*.mmcollection)|*.mmcollection|All Files (*.*)|*.*";
                    ofd.Title = "Select Collection File";
                    
                    IntPtr hwnd = GetForegroundWindow();
                    WindowWrapper wrapper = new WindowWrapper(hwnd);
                    
                    DialogResult result = ofd.ShowDialog(wrapper);
                    
                    if (result == DialogResult.OK)
                    {
                        Console.Write(ofd.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }
    }
}
