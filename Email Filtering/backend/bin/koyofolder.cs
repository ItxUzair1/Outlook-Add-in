using System;
using System.Windows.Forms;
using System.IO;

namespace KoyoFolder
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            
            string title = "Select Destination Folder";
            string initialDir = "";
            
            if (args.Length > 0) title = args[0];
            if (args.Length > 1) initialDir = args[1];

            try
            {
                using (OpenFileDialog ofd = new OpenFileDialog())
                {
                    ofd.ValidateNames = false;
                    ofd.CheckFileExists = false;
                    ofd.CheckPathExists = true;
                    // Always default to a dummy file name so user can just hit Open
                    ofd.FileName = "Folder Selection.";
                    ofd.Title = title;
                    
                    if (!string.IsNullOrEmpty(initialDir) && Directory.Exists(initialDir))
                    {
                        ofd.InitialDirectory = initialDir;
                    }
                    
                    Form form = new Form();
                    form.TopMost = true;
                    form.ShowInTaskbar = false;
                    form.WindowState = FormWindowState.Minimized;
                    form.Show();
                    form.Focus();
                    
                    DialogResult result = ofd.ShowDialog(form);
                    
                    if (result == DialogResult.OK)
                    {
                        string path = ofd.FileName;
                        if (path.EndsWith("Folder Selection."))
                        {
                            path = path.Substring(0, path.Length - "Folder Selection.".Length);
                            if (path.EndsWith("\\") && path.Length > 3)
                                path = path.Substring(0, path.Length - 1);
                        }
                        else 
                        {
                            path = Path.GetDirectoryName(path);
                        }
                        Console.WriteLine(path);
                    }
                    
                    form.Close();
                }
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex.Message);
            }
        }
    }
}
