using System;
using System.Windows.Forms;

namespace KoyoFile
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Collection Files (*.mmcollection)|*.mmcollection|All Files (*.*)|*.*";
                ofd.Title = "Select Collection File";
                
                Form form = new Form();
                form.TopMost = true;
                form.ShowInTaskbar = false;
                form.WindowState = FormWindowState.Minimized;
                form.Show();
                form.Focus();
                
                DialogResult result = ofd.ShowDialog(form);
                
                if (result == DialogResult.OK)
                {
                    Console.WriteLine(ofd.FileName);
                }
                
                form.Close();
            }
        }
    }
}
