using System;
using System.Diagnostics;
using System.IO;

namespace KoyoLauncher
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Find where the launcher is currently running
                string appDir = AppDomain.CurrentDomain.BaseDirectory;
                string backendPath = Path.Combine(appDir, "koyomail-backend.exe");
                
                if (File.Exists(backendPath))
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = backendPath;
                    psi.WorkingDirectory = appDir;
                    psi.CreateNoWindow = true;
                    psi.WindowStyle = ProcessWindowStyle.Hidden;
                    psi.UseShellExecute = false; // Runs directly without launching cmd.exe
                    
                    Process.Start(psi);
                }
            }
            catch
            {
                // Fail silently
            }
        }
    }
}
