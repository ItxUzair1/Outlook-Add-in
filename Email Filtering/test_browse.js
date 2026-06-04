const {exec} = require('child_process');
const psScript = `
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class FocusHelper {
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr ProcessId);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    public static void ForceForeground(IntPtr hWnd) {
        IntPtr fg = GetForegroundWindow();
        if (fg == hWnd) return;
        uint fgThread = GetWindowThreadProcessId(fg, IntPtr.Zero);
        uint curThread = GetCurrentThreadId();
        if (fgThread != curThread) {
            AttachThreadInput(curThread, fgThread, true);
            SetForegroundWindow(hWnd);
            AttachThreadInput(curThread, fgThread, false);
        } else {
            SetForegroundWindow(hWnd);
        }
    }
}
"@

Add-Type -AssemblyName System.Windows.Forms
$f = New-Object System.Windows.Forms.Form
$f.TopMost = $true
$f.Show()
[FocusHelper]::ForceForeground($f.Handle)
$d = New-Object System.Windows.Forms.FolderBrowserDialog
$d.Description = "Select Destination Folder"
$d.ShowNewFolderButton = $true
$result = $d.ShowDialog($f)
$f.Dispose()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath }
`;
const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
exec(`powershell -Sta -NoProfile -EncodedCommand ${encoded}`, (err, out, stderr) => {
    console.log("ERR:", err);
    console.log("OUT:", out);
    console.log("STDERR:", stderr);
});
