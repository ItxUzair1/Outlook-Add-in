Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
$f = New-Object System.Windows.Forms.Form
$f.TopMost = $true
$f.ShowInTaskbar = $true
$f.Text = "Koyomail Browse"
$f.Show()
[Microsoft.VisualBasic.Interaction]::AppActivate($f.Text)

$d = New-Object System.Windows.Forms.FolderBrowserDialog
$d.Description = "Select Destination Folder"
$d.ShowNewFolderButton = $true
$d.ShowDialog($f)

$f.Dispose()
