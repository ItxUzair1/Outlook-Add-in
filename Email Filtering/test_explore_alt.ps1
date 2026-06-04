$wshell = New-Object -ComObject wscript.shell
$wshell.SendKeys('%')
$shell = New-Object -ComObject Shell.Application
$shell.Open("C:\")
