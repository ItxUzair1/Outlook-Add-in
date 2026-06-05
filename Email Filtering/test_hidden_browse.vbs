' This simulates how the installer runs things — hidden window (SW_HIDE = 0)
' If the Browse dialog pops to FOREGROUND from this, it will work in production.
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cscript //nologo """ & Replace(WScript.ScriptFullName, "test_hidden_browse.vbs", "test_vbs_browse.vbs") & """", 0, False
