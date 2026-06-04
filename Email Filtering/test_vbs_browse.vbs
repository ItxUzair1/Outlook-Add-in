Set objShell = WScript.CreateObject("Shell.Application")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.SendKeys "%"
Set objFolder = objShell.BrowseForFolder(0, "Select Destination Folder", 0, 0)
If Not objFolder Is Nothing Then
    WScript.Echo objFolder.Self.Path
End If
