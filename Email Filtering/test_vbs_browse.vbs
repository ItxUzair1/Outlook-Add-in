Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

' Create a focus-helper script that runs ASYNCHRONOUSLY.
' It waits 300ms for the BrowseForFolder dialog to appear,
' then sends Alt key and uses AppActivate to force the dialog to foreground.
Dim helperPath
helperPath = fso.GetSpecialFolder(2) & "\koyo_focus_" & Replace(Timer, ",", "") & ".vbs"

Set f = fso.CreateTextFile(helperPath, True)
f.WriteLine "WScript.Sleep 300"
f.WriteLine "Set ws = CreateObject(""WScript.Shell"")"
f.WriteLine "ws.SendKeys ""%"""
f.WriteLine "WScript.Sleep 100"
f.WriteLine "ws.AppActivate ""Browse For Folder"""
f.Close

' Launch the focus helper in the background (non-blocking)
WshShell.Run "wscript """ & helperPath & """", 0, False

' Break foreground lock and show dialog
WshShell.SendKeys "%"

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.BrowseForFolder(0, "Select Destination Folder", &H50, 0)

' Clean up helper file
On Error Resume Next
fso.DeleteFile helperPath, True
On Error Goto 0

If Not objFolder Is Nothing Then
    WScript.Echo objFolder.Self.Path
End If
