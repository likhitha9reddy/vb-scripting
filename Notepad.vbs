'Script to open Notepad
Dim Shell
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run "notepad.exe"
WScript.Sleep 1000

'to open second Notepad window, uncomment the commented lines
'Shell.AppActivate "Notepad"
'WScript.Sleep 1000

'to send keystrokes Testing, enter and date in currently active Notepad window
Shell.SendKeys "Testing"
Shell.SendKeys "{ENTER}"
Shell.SendKeys Now()