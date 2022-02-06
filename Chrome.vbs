'Script to open Chrome
Dim Shell
Set Shell = WScript.CreateObject("WScript.Shell")
Shell.Run "chrome.exe"
WScript.Sleep 1000

'to open second Chrome tab, uncomment the commented lines
'Shell.AppActivate "Chrome"
'WScript.Sleep 1000

'search VBScript
Shell.SendKeys "VBScript"
Shell.SendKeys "{ENTER}"
