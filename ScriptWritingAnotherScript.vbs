'script to create another script
Dim filePath, fso, file
filePath = "Greeting.vbs"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(filePath, True)
With file
.Write("MsgBox ""Hello!""")
.Close
End With

'script to run the other script
Dim shell
set shell = WScript.CreateObject("WScript.Shell")
shell.run filePath