'Script to write to a file
Dim filePath, fso, file

'creates a file called MoreText.txt in the same folder as of script
filePath = "MoreText.txt"

'CreateObject creates FileSystemObject (fso)
Set fso = CreateObject("Scripting.FileSystemObject")

'fso has 2 parameters - path of file and boolean
'boolean True -> created file should be overwritten
'boolean Flase -> craeted file should not be overwitten
Set file = fso.CreateTextFile(filePath, True)
file.Write("Success!")
file.Close

-----------------------------------------------------------------------------------------------------------

'Script to write to a file using with... end with stmts
Dim filePath, fso, file
filePath = "MoreText.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(filePath, True)

'with allows to reference an object once and use it repeatedly
With file
.Write("Success!")
.Close
End With