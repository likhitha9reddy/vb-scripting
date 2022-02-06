Dim num : num = 0
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
Dim f : Set f = fso.OpenTextFile("BigFile.txt")
Dim line
Do Until f.AtEndOfStream
  line = f.ReadLine
  If inStr(line, "00 ") Then
   num = num + 1
  End If
Loop
f.Close
MsgBox num & " lines contain matching text"
