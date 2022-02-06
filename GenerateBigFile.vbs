'script to generate big file
Dim filePath, fso, file, i
filePath = "BigFile.html"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(filePath, True)
For i = 1 to 1000

'making every 10th line is html comment
If i Mod 10 = 0 Then
file.Write("<!--" & i & " - dummy text" & "-->" & vbCrLf)

'making other lines html paragraphs
Else
file.Write("<p>" & i & " - dummy text" & "</p>" & vbCrLf)
End If
Next
file.Close