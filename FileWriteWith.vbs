Dim filePath, fso, file
filePath = "SomeText.txt"
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.CreateTextFile(filePath, True)
With file
  .Write("File named: " & filePath & vbCrLf)
  .Write("Created at: " & Now())
  .Close
End With
