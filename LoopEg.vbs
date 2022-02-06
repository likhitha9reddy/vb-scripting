Dim input, position, length, char, result
result = "phone number"
input = InputBox("Enter telephone number or email:")
position = 1
length = Len(input)
Do Until position > length
  char = Mid(input, position, 1)
    If char = "@" Then
      result = "email"
      Exit Do
    End If
  position = position + 1
Loop
MsgBox "Thank you for entering your " & result
