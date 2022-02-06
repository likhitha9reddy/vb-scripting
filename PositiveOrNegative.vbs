'Script to check for positive or negative number
Dim num, result
num = InputBox("Enter a number:")
If num > 0 Then
  result = num & " is a positive number."
Else
  If num = 0 Then
    result = num & " is zero."
  Else
    result = num & " is a negative number."
  End if
End if
MsgBox result
