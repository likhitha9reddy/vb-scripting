'script to run FizzBuzz Game
Dim i
For i = 10 To 15

'when number is a multiple of 3 & 5
If i Mod 3 = 0 and i Mod 5 = 0 Then
MsgBox "fizz buzz"

'when number is a multiple of 3
ElseIf i Mod 3 = 0 Then
MsgBox "fizz"

'when number is a multiple of 5
ElseIf i Mod 5 = 0 Then
MsgBox "buzz"

Else
MsgBox i
End If
Next