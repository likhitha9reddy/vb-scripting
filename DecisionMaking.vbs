'Scripts for decision making
MsgBox 1 = 1
MsgBox 1 <> 1
MsgBox 1 < 2
MsgBox 1 > 2
MsgBox 1 <= 2
MsgBox 1 >= 2

Const divisor = 3
Dim num, result
num = InputBox("Enter a multiple of " & divisor)
result = ((num Mod divisor) = 0)
MsgBox result
