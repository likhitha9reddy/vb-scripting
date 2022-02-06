'The "Mid" function gets you text from midway through a string. It takes arguments as follows:
'Mid(string, start [, length])
'The "start" argument says which character position to start from, e.g. "L" is at position one, "e" is at two, etc.
'The "length" argument is optional. Thats why its shown here with square brackets around it.

MsgBox Mid("Learn VBScript Fast", 2)
MsgBox Mid("Learn VBScript Fast", 1, 14)
MsgBox Mid("Learn VBScript Fast", 1, 1)
MsgBox Mid("Learn VBScript Fast", 7, 2)

Dim input
input = InputBox("Enter some text")
MsgBox Len(input)
MsgBox Mid(input, Len(input))
MsgBox Right(input, 3)
