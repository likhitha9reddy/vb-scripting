'Script to display the message
MsgBox: short for "message box", makes a script provide information
MsgBox "Hello!"

'Make the script wait
Sleep: pauses the program for a specified time (measured in milliseconds (ms))
WScript.Sleep(1000)
MsgBox "VBScript!"

'Alt. way to make the script wait
WScript.Sleep 2000
MsgBox("Hi!")

'Script to join 2 strings
'Const: short for "constant"; declare the word after it as a constant
'& symbol (ampersand); takes the values to either side of itself and joins them into a single text string/ concatenation 
'= symbol: assignment operator
Const Greeting = "Hi"
MsgBox Greeting & "!"
WScript.Sleep(2000)
MsgBox Greeting & " again!"

Const WaitTime = 2000
MsgBox "Click OK to wait for " & (WaitTime / 1000) & " seconds."
WScript.Sleep(WaitTime)
MsgBox "Wait completed!"

'InputBox lets the script ask a question and receive an answer.
'Dim is short for dimension, and is used to create, or "declare", a variable
Dim msg
msg = InputBox("Enter a message:")
MsgBox "You said: " & msg

Dim msg
Dim shout
msg = InputBox("Enter message to shout")
MsgBox "Shout & Whisper: " & msg & "?"
shout = UCase (msg)
MsgBox shout
whisper = LCase (msg)
MsgBox whisper
