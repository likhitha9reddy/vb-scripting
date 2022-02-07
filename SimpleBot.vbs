'script to create a simple bot
script and other dependecy files are placed in C:\xampp\htdocs\php\vb_scripts

Dim l,rand
l=1
Do While l<4

'Open Internet Explorer
Set obj = CreateObject("InternetExplorer.Application")
obj.Naviagte("http://localhost/php/vb_scripts/Form.html")
obj.Visible = True

WScript.Sleep 1000

Randomize
rand = Int((10000)*Rnd+10000)

obj.Document.all.item("email").Value =rand&"@something.com"
obj.Document.all.item("pwd").Value ="$tr0nGP@$$w0rd"

WScript.Sleep 1000

obj.Document.all.Item("submit").click
WScript.Sleep 500

obj.Quit
l = l + 1