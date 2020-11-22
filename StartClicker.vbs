set wshshell = wscript.CreateObject("wscript.shell")
wscript.sleep 2000
do
WshShell.SendKeys "KEY"
WScript.Sleep 50
loop