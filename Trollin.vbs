'This script is to poke fun at coworkers when they slack off and close their youtube and go to our work system
Set objShell = CreateObject("Shell.Application")
Set objShellWindows = objShell.Windows
intAlwaysOn = 1

Do While intAlwaysOn <= 50000000000
binFirstFound = False
For x = 0 to objShellWindows.Count - 1
Set objIE = objShellWindows.Item(x)
If InStr(objIE.LocationURL, "https://www.youtube.com/") then
If binFirstFound = False then
WScript.Sleep 100
objIE.quit
msgBox("Get back to work, Slacker!")
CreateObject("WScript.Shell").Run "http://worksystem/admin"
End If
End If
Next

intAlwaysOn = intAlwaysOn + 1
WScript.Sleep 100

Loop