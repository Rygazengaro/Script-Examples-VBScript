'This script redirects a survey laptop to the start of the survey once it is taken so the next person can take it.
Set objShell = CreateObject("Shell.Application")
Set objShellWindows = objShell.Windows
intAlwaysOn = 1

Do While intAlwaysOn <= 5000000
binFirstFound = False
For x = 0 to objShellWindows.Count - 1
Set objIE = objShellWindows.Item(x)
If InStr(objIE.LocationURL, "https://www.surveymonkey.com/survey-thanks") then
If binFirstFound = False then
WScript.Sleep 5000
objIE.quit
CreateObject("WScript.Shell").Run "https://www.surveymonkey.com/r/SURVEYSTART"
End If
End If
Next

intAlwaysOn = intAlwaysOn + 1
WScript.Sleep 5000

Loop