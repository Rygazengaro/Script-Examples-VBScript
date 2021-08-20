'This script could be used to have user present information to helpdesk team
On Error Resume Next 

Set objWMIService = GetObject("winmgmts://./root/CIMV2")
Set colItems      = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration",,48)
For Each objItem In colItems
    If IsArray( objItem.IPAddress ) Then
        If UBound( objItem.IPAddress ) = 0 Then
            strIP = "IP Address: " & objItem.IPAddress(0)
        Else
            strIP = "IP Address: " & Join( objItem.IPAddress, "," )
        End If
    End If
Next

set colItems2 = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48) 
For Each objItem2 in colItems2 
strUsername = objItem2.UserName 
strComputer = objItem2.Name
Next 

strResults = "Please call the helpdesk at extension XXXX and present the following information:" & vbCrLf
strResults = strResults & "Server Name: " & strComputer & vbCrLf
strResults = strResults & Left(strIp,27) & vbCrLf
strResults = strResults & "Currently Logged-on User: " & strUsername & vbCrLf

msgBox strResults 