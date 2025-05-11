Skip to content
decoge/WindowsShutdownScript
Created 11 years ago • Report abuse
Clone this repository at &lt;script src=&quot;https://gist.github.com/decoge/9597061.js&quot;&gt;&lt;/script&gt;
<script src="https://gist.github.com/decoge/9597061.js"></script>
Code
Revisions
1
Windows Shutdown Script
WindowsShutdownScript
Const wshOK = 0 
Const wshExclamation = 48 
 
strComputer = "." 

strCounter = 1 
Set objShell = CreateObject("Wscript.Shell") 

strTimeLeft = 1 
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
 
Set colItems = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = 'explorer.exe'") 
 

 
while strCounter < 3
     
    intReturn = objShell.Popup("Your computer will shutdown in " & strTimeLeft & " minutes. Click OK to cancel the shutdown.",30, "Shutdown Pending", wshOK + wshExclamation) 
 
    If intReturn = 1 Then 
    objShell.Popup "Shutdown has been cancelled!", 6,"SHUTDOWN CANCELLED", wshOK + wshExclamation 
    Wscript.quit 0 
    End If 

    strTimeLeft = strTimeLeft - .5 
    strCounter = strCounter + 1 
 
Wend 

objShell.Run "C:\WINDOWS\system32\shutdown.exe -f -s"
Comment
 
Leave a comment
 
Footer
© 2025 GitHub, Inc.
Footer navigation
Terms
Privacy
Security
Status
Docs
Contact
Manage cookies
Do not share my personal information
