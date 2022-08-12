on error resume next
msgbox "Thank you for using iBug SBCopy"&vbcrlf&"This program can help you make million+ copy-and-pastes quickly",,"Welcome"
set sshell = createobject("wscript.shell")
dim a,b,pinv
dim i
a=inputbox("Enter repeat times","Repeat Times",100)
b=inputbox("Enter repeat interval","Repeat Interval",0.5)
pinv=inputbox("Enter protection interval","Protection Interval",0.01)
b=b-pinv
if msgbox("Are you sure to continue iBug SBCopy?"&vbcrlf&"Total: "&a&vbcrlf&"Frequency: "&(b+pinv),vbyesno,"Confirm")=vbyes then
inputbox "Now you have 1.5 second to prepare after clicking [OK]."&vbcrlf&"Test your clipboard in the box below before clicking [OK].",,"Prepare"
wsh.sleep 1500
for i=1 to a
sshell.sendkeys "^v"
wsh.sleep pinv*1000
sshell.sendkeys "{ENTER}"
wsh.sleep b*1000
next
end if
wsh.quit