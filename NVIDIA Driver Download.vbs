Option Explicit

Dim Title, Shell
Title = "NVIDIA Driver Link Generator"
Set Shell = CreateObject("WScript.Shell")


Dim Version, System, Has64Bit
Version = InputBox("Driver version", Title, "417.71")
System = MsgBox("Win10?", vbYesNo, Title)
Has64Bit = MsgBox("64-bit?", vbYesNo, Title)

If System = vbYes Then
  System = "win10"
Else
  System = "win8-win7-winvista"
End If
If Has64Bit = vbYes Then
  Has64Bit = "64bit"
Else
  Has64Bit = "32bit"
End If

Dim Target, Shortcut, Action
' http://cn.download.nvidia.com/Windows/417.71/417.71-desktop-win10-64bit-international-whql.exe
Target = "http://cn.download.nvidia.com/Windows/" & Version & "/" & Version & "-desktop-" & System & "-" & Has64Bit & "-international-whql.exe"

Action = MsgBox("Download directly?", vbYesNo, Title)
If Action = vbYes Then
  Shell.Run Target, 9
Else
  Set Shortcut = Shell.CreateShortcut("NVIDIA Driver Download.lnk")
  Shortcut.TargetPath = Target
  Shortcut.Save
End If