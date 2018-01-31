' Author: iBug

Option Explicit

Dim Shell, FS
Set Shell = CreateObject("WScript.Shell")
Set FS = CreateObject("Scripting.FileSystemObject")

Sub Conv(FileName)
  Dim PPT
  Set PPT = PowerPoint.Presentations.Open(FileName)
  'WScript.Sleep 2000
  PPT.SaveCopyAs FileName & "x"
  PPT.Close
End Sub

Sub ConvAll(Dir)
  Dim Item
  For Each Item In Dir.Files
    If LCase(FS.GetExtensionName(Item.Path)) = "ppt" Then
      Conv Item.Path
    End If
  Next
  For Each Item In Dir.SubFolders
    ConvAll Item
  Next
End Sub

Dim PowerPoint
Set PowerPoint = CreateObject("PowerPoint.Application")
PowerPoint.Visible = True
ConvAll FS.GetFolder(".")
PowerPoint.Quit
