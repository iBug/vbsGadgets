' Author: iBug

Option Explicit

Dim Shell, FS
Set Shell = CreateObject("WScript.Shell")
Set FS = CreateObject("Scripting.FileSystemObject")

Sub Conv(FileName)
  Dim XLS
  Set XLS = Excel.Workbooks.Open(FileName)
  'WScript.Sleep 2000
  XLS.SaveCopyAs FileName & "x"
  XLS.Close
End Sub

Sub ConvAll(Dir)
  Dim Item
  For Each Item In Dir.Files
    If LCase(FS.GetExtensionName(Item.Path)) = "xls" Then
      Conv Item.Path
    End If
  Next
  For Each Item In Dir.SubFolders
    ConvAll Item
  Next
End Sub

Dim Excel
Set Excel = CreateObject("Excel.Application")
Excel.Visible = True
ConvAll FS.GetFolder(".")
Excel.Quit
