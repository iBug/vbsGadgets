' Author: iBug

Option Explicit

Dim Shell, FS
Set Shell = CreateObject("WScript.Shell")
Set FS = CreateObject("Scripting.FileSystemObject")

Sub Conv(FileName)
  Dim DOC
  Set DOC = Word.Documents.Open(FileName)
  'WScript.Sleep 2000
  ' MS Word is a little strange and requires a workaround
  DOC.SaveAs2 FileName & "x", 12
  DOC.Close
End Sub

Sub ConvAll(Dir)
  Dim Item
  For Each Item In Dir.Files
    If LCase(FS.GetExtensionName(Item.Path)) = "doc" Then
      Conv Item.Path
    End If
  Next
  For Each Item In Dir.SubFolders
    ConvAll Item
  Next
End Sub

Dim Word
Set Word = CreateObject("Word.Application")
Word.Visible = True
ConvAll FS.GetFolder(".")
Word.Quit
