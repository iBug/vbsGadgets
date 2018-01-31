' Author: iBug - https://stackoverflow.com/user/5958455/ibug

Option Explicit

Const ProgramName = "iBug PhotoManager"
Const Version = "1.0"
Const VersionNumber = 2
Dim Title : Title = ProgramName & " v" & Version & "." & VersionNumber

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim I_ConfigFileName, I_ValidExtensions, I_NumberLength, I_TargetPattern, I_LastNumber
I_ConfigFileName = "photo.ini"
I_ValidExtensions = Array("jpg", "raw", "jpeg", "jpe")
I_NumberLength = 8
I_TargetPattern = "IMG_%"
I_LastNumber = -1

Dim ConfigFile
Dim C_In, C_Out, C_Count, F_In, F_Out

''''''''''''''''''''''''''''''
GetConfig I_ConfigFileName
SearchFolder F_In, F_Out
''''''''''''''''''''''''''''''

Sub CreateDefaultConfig(ByVal ConfigFileName)
  Dim ConfigFile
  Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 2, True)
  ConfigFile.WriteLine "In="
  ConfigFile.WriteLine "Out="
  ConfigFile.WriteLine "Extension=jpg,jpeg,raw"
  ConfigFile.WriteLine "Pattern=IMG_%"
  ConfigFile.Close
End Sub

Sub ReadConfig(ConfigFileName)
  Dim ConfigFile, Config
  Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 1, False)
  Do Until ConfigFile.AtEndOfStream
    Config = Split(ConfigFile.ReadLine(), "=", 2)
    If UBound(Config) <> 1 Then
      ' Do nothing
    ElseIf Mid(LTrim(Config(0)), 1, 1) <> ";" Then
      Select Case LCase(Trim(Config(0)))
        Case "in"
          C_In = StripSlashes(Trim(Config(1)))
        Case "out"
          C_Out = StripSlashes(Trim(Config(1)))
        Case "extension"
          I_ValidExtensions = Split(Config(1), ",")
          Dim i
          For i = 0 To UBound(I_ValidExtensions)
            I_ValidExtensions(i) = Trim(I_ValidExtensions(i))
          Next
        Case "pattern"
          If UBound(Split(Config(1), "%")) = 1 Then
            I_TargetPattern = Config(1)
          Else
            MsgBox "Invalid pattern!", vbExclamation, Title
          End If
        Case Else
          MsgBox "Unknown option """ & Config(0) & """", vbExclamation, Title
          ' And... Do nothing
      End Select
    End If
  Loop
  ConfigFile.Close
End Sub

Function ValidateConfig()
  ValidateConfig = True
  If Not Fso.FolderExists(C_In) Then
    MsgBox "The source folder """ & C_In & """ does not exist!", 16, Title
    ValidateConfig = False
    Exit Function
  End If
  If Not Fso.FolderExists(C_Out) Then
    MsgBox "The target folder """ & C_Out & """ does not exist!", 16, Title
    ValidateConfig = False
    Exit Function
  End If
End Function

Sub GetExistingConfig(ByVal ConfigFileName, ByVal DeleteAfter)
  ' Read and apply the config
  ReadConfig ConfigFileName
  If DeleteAfter Then
    Fso.DeleteFile ConfigFileName
  End If
  
  ' Validate config
  If Not ValidateConfig Then
    WScript.Quit 1
  End If
  Set F_In = Fso.GetFolder(C_In)
  Set F_Out = Fso.GetFolder(C_Out)
End Sub

Sub GetConfig(ByVal ConfigFileName)
  If Not Fso.FileExists(ConfigFileName) Then
    CreateDefaultConfig ConfigFileName
    Shell.Run ConfigFileName,, True
  End If
  GetExistingConfig ConfigFileName, False
  C_Count = 0
End Sub

Sub PatchFile(ByRef File)
  If LCase(Fso.GetExtensionName(File.Path)) = "jpeg" Then
    File.Move GetFilenameWithoutExtension(File.Path) & ".jpg"
  End If
End Sub

Sub MoveFile(ByRef File, ByRef Target)
  PatchFile File
  File.Move Target.Path & "\" & FindNextAvailableName(F_Out) & "." & Fso.GetExtensionName(File.Path)
End Sub

Sub SearchFolder(ByRef Folder, ByRef Target)
  Dim Item
  For Each Item In Folder.Files
    If InArray(I_ValidExtensions, LCase(Fso.GetExtensionName(Item.Path))) Then
      MoveFile Item, Target
    End If
  Next
  For Each Item In Folder.Subfolders
    SearchFolder Item, Target
  Next
End Sub

Function FileIDExists(ByVal Number, ByVal Extension)
  FileIDExists = Fso.FileExists(C_Out & "\" & FormatFileName(Number) & Extension)
End Function

Function FormatNumber(ByVal Pattern, ByVal Number, ByVal NLength)
  Dim Filler, SNum
  Filler = "0000000000000000"
  SNum = CStr(Number)
  SNum = Mid(Filler, 1, NLength - Len(SNum)) & SNum
  FormatNumber = Replace(Pattern, "%", SNum)
End Function

Function UnformatNumber(ByVal Pattern, ByVal Format)
  If Len(Format) < Len(Pattern) Then
    UnformatNumber = -1
    Exit Function
  End If
  Dim LPattern, RPattern, S
  S = Split(Pattern, "%")
  LPattern = S(0)
  RPattern = S(UBound(S))
  If Left(Format, Len(LPattern)) = LPattern _
  And Right(Format, Len(RPattern)) = RPattern Then
    UnformatNumber = CLng(Mid(Format, 1+Len(LPattern), Len(Format)-Len(LPattern)-Len(RPattern)))
  Else
    UnformatNumber = -1
  End If
End Function

Function GetNextAvailableNumber(ByRef Folder)
  If I_LastNumber < 0 Then
    Dim Num, Item
    I_LastNumber = 0
    Num = 0
    For Each Item In Folder.Files
      Num = UnformatNumber(I_TargetPattern, Fso.GetBaseName(Item.Path))
      If InArray(I_ValidExtensions, LCase(Fso.GetExtensionName(Item.Path))) And _
        Num > I_LastNumber Then I_LastNumber = Num
    Next
  End If
  I_LastNumber = I_LastNumber + 1
  GetNextAvailableNumber = I_LastNumber
End Function

Function FindNextAvailableName(ByRef Folder)
  FindNextAvailableName = FormatNumber(I_TargetPattern, GetNextAvailableNumber(Folder), I_NumberLength)
End Function

Function StripSlashes(ByVal Str)
  Dim cutPoint
  cutPoint = Len(Str)
  Do While cutPoint > 0
    If Not Mid(Str, cutPoint, 1) = "/" Then Exit Do
    cutPoint = cutPoint - 1
  Loop
  StripSlashes = Left(Str, cutPoint)
End Function

Function GetFilenameWithoutExtension(ByVal FileName)
  Dim Result, i, j
  Result = FileName
  i = InStrRev(FileName, ".")
  j = InStrRev(FileName, "\")
  If i > 0 And i > j Then
    Result = Mid(FileName, 1, i - 1)
  End If
  GetFilenameWithoutExtension = Result
End Function

Function InArray(ByRef Arr, ByRef Match)
  Dim Item
  For Each Item In Arr
    If Item = Match Then
      InArray = True
      Exit Function
    End If
  Next
  InArray = False
End Function

