Option Explicit

Const FFMPEG_EXE = "ffmpeg.exe"


Function CatMP3(ffmpeg_file, source_file, output_file, start_sec, length_sec)
  ' ■ WScript.Shell オブジェクトについて
  '   「Windows管理者のためのWindows Script Host入門：第5回　WshShellオブジェクトの詳細（1） (2/4) - ＠IT」
  '      http://www.atmarkit.co.jp/ait/articles/0407/08/news101_2.html
  Dim objShell
  Set objShell = WScript.CreateObject("WScript.Shell")
  
  '   Call objShell.Run(sFFmpegRun , 1, true)
End Function


' ======================================================================
'   Functions for System
' ======================================================================

Sub Include(ByVal strFile)
  Dim objStream
  Set objStream = CreateObject("Scripting.FileSystemObject").OpenTextFile(strFile, 1)
  ExecuteGlobal objStream.ReadAll() 
  objStream.Close 
  Set objStream = Nothing 
End Sub

Function GetModulePath()
  Dim PathLen, ExePath
  PathLen = Len(WScript.Scriptfullname) - Len(WScript.Scriptname) - 1
  GetModulePath = Left(Wscript.Scriptfullname, PathLen)
End Function

Function GetLibPath()
  GetLibPath = GetModulePath() & "\" & "lib"
End Function

Function GetUserFilePath()
  Dim Arguments
  Set Arguments = WScript.Arguments
  GetUserFilePath = WScript.Arguments.item(0)
End Function

Function GetFfmpegBinFile()
  Dim file, Fso
  Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
  file = GetModulePath() & "\" & FFMPEG_EXE
  If Not Fso.FileExists(file) Then
    Error("指定された場所に" & FFMPEG_EXE & "が存在しません")
  End If
  GetFfmpegBinFile = file
End Function

Function InitErrorCheck()
  Dim Arguments, Fso
  Set Arguments = WScript.Arguments
  Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
  If Not Arguments.count = 1 Then
    Error("ファイルが指定されていません")
  End If
  If Not Fso.FileExists( Arguments.item(0) ) Then
    Error("指定されたファイルは存在しません")
  End If
End Function

Function Error(Message)
  MsgBox(Message)
  WScript.Quit 1
End Function

Function RemoveSpace(str)
  Dim re
  Set re = New RegExp
  re.Pattern = "\s*"
  re.Global = True
  RemoveSpace = re.Replace(str, "")
End Function

Function TrimEx(str)
  TrimEx= LTrimEx(str)
  TrimEx= RTrimEx(str)
End Function

Function LTrimEx(str)
  Dim re
  Set re = New RegExp
  re.Pattern = "^\s*"
  re.Multiline = False
  LTrimEx = re.Replace(str, "")
End Function

Function RTrimEx(str)
  Dim re
  Set re = New RegExp
  re.Pattern = "\s*$"
  re.Multiline = False
  RTrimEx = re.Replace(str, "")
End Function

Function ConfirmSplit(UserFile)
  Dim message, i, elements, is1st

  message =                    "【フォルダ　　】"
  message = message & vbCrLf & "　" & UserFile.Dir
  message = message & vbCrLf & "【分割元ファイル】"
  message = message & vbCrLf & "　" & UserFile.Mp3File
  message = message & vbCrLf & "【分割先ファイル】"

  elements = UserFile.Mp3Elements

  For i = 1 To UBound(elements)
	If Not elements(i).IsExists() Then
	  message = message & vbCrLf & "　" & elements(i).TimeBegin & " - " & elements(i).TimeEnd & " / " & elements(i).FileName
	End If
  Next

  is1st = True
  For i = 1 To UBound(elements)
	If elements(i).IsExists() Then
	  If is1st Then
		message = message & vbCrLf & vbCrLf & "以下は存在するためスキップします"
		is1st = False
	  End If
	  message = message & vbCrLf & "－ " & elements(i).FileName
	End If
  Next

  message = message & vbCrLf & vbCrLf & "分割を開始してよろしいですか？"
  
  MsgBox(message)
  
End Function

' ======================================================================
'   Functions for Debug
' ======================================================================

Function GetSamplePath()
  ' テストディレクトリの取得
  '     このVBSファイルから ../sample の位置を想定
  GetSamplePath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(GetModulePath()) & "\sample"
End Function

Function debug(Message)
  MsgBox("Debug: " & Message)
End Function

' ======================================================================
'   Main_Exec
' ======================================================================
Function Main_Exec()
  Dim UserFile
  Set UserFile = New UserFile

  ' Call InitErrorCheck()
  ' UserFile.Path = GetUserFilePath()
  UserFile.Path = GetSamplePath() & "\entry.txt"
  UserFile.LoadSettings
  ConfirmSplit(UserFile)

End Function


' ======================================================================
'   Main
' ======================================================================
Include GetLibPath() & "\UserFile.class.vbs"
Include GetLibPath() & "\Mp3Element.class.vbs"
Call Main_Exec

  ' Ffmpeg = GetFfmpegBinFile()

