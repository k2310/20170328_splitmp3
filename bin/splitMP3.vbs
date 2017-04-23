Option Explicit

Const FFMPEG_EXE = "ffmpeg.exe"


Function CatMP3(ffmpeg_file, source_file, output_file, start_sec, length_sec)
  ' �� WScript.Shell �I�u�W�F�N�g�ɂ���
  '   �uWindows�Ǘ��҂̂��߂�Windows Script Host����F��5��@WshShell�I�u�W�F�N�g�̏ڍׁi1�j (2/4) - ��IT�v
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
    Error("�w�肳�ꂽ�ꏊ��" & FFMPEG_EXE & "�����݂��܂���")
  End If
  GetFfmpegBinFile = file
End Function

Function InitErrorCheck()
  Dim Arguments, Fso
  Set Arguments = WScript.Arguments
  Set Fso = WScript.CreateObject("Scripting.FileSystemObject")
  If Not Arguments.count = 1 Then
    Error("�t�@�C�����w�肳��Ă��܂���")
  End If
  If Not Fso.FileExists( Arguments.item(0) ) Then
    Error("�w�肳�ꂽ�t�@�C���͑��݂��܂���")
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

  message =                    "�y�t�H���_�@�@�z"
  message = message & vbCrLf & "�@" & UserFile.Dir
  message = message & vbCrLf & "�y�������t�@�C���z"
  message = message & vbCrLf & "�@" & UserFile.Mp3File
  message = message & vbCrLf & "�y������t�@�C���z"

  elements = UserFile.Mp3Elements

  For i = 1 To UBound(elements)
	If Not elements(i).IsExists() Then
	  message = message & vbCrLf & "�@" & elements(i).TimeBegin & " - " & elements(i).TimeEnd & " / " & elements(i).FileName
	End If
  Next

  is1st = True
  For i = 1 To UBound(elements)
	If elements(i).IsExists() Then
	  If is1st Then
		message = message & vbCrLf & vbCrLf & "�ȉ��͑��݂��邽�߃X�L�b�v���܂�"
		is1st = False
	  End If
	  message = message & vbCrLf & "�| " & elements(i).FileName
	End If
  Next

  message = message & vbCrLf & vbCrLf & "�������J�n���Ă�낵���ł����H"
  
  MsgBox(message)
  
End Function

' ======================================================================
'   Functions for Debug
' ======================================================================

Function GetSamplePath()
  ' �e�X�g�f�B���N�g���̎擾
  '     ����VBS�t�@�C������ ../sample �̈ʒu��z��
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

