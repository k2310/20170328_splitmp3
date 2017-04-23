' ======================================================================
'   Class UserFile
' ======================================================================
Class UserFile
  Public Sub Class_Initialize()
  End Sub

  ' ----- Path
  Public Property Let Path(value)
    m_path = value
  End Property
  Public Property Get Path
    Path = m_path
  End Property
  ' ----- Dir
  Public Property Get Dir
    Dir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Me.Path)
  End Property

  ' ----- Mp3File
  Public Property Let Mp3File(value)
    m_mp3_file = value
  End Property
  Public Property Get Mp3File
    Mp3File = m_mp3_file
  End Property
  Public Property Get Mp3FileWithPath
    Mp3FileWithPath =  Me.Dir & "\" & Me.Mp3File
  End Property

  ' ----- Mp3Elements
  Public Property Get Mp3Elements
    Mp3Elements = m_mp3_elements
  End Property

  Public Sub LoadSettings
    ' Load from Textfile
    Dim File
    Set File = CreateObject("Scripting.FileSystemObject").OpenTextFile(Me.Path, 1 , true) ' 1 = ForReading
    If Err.Number = 0 Then
      debug("OK")
      LoadTextContents(File)
    Else
      Error("ファイル " & Me.Path & "を開くことができません" &  Err.Description)
    End If
    File.Close
    Set File = Nothing
  End Sub

  Private Sub LoadTextContents(FileObject)
    Dim line, Is1stLine, new_index
    debug("LoadTextContents")
    ReDim m_contents(0)
    Is1stLine = True
    Do While FileObject.AtEndOfStream <> True
      line = FileObject.ReadLine
      If IsValidLine(line) Then
        new_index = UBound(m_contents) + 1
        ReDim Preserve m_contents(new_index)
        m_contents(new_index) = line
      End If
    Loop
    Call InitMp3FileSettings
    Call InitMp3Elements(Me.Dir)
  End Sub
  
  Private Sub InitMp3FileSettings()
    Dim mp3file
    If UBound(m_contents) > 0 Then
      mp3file = m_contents(1)
      Me.Mp3File = m_contents(1)
      If Not CreateObject("Scripting.FileSystemObject").FileExists(Me.Mp3FileWithPath) Then
        Error("指定された音声ファイル(" & Me.Mp3FileWithPath & ")が存在しません")
      End If
      debug("MP3 file is " & Me.Mp3FileWithPath)
    End If
  End Sub

  Private Sub InitMp3Elements(dir)
    Dim new_index, line, i
    ReDim m_mp3_elements(0)
    If UBound(m_contents) > 1 Then
      For i = 2 To UBound(m_contents)
        new_index = UBound(m_mp3_elements) + 1
        ReDim Preserve m_mp3_elements(new_index)
        Set m_mp3_elements(new_index) = New Mp3Element
        m_mp3_elements(new_index).Line = m_contents(i)
        m_mp3_elements(new_index).Dir = Dir
      Next
    Else
      Error("音声項目が指定されていません")
    End If
  End Sub
  
  Private Function IsValidLine(line)
    Dim trimed_line
    trimed_line = LTrimEx(line)
    If Len(trimed_line) > 0 Then
      If Left(trimed_line, 1) <> "#" Then
        IsValidLine = True
        Exit Function
      End If
    End If
    IsValidLine = False
  End Function
  
  Private Function Error(Message)
    MsgBox(Message)
    WScript.Quit 1
  End Function

  Private m_path
  Private m_mp3_file
  Private m_contents()
  Private m_mp3_elements()
  
End Class

