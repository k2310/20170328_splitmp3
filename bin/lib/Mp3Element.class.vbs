' ======================================================================
'   Class Mp3Element
' ======================================================================
Class Mp3Element

  Public Property Let Line(value)
    m_line = value
    Call InitProperty
  End Property
  Public Property Get Line
    Line = m_line
  End Property

  Public Property Let Dir(value)
    m_dir = value
  End Property
  Public Property Get Dir
    Dir = m_dir
  End Property

  Public Property Get Name
    Name = m_name
  End Property

  Public Property Get FileName
    FileName = Me.Name & ".mp3"
  End Property

  Public Property Get FilePath
    FilePath = Me.Dir & "\" & Me.FileName
  End Property

  Public Function IsExists()
	If CreateObject("Scripting.FileSystemObject").FileExists(Me.FilePath) Then
	  IsExists = True
	Else
	  IsExists = False
	End If
  End Function
  

  Public Property Get TimeBegin
    TimeBegin = m_time_begin
  End Property
  Public Property Get TimeEnd
    TimeEnd = m_time_end
  End Property

  Public Property Get TimeBeginSec
    TimeBeginSec = m_time_begin_sec
  End Property
  Public Property Get TimeEndSec
    TimeEndSec = m_time_end_sec
  End Property

  Private Sub InitProperty()
    Dim elements, buf
    elements = Split(Me.Line, ",")
    If UBound(elements) < 2 Then
      Error("行[" & Me.Line & "]は書式違反です。" & VbCrLf & "    書式: <開始時間>,<終了時間>,<ファイル名(.mp3は除く)")
    End If

    ' ----- Begin Time
    buf = RemoveSpace(elements(0))
    If Not IsTimeFormat(buf) Then
      Error("行中 [" & buf & "]が時間書式になっていません。")
    Else
      m_time_begin = buf
      m_time_begin_sec = GetSecond(buf)
    End If

    ' ----- End Time
    buf = RemoveSpace(elements(1))
    If Not IsTimeFormat(buf) Then
      Error("行中 [" & buf & "]が時間書式になっていません。")
    Else
      m_time_end = buf
      m_time_end_sec = GetSecond(buf)
    End If

    ' ----- File Name
    buf = RemoveSpace(elements(2))
    m_name = buf
  End Sub

  Private Function GetSecond(timestr)
    GetSecond = DateDiff("s", 0, timestr)
  End Function
  
  Private Function IsTimeFormat(str)
    Dim re
    Set re = New RegExp
    re.IgnoreCase = True
    re.pattern = "^\d{1,2}:\d{1,2}:\d{1,2}$"
    IsTimeFormat = re.Test(str)
  End Function

  Private m_line
  Private m_dir
  Private m_name
  Private m_time_begin
  Private m_time_end
  Private m_time_begin_sec
  Private m_time_end_sec

End Class


