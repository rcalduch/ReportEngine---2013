' GetDotNetCode Play WAV With Widows Forms

Public Class WavPlayer

  ' You must use some Windows API to play a WAV file.
  Const SND_ASYNC As Integer = &H1
  ' Import the Windows PlaySound fuction API.
  <System.Runtime.InteropServices.DllImport("winmm.dll")> _
  Shared Function PlaySound(ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
  End Function

  Function PlayWAVAsynchronously(ByVal fileName As String) As Boolean
    ' Play asynchronously.
    Return (PlaySound(fileName, 0, SND_ASYNC) = 0)
  End Function

  Function PlayWAV(ByVal fileName As String, ByVal synchronousMode As Boolean) As Boolean
    ' Play asnhchronously or synchronously depending on synchronousMode parameter
    ' received by this funtion.
    If synchronousMode Then
      ' Play asynchronously.
      Return (PlaySound(fileName, 0, 0) = 0)
    Else
      ' Play synchronously.
      Return (PlaySound(fileName, 0, SND_ASYNC) = 0)
    End If
  End Function

End Class
