Module csBenchMark
  Delegate Sub BenchmarkDelegate()

  Enum BenchmarkModes
    DontShow
    Console
    MessageBox
  End Enum

  Public Function BenchmarkIt(ByVal routine As BenchmarkDelegate, _
      ByVal mode As BenchmarkModes, Optional ByVal msg As String = Nothing) As _
      TimeSpan
    ' remember starting time
    Dim start As Date = Now
    ' run the procedure to be benchmarked
    routine.Invoke()
    ' evaluate elapsed time, assign to result
    Dim elapsed As TimeSpan = Now.Subtract(start)
    ' exit if nothing else to do
    If mode = BenchmarkModes.DontShow Then Return elapsed

    ' build a suitable string if none was provided
    If msg Is Nothing OrElse msg.Length = 0 Then
      ' use the name of the target method
      msg = routine.Method.Name
    End If
    ' append a placeholder for elapsed time, if not there
    If msg.IndexOf("{0}") < 0 Then
      msg &= ": {0} secs"
    End If
    ' display on the console window or a message box
    If mode = BenchmarkModes.Console Then
      Console.WriteLine(msg, elapsed)
    ElseIf mode = BenchmarkModes.MessageBox Then
      System.Windows.Forms.MessageBox.Show(String.Format(msg, elapsed))
    End If
    ' return result to caller
    Return elapsed
  End Function

End Module
