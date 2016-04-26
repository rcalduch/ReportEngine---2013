Friend Class daUtils

  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As DateTime) As DateTime
    If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
      Return DefaultValue
    End If
    Return Convert.ToDateTime(Value)
  End Function

  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Decimal) As Decimal
    If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
      Return DefaultValue
    End If
    Return Convert.ToDecimal(Value)
  End Function

  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Integer) As Integer
    If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
      Return DefaultValue
    End If
    Return Convert.ToInt32(Value)
  End Function


  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As String) As String
    If Value Is DBNull.Value OrElse Value Is Nothing Then
      Return DefaultValue
    End If
    Return CStr(Value)
  End Function

  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Object) As Object
    If Value Is DBNull.Value OrElse Value Is Nothing Then
      Return DefaultValue
    End If
    Return Value
  End Function

  Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Boolean) As Boolean
    If Value Is DBNull.Value OrElse Value Is Nothing Then
      Return DefaultValue
    End If
    Return CBool(Value)
  End Function

  Shared Function CNull(ByVal Value As Object) As String
    If Value Is DBNull.Value OrElse Value Is Nothing Then
      Return ""
    End If
    Return CStr(Value)
  End Function

  Shared Function CDNull(ByVal Value As Object, Optional ByVal FormatString As String = "dd/MM/yyyy") As String
    If Value Is DBNull.Value OrElse Value Is Nothing Then
      Return ""
    End If
    Return CDate(Value).ToString(FormatString)
  End Function


End Class
