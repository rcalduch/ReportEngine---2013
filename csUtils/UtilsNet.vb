Imports System.net

Public Class UtilsNet
  Shared Function GetIpAddres() As String
    Dim iphe As IPHostEntry = Dns.GetHostEntry(Dns.GetHostName())
    Dim ipAddr() As IPAddress = iphe.AddressList
    Dim Ip4 As String = Nothing
    For Each ip As IPAddress In ipAddr
      If ip.AddressFamily = Sockets.AddressFamily.InterNetwork Then
        Ip4 = ip.ToString
        Exit For
      End If
    Next
    Return Ip4
  End Function

End Class
