Imports csUtils
Imports csUtils.Utils
Imports System.Data.OleDb
Imports csAppData

Public Class C00_gst_sys

  Public Function GetSystemValue(ByVal Item As String) As Object
    Dim OleDbComm As New OleDbCommand
    Dim cmd As String
    Dim row As OleDbDataReader
    Dim retValue As Object

    cmd = String.Format("SELECT * FROM exp_sys WHERE st_data = '{0}'", Item)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    Try
      OleDbComm.Connection.Open()
      row = OleDbComm.ExecuteReader
      If row.Read() Then
        Select Case row("st_tipo").ToString
          Case "C"
            retValue = row("st_char").ToString.Trim
          Case "N"
            retValue = CNull(row("st_num"), 0D)
          Case Else
            retValue = Nothing
        End Select
      Else
        retValue = Nothing
      End If
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_exp_sis " + ex.Message)
      retValue = Nothing
    Finally
      OleDbComm.Connection.Close()
    End Try

    Return retValue

  End Function

End Class

