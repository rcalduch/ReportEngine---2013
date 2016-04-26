Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_reb

  Public Function GetRebuts(cSerie As String, cAny As String, nNumero As Integer) As DataTable

    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtReb As New DataTable
    Dim cmd As String

    cmd = String.Format("SELECT * FROM gst_reb WHERE !DELETED() AND re_tipo = 'C' AND re_serie = '{0}' AND re_any = '{1}' AND re_numero = {2} ORDER BY re_nvto", cSerie, cAny, nNumero)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm
    OleDbComm.Connection.Open()
    Try
      oledbDa.Fill(dtReb)
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_reb" + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try
    Return dtReb
  End Function

End Class
