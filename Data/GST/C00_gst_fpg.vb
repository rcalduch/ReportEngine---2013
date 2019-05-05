Imports csAppData
Imports csUtils.Utils

Public Class C00_gst_fpg

    Public Function GetFormaPagament(codiFormaPagament As String) As String
        Dim oledbDa As New OleDbDataAdapter
        Dim oleDbComm As New OleDbCommand
        Dim cmd As String
        Dim formaPagament As String

        cmd = $"SELECT fp_desc FROM gst_fpg WHERE !DELETED() AND fp_codi = '{codiFormaPagament}' "

        oleDbComm.Connection = AppData.OleDbConnFAC
        oleDbComm.CommandText = cmd
        oleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = oleDbComm
        Try
            oleDbComm.Connection.Open()
            formaPagament = CNull(oleDbComm.ExecuteScalar())
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fcp " + ex.Message)
            Throw ex
        Finally
            oleDbComm.Connection.Close()
        End Try
        Return formaPagament

    End Function

End Class

