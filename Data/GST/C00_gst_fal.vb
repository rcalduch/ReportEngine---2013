Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_fal

  Public Function GetLiniesFactura(origenDades As String, serie As String, any As String, Numero As Integer) As DataTable
    Dim DA As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    'Dim dr As OleDbDataReader
    Dim cmd As String
    Dim dtLin As New DataTable("lin")

    If origenDades = "actual" Then
      cmd = String.Format("SELECT * FROM gst_fal WHERE !DELETED() AND fl_serie = '{0}' and fl_any = '{1}' AND fl_numero = {2} ", serie, any, Numero)
    Else
      cmd = String.Format("SELECT * FROM gst_hfl WHERE !DELETED() AND fl_serie = '{0}' and fl_any = '{1}' AND fl_numero = {2} ", serie, any, Numero)
    End If

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    DA.SelectCommand = OleDbComm
    OleDbComm.Connection.Open()
    Try
      DA.Fill(dtLin)
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fal " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try

    Return dtLin

  End Function

End Class

