Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_cim

  Public Function get_page(report_id As String) As DataRow
    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtCim As New DataTable
    Dim cmd As String
    Dim prop_dr As DataRow = Nothing

    cmd = String.Format("SELECT * FROM gst_cim WHERE !DELETED() AND ic_codi = '{0}' ", report_id)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm
    OleDbComm.Connection.Open()
    Try
      oledbDa.Fill(dtCim)
      If dtCim.Rows.Count > 0 Then
        prop_dr = dtCim.Rows(0)
      End If
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_cim-get_page: " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try
    Return prop_dr
  End Function

  Public Function get_fields(report_id As String) As DataTable
    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtLim As New DataTable
    Dim cmd As String

    cmd = String.Format("SELECT * FROM gst_cim WHERE !DELETED() AND ic_codi = '{0}' ", report_id)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm
    OleDbComm.Connection.Open()
    Try
      oledbDa.Fill(dtLim)
     
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_cim-get_fields: " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try
    Return dtLim
  End Function

End Class

