Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_cli

  Public Function GetClientByNIF(ByVal NifClient As String) As DataTable
    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtFln As New DataTable
    Dim cmd As String

    cmd = String.Format("SELECT * FROM gst_cli WHERE !DELETED() AND cl_nif = '{0}' ", NifClient)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm
    OleDbComm.Connection.Open()
    Try
      oledbDa.Fill(dtFln)
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_cli-GetClientByNIF: " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try
    Return dtFln
  End Function

  Public Function Lookup(ByVal CodiClient As String, ByVal CampRetorn As String) As Object
    Dim oledbComm As New OleDbCommand
    Dim cmd As String
    Dim retValue As Object

    cmd = String.Format("SELECT {1} FROM gst_cli WHERE cl_codcli = '{0}' ", CodiClient, CampRetorn)

    oledbComm.Connection = AppData.OleDbConnFAC
    oledbComm.CommandText = cmd
    oledbComm.CommandType = CommandType.Text

    oledbComm.Connection.Open()
    Try
      retValue = oledbComm.ExecuteScalar
    Catch ex As Exception
      ' Throw ex
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_cli-Lookup " + ex.Message)
      retValue = Nothing
    Finally
      oledbComm.Connection.Close()
    End Try

    Return retValue

  End Function

End Class

