Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_fac

  Public Function GetFactures(OrigenDades As String, Serie As String, deClient As String, aClient As String, deAny As String, aAny As String, deNumero As Integer, aNumero As Integer, deData As DateTime, aData As DateTime, ByVal EstatFactura As String) As DataTable

    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtFcp As New DataTable
    Dim cmd As String = String.Empty

    If OrigenDades = "actual" Then
      cmd = "SELECT * FROM gst_fac WHERE !DELETED() "
    Else
      cmd = "SELECT * FROM gst_hfc WHERE !DELETED() "
    End If

    If Not String.IsNullOrEmpty(deClient) Then
      cmd += String.Format("AND fc_codcli >= '{0}' ", deClient)
      If Not String.IsNullOrEmpty(aClient) Then
        cmd += String.Format("AND fc_codcli <= '{0}' ", aClient)
      End If
    End If

    If Not String.IsNullOrEmpty(Serie) Then
      cmd += String.Format("AND ALLTRIM(fc_serienif) = '{0}' ", Serie)
    End If

    If Not String.IsNullOrEmpty(deAny) Then
      cmd += String.Format("AND fc_any >= '{0}' ", deAny)
      If Not String.IsNullOrEmpty(aAny) Then
        cmd += String.Format("AND fc_any <= '{0}' ", aAny)
      End If
    End If


    If (deNumero > 0) Or (aNumero > 0) Then
      If ((deNumero > 0) And (aNumero = 0)) Or (deNumero = aNumero) Then
        cmd += String.Format("AND fc_numero = {0} ", deNumero)
      Else
        If deNumero > 0 Then
          cmd += String.Format("AND fc_numero >= {0} ", deNumero)
        End If
        If aNumero > 0 Then
          cmd += String.Format("AND fc_numero <= {0} ", aNumero)
        End If
      End If
    End If

    Select Case EstatFactura.Substring(0, 1).ToLower
      Case "t"
        ' Nothing
      Case "p"
        cmd += "AND !fc_impres "
    End Select

    If Not IsNullOrEmptyValue(deData) Then
      cmd += String.Format("AND fc_data >= {1}^{0:yyyy-MM-dd}{2} ", deData, "{", "}")
    End If

    If Not IsNullOrEmptyValue(aData) Then
      cmd += String.Format("AND fc_data <= {1}^{0:yyyy-MM-dd}{2} ", aData, "{", "}")
    End If

    cmd += " ORDER BY fc_numero"

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm
    Try
      OleDbComm.Connection.Open()
      oledbDa.Fill(dtFcp)
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fcp " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try
    Return dtFcp
  End Function

  Public Function GetFactura(origenDades As String, serie As String, any As String, Numero As Integer) As DataRow
    Dim oledbDa As New OleDbDataAdapter
    Dim OleDbComm As New OleDbCommand
    Dim dtFcp As New DataTable
    Dim cmd As String = String.Empty
    Dim Result As DataRow

    If origenDades = "actual" Then
      cmd = "SELECT * FROM gst_fac WHERE !DELETED() "
    Else
      cmd = "SELECT * FROM gst_hfc WHERE !DELETED() "
    End If

    cmd += String.Format("AND fc_serie = '{0}' ", serie)
    cmd += String.Format("AND fc_any = '{0}' ", any)
    cmd += String.Format("AND fc_numero = {0} ", Numero)

    OleDbComm.Connection = AppData.OleDbConnFAC
    OleDbComm.CommandText = cmd
    OleDbComm.CommandType = CommandType.Text

    oledbDa.SelectCommand = OleDbComm

    Try
      OleDbComm.Connection.Open()
      oledbDa.Fill(dtFcp)
    Catch ex As Exception
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fac " + ex.Message)
      Throw ex
    Finally
      OleDbComm.Connection.Close()
    End Try

    If dtFcp.Rows.Count > 0 Then
      Result = dtFcp.Rows(0)
    Else
      Result = Nothing
    End If

    Return Result

  End Function

End Class

