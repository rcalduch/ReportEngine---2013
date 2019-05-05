Imports csAppData
Imports System.Data.OleDb
Imports csUtils.Utils

Public Class C00_gst_fac

    Public Function GetFactures(origenDades As String, serie As String, deClient As String, aClient As String, deAny As String, aAny As String, deNumero As Integer, aNumero As Integer, deData As DateTime, aData As DateTime, estatFactura As String) As DataTable

        Dim oledbDa As New OleDbDataAdapter
        Dim oleDbComm As New OleDbCommand
        Dim dtFcp As New DataTable
        Dim cmd As String = String.Empty

        If origenDades = "actual" Then
            cmd = "SELECT * FROM gst_fac WHERE !DELETED() "
        Else
            cmd = "SELECT * FROM gst_hfc WHERE !DELETED() "
        End If

        If Not String.IsNullOrEmpty(deClient) Then
            cmd += $"AND fc_codcli >= '{deClient}' "
            If Not String.IsNullOrEmpty(aClient) Then
                cmd += $"AND fc_codcli <= '{aClient}' "
            End If
        End If

        If Not String.IsNullOrEmpty(serie) Then
            cmd += $"AND ALLTRIM(fc_serienif) = '{serie}' "
        End If

        If Not String.IsNullOrEmpty(deAny) Then
            cmd += $"AND fc_any >= '{deAny}' "
            If Not String.IsNullOrEmpty(aAny) Then
                cmd += $"AND fc_any <= '{aAny}' "
            End If
        End If


        If (deNumero > 0) Or (aNumero > 0) Then
            If ((deNumero > 0) And (aNumero = 0)) Or (deNumero = aNumero) Then
                cmd += $"AND fc_numero = {deNumero} "
            Else
                If deNumero > 0 Then
                    cmd += $"AND fc_numero >= {deNumero} "
                End If
                If aNumero > 0 Then
                    cmd += $"AND fc_numero <= {aNumero} "
                End If
            End If
        End If

        Select Case estatFactura.Substring(0, 1).ToLower
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

        oleDbComm.Connection = AppData.OleDbConnFAC
        oleDbComm.CommandText = cmd
        oleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = oleDbComm
        Try
            oleDbComm.Connection.Open()
            oledbDa.Fill(dtFcp)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fcp " + ex.Message)
            Throw ex
        Finally
            oleDbComm.Connection.Close()
        End Try
        Return dtFcp
    End Function

    Public Function GetFactura(origenDades As String, serie As String, any As String, numero As Integer) As DataRow
        Dim oledbDa As New OleDbDataAdapter
        Dim oleDbComm As New OleDbCommand
        Dim dtFcp As New DataTable
        Dim cmd As String = String.Empty
        Dim result As DataRow

        If origenDades = "actual" Then
            cmd = "SELECT * FROM gst_fac WHERE !DELETED() "
        Else
            cmd = "SELECT * FROM gst_hfc WHERE !DELETED() "
        End If

        cmd += $"AND fc_serie = '{serie}' AND fc_any = '{any}' AND fc_numero = {numero} "

        oleDbComm.Connection = AppData.OleDbConnFAC
        oleDbComm.CommandText = cmd
        oleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = oleDbComm

        Try
            oleDbComm.Connection.Open()
            oledbDa.Fill(dtFcp)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fac " + ex.Message)
            Throw ex
        Finally
            oleDbComm.Connection.Close()
        End Try

        If dtFcp.Rows.Count > 0 Then
            result = dtFcp.Rows(0)
        Else
            result = Nothing
        End If

        Return result

    End Function

End Class

