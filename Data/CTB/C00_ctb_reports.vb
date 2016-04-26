Imports csAppData
Imports csUtils.Utils

Public Class C00_ctb_reports
    Public Function get_extracte_comptes() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb41ext"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_ctb_get_extracte_comptes: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

    Public Function get_diari_oficial() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb41ofi"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_diari_oficial: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function
    Public Function get_balans_sumes_saldos() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb42sys"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_balans_sumes_saldos: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

    Public Function get_balans_situacio() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb42sit"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_balans_situacio: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

    Public Function get_balans_explotacio() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb42exp"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_balans_explotacio: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

    Public Function get_comptes_anuals_balans() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb42bal"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_comptes_anuals_balans: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

    Public Function get_comptes_anuals_perdues() As DataTable
        Dim oledbDa As New OleDbDataAdapter
        Dim OleDbComm As New OleDbCommand
        Dim dtExt As New DataTable
        Dim cmd As String

        cmd = "SELECT * FROM ctb42pyg"

        OleDbComm.Connection = AppData.OleDbConnCTB
        OleDbComm.CommandText = cmd
        OleDbComm.CommandType = CommandType.Text

        oledbDa.SelectCommand = OleDbComm
        OleDbComm.Connection.Open()
        Try
            oledbDa.Fill(dtExt)
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " get_comptes_anuals_perdues: " + ex.Message)
            Throw ex
        Finally
            OleDbComm.Connection.Close()
        End Try
        Return dtExt
    End Function

End Class
