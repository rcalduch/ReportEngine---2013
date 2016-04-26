Imports Microsoft.Win32
Imports csAppData
Imports csUtils

Module StartUp

    <STAThread()>
    Sub Main(ByVal args() As String)

        AddHandler Application.ThreadException, AddressOf App_ThreadException

        Application.EnableVisualStyles()

        Application.DoEvents()

        SetUpAppData()

        SetDebugMode()

        Dim frm As New P00MN00_Main

        Application.Run(frm)

        frm.Dispose()

        AppData.SaveAppData()

        Application.Exit()

    End Sub

    Private Sub SetDebugMode()
        ' This will only be done if in the IDE
        Debug.Assert(InDebugMode)
    End Sub

    Private Function InDebugMode() As Boolean
        AppData.RunningFromIDE = True
        Return True
    End Function

    Private Sub SetUpAppData()

        AppData.CustomerName = ""
        AppData.CurrentEmpresaName = ""
        AppData.AppName = "REPORT ENGINE"
        AppData.RegKey = "Software\Custom Software\csDosReportEngine"
        AppData.ServerCatalog = "csGestionsCat"
        AppData.SQLAppUserName = "Custom"
        AppData.SQLAppPassword = "Custom.89"

        AppData.OleDbConnFAC = New OleDbConnection(String.Format(AppData.OleDbConnString, My.Settings.OleDbDirFAC))
        AppData.OleDbConnCTB = New OleDbConnection(String.Format(AppData.OleDbConnString, My.Settings.OleDbDirCTB))

    End Sub

    ' Default Exception Handler
    Sub App_ThreadException(ByVal sender As Object, ByVal e As System.Threading.ThreadExceptionEventArgs)

        MessageBox.Show("La aplicació es tancarà " + e.Exception.Message, "Error greu", MessageBoxButtons.OK, MessageBoxIcon.Error)

        Try

            ' Log error into database
            Dim TexteError As String
            Dim StackTrace As String
            Dim ErrorLogID As Integer
            Dim IpAdr As String
            Dim ImgScr As Image
            Dim dialogRes As DialogResult

            ImgScr = csScreenImage.GetScreenSnapshot()

            If e.Exception.Message.Length > 8000 Then
                TexteError = e.Exception.Message.Substring(0, 8000)
            Else
                TexteError = e.Exception.Message
            End If
            If e.Exception.StackTrace.Length > 8000 Then
                StackTrace = e.Exception.StackTrace.Substring(0, 8000)
            Else
                StackTrace = e.Exception.StackTrace
            End If

            IpAdr = String.Empty
            Try
                IpAdr = UtilsNet.GetIpAddres
            Catch ex As Exception
                If AppData.Debug Then
                    MessageBox.Show("Error: " + IpAdr + ". Avisar a dep. de sistemes." + vbCrLf + vbCrLf + ex.ToString)
                End If
            End Try

            If ErrorLogID > 0 Then

            End If

            If dialogRes = DialogResult.Abort Then
                Application.Exit()
            End If

        Catch ex As Exception
            Try
            Finally
                MessageBox.Show("La aplicació es tancarà" + vbCrLf + vbCrLf + ex.Message, "Error greu", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Application.Exit()
            End Try
        End Try
    End Sub

    Public Class FetchXmlParameter
        Private params As New Dictionary(Of String, String)

        Public Sub New(ByVal XmlFileName As String)

            Dim xml As New Xml.XmlDocument
            Dim doc As Xml.XmlElement
            Dim value As String

            ' MsgBox(XmlFileName)

            xml.Load(XmlFileName)
            doc = xml.DocumentElement

            For Each n As Xml.XmlNode In doc.ChildNodes
                Try
                    value = n.FirstChild.Value.Trim
                Catch ex As Exception
                    value = String.Empty
                End Try
                params.Add(n.Name, value)
            Next

        End Sub

        Public Function GetValue(ByVal ParameterName As String) As String
            Dim ParamValue As String

            Try
                ParamValue = params(ParameterName)
            Catch ex As Exception
                ParamValue = String.Empty
            End Try

            Return ParamValue

        End Function

    End Class

End Module

