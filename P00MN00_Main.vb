Imports csAppData
Imports csUtils.Utils
Imports System.IO
Imports System.Threading

Public Class P00MN00_Main
    Private CancelProcess As Boolean
    Private smtpInfo As New InfoEMailServerStruct

    Private report As ReportBaseClass
    Private JobDone As Boolean

    Private watchfolderCTB As FileSystemWatcher
    Private watchfolderGST As FileSystemWatcher

#Region " General "

    Private Sub P00MN00_Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CancelProcess = False
        lblDebug.Visible = False

        watchfolderGST = New FileSystemWatcher()
        watchfolderGST.Path = My.Settings.PathToMonitor
        watchfolderGST.Filter = "*.xml"

        AddHandler watchfolderGST.Created, AddressOf ProcessReports
        watchfolderGST.EnableRaisingEvents = True

        lblPTM.Text = "Monitoritzant: " + My.Settings.PathToMonitor

    End Sub

    Private Sub cmdSortir_Click(sender As Object, e As EventArgs) Handles cmdSortir.Click
        If MessageBox.Show($"Si tanqueu la aplicació es pararà el servei d'impressió de la aplicació de CUSTOM. Voleu sortir?", $"Gestor de llistats de CUSTOM", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            AppData.CanLogout = True
            CancelProcess = True
            Me.Close()
        End If
    End Sub

    Private Sub lblTitle_DoubleClick(sender As Object, e As EventArgs) Handles lblTitle.DoubleClick
        If lblDebug.Visible Then
            lblDebug.Visible = False
            AppData.Debug = False
        Else
            lblDebug.Visible = True
            AppData.Debug = True
        End If
    End Sub

    Private Sub ProcessReports(source As Object, e As FileSystemEventArgs)

        Thread.Sleep(1000)

        JobDone = False
        DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90FAC0001A_Factura " + "ProcessReports: " + e.Name)

        Select Case e.Name.Substring(0, 2).ToLower
            Case "fc"
                Select Case My.Settings.ClientCustom
                    Case "000435"
                        ' ViesTendals
                        report = New R90FAC0001A_Factura
                    Case "201502"
                        ' Riualebre
                        report = New R90FAC0002A_Factura
                    Case "000400"
                        ' ipm
                        report = New R90FAC0003A_Factura
                End Select

            Case "ct"

                Dim codiLlistat As String
                Dim params As FetchXmlParameter

                params = New FetchXmlParameter(e.FullPath)
                codiLlistat = CNull(params.GetValue("CodiLlistat"))

                Select Case codiLlistat.ToLower
                    Case "ctb41ext"
                        ' extracte de comptes
                        report = New R90CTB0041A_Extracte
                    Case "ctb41ofi"
                        ' Diari oficial
                        report = New R90CTB0041A_DiariOficial
                    Case "ctb42sys"
                        ' Balanç sumes i saldos
                        report = New R90CTB0042A_SumesiSaldos
                    Case "ctb42sit"
                        ' Balanç de situació
                        report = New R90CTB0042A_Situacio
                    Case "ctb42exp"
                        ' Balanç de explotació
                        report = New R90CTB0042A_Explotacio
                    Case "ctb42bal"
                        ' Balanç de comptes anuals: balanç
                        report = New R90CTB0042A_Balanç
                    Case "ctb42pyg"
                        ' Balanç de comptes anuals: perdues i guanys
                        report = New R90CTB0042A_Perdues
                End Select
        End Select

        Try
            report.Execute(e.FullPath)
            JobDone = True
        Catch ex As Exception
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90FAC0001A_Factura " + ex.Message)
            JobDone = False
        Finally
            report = Nothing
        End Try


        If JobDone Then
            File.Delete(e.FullPath)
        End If

        Application.DoEvents()

    End Sub

#End Region

    Private Sub cmdTest_Click(sender As Object, e As EventArgs) Handles cmdTest.Click
        report = New R90FAC0001A_Factura
        report.Execute("")
        report = Nothing
        JobDone = True
    End Sub

End Class
