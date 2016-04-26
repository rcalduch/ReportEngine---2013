Imports csAppData
Imports csUtils
Imports csUtils.Utils
Imports System.IO

Public Class P00MN00_Main
    Private CancelProcess As Boolean
    Private smtpInfo As New InfoEMailServerStruct

    Private report As ReportBaseClass
    Private JobDone As Boolean

    Private watchfolderCTB As FileSystemWatcher
    Private watchfolderGST As FileSystemWatcher

#Region " General "

    Private Sub P00MN00_Main_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CancelProcess = False
        lblDebug.Visible = False

        watchfolderGST = New System.IO.FileSystemWatcher()
        watchfolderGST.Path = My.Settings.PathToMonitor
        watchfolderGST.Filter = "*.xml"

        AddHandler watchfolderGST.Created, AddressOf ProcessReports
        watchfolderGST.EnableRaisingEvents = True

        lblPTM.Text = "Monitoritzant: " + My.Settings.PathToMonitor

    End Sub

    Private Sub cmdSortir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSortir.Click
        If MessageBox.Show("Si tanqueu la aplicació es pararà el servei d'impressió de la aplicació de CUSTOM. Voleu sortir?", "Gestor de llistats de CUSTOM", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            AppData.CanLogout = True
            CancelProcess = True
            Me.Close()
        End If
    End Sub

    Private Sub lblTitle_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblTitle.DoubleClick
        If lblDebug.Visible Then
            lblDebug.Visible = False
            AppData.Debug = False
        Else
            lblDebug.Visible = True
            AppData.Debug = True
        End If
    End Sub

    Private Sub ProcessReports(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs)

        System.Threading.Thread.Sleep(1000)

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
                End Select

            Case "ct"

                Dim CodiLlistat As String
                Dim params As FetchXmlParameter

                params = New FetchXmlParameter(e.FullPath)
                CodiLlistat = CNull(params.GetValue("CodiLlistat"))

                Select Case CodiLlistat.ToLower
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
            IO.File.Delete(e.FullPath)
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
