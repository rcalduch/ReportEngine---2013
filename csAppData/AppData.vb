Public Class AppData

    Public Shared AppName As String = ""
    Public Shared RegKey As String = ""
    Public Shared CustomerName As String = ""
    Public Shared Developer As Boolean
    Public Shared DeveloperDB As Boolean
    Public Shared CanLogout As Boolean = True
    Public Shared DeveloperName As String
    Public Shared Debug As Boolean

    Public Shared Servers As String = ""
    Public Shared ServerNameOrIp As String = ""
    Public Shared ServerNameOrIpBackup As String = ""
    Public Shared ServerCatalog As String
    Public Shared ServerCatalogCtb As String
    Public Shared ServerWeb As Boolean

    Public Shared TWAINUser As String = "Ramon Calduch"
    Public Shared TWAINEmail As String = "rcalduch@custom-sw.com"
    Public Shared TWAINRegCode As String = "Sf1kmlXO+N4NIm+nDexND7nKyZzWOMvJx+bMGYA5VwDgqciC3UqwEN/VGeDLyBRlMGyg/hXTFJWyf+jwbRw460SqhCbaPO6P6aeSDsbIQ1KeMLoNGyogHCqMnyYgDEjyNe4CAs4tBr6iuBSbOTX3sFntgwuqJav1dBT3Kah3YCM8"

    Public Shared FoxPath As String
    Public Shared OleDbExpedients As String
    Public Shared OleDbCtb As String
    Public Shared FoxCtbPath As String
    Public Shared FoxPtmPath As String
    Private Shared m_PrinterLastUsed As String = ""
    Public Shared SQLConnection As New SqlClient.SqlConnection

    Public Shared OleDbConnString As String = "Provider=VFPOLEDB.1;Collating Sequence=MACHINE;Data Source={0};Mode=ReadWrite|Share Deny None;"

    Public Shared OleDbConnRPT As OleDbConnection
    Public Shared OleDbConnFAC As OleDbConnection
    Public Shared OleDbConnCTB As OleDbConnection
    Public Shared OleDbConnEOS As OleDbConnection
    Public Shared OleDbConnEOS_ As OleDbConnection
    Public Shared OleDbConnPTM As OleDbConnection

    Public Shared SQLAppUserName As String
    Public Shared OleDBAcces As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=L:\vs.net\_csw\Gestions.cat\csImpressos\csImpressos.mdb"
    Private Shared mSqlAppPassword As String

    Public Shared Property SQLAppPassword() As String
        Get
            Return mSqlAppPassword
        End Get
        Set(ByVal value As String)
            mSqlAppPassword = value
        End Set
    End Property

    Public Shared AppIcon As System.Drawing.Icon
    Public Shared HelpNameSpace As String
    Public Shared RunningFromIDE As Boolean

    'Printers
    Public Shared Prt_Albarans As String
    Public Shared Prt_Factures As String
    Public Shared Prt_Llistats As String
    Public Shared Prt_Rebuts As String
    Public Shared Prt_LabelProducte As String
    Public Shared Prt_LabelEnvio As String
    Public Shared Prt_LabelGefco As String
    Public Shared Prt_LabelSEUR As String

    'Colors
    Public Shared TitleBackColor As System.Drawing.Color
    Public Shared TitleForeColor As System.Drawing.Color
    Public Shared GridEvenRowBackColor As System.Drawing.Color
    Public Shared GridEvenRowForeColor As System.Drawing.Color
    Public Shared GridHightLightRowBackColor As System.Drawing.Color
    Public Shared GridHightLightRowForeColor As System.Drawing.Color

    Public Shared UserName As String = ""
    Public Shared UserPassword As String = ""
    Public Shared UserIsAdmin As Boolean = False
    Public Shared UserIP As String = ""

    'e-mail
    Public Shared SMTPServer As String
    Public Shared SMTPUser As String
    Public Shared SMTPPassword As String
    'fax
    Public Shared EmailServerFax As String

    'Application Specific
    '====================================================================================
    Public Shared ShowAvisosOnStartUp As Boolean
    Public Shared Prf_LlistatPijama As Boolean

    Public Shared UltimaEmpresa As Integer
    Public Shared mCurrentEmpresaID As Integer
    Public Shared CurrentEmpresaName As String
    Public Shared CurrentEmpresaLogo As String
    Public Shared CurrentEmpresaLogoFax As String
    Public Shared LogoFax As String
    Public Shared CurrentEmpresaBackColor As System.Drawing.Color
    Public Shared CurrentEmpresaForeColor As System.Drawing.Color
    Public Shared Preferencies As New AppPreferencies

#Region " Directoris "

    ' Directoris aplicació
    ' ===================================================================================
    Private Shared mDocs_RootFolder As String
    Public Shared Property Docs_RootFolder() As String
        Get
            Return mDocs_RootFolder
        End Get
        Set(ByVal value As String)
            mDocs_RootFolder = value
        End Set
    End Property

    Private Shared mDocs_Client_Fotos As String = "CLIENT_FOTOS"
    Private Shared mDocs_Client_Documents As String = "CLIENT_DOCUMENTS"
    Private Shared mDocs_Documents_Cobro As String = "DOCUMENTS_COBRO"
    Private Shared mDocs_Producte_Fitxes As String = "PRODUCTE_FITXES"
    Private Shared mDocs_Producte_Fotos As String = "PRODUCTE_FOTOS"
    Private Shared mDocs_Produccio_Fotos As String = "PRODUCCIO_FOTOS"
    Private Shared mDocs_Labels_Info As String = "LABELS_INFO"
    Private Shared mDocs_App_Error_Screen As String = "APP_ERROR_SCREEN"
    Private Shared mDocs_EDI_Outbox As String = "EDI_OUTBOX"
    Private Shared mDocs_EDI_Inbox As String = "EDI_INBOX"
    Private Shared mDocs_Logo_Corporatiu As String = "LOGO_EMPRESA"
    Private Shared mDocs_Envios_Expedicions As String = "FITXERS_EXPEDICIONS"
    Private Shared mDocs_Envios_SEUR As String = "TRANSPORT_SEUR"
    Private Shared mDocs_Envios_SOUTO As String = "TRANSPORT_SOUTO"
    Private Shared mDocs_Envios_GEFCO As String = "TRANSPORT_GEFCO"
    Private Shared mDocs_Envios_TOURLINE As String = "TRANSPORT_TOURLINE"
    Private Shared mDocs_Envios_DHL As String = "TRANSPORT_TOURLINE"
    Private Shared mDocs_CSB As String = "REMESES_BANC"

    Public Shared ReadOnly Property Docs_Logo_Corporatiu(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Logo_Corporatiu, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Client_Fotos(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Client_Fotos, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Client_Documents(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Client_Documents, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Documents_Cobro(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Client_Documents, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Producte_Fitxes(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Producte_Fitxes, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Producte_Fotos(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Producte_Fotos, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Produccio_Fotos(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Produccio_Fotos, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Labels_Info(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_Labels_Info, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_App_Error_Screen(ByVal Filename As String) As String
        Get
            Return BuildPath(Docs_RootFolder, mDocs_App_Error_Screen, Filename)
        End Get
    End Property

    Public Shared ReadOnly Property Docs_EDI_Outbox() As String
        Get
            ' Overridable value
            Dim path As String
            path = Preferencies.DirectoriEdiOutbox.Valor.ToString
            If String.IsNullOrEmpty(path) Then
                path = BuildPath(Docs_RootFolder, mDocs_EDI_Outbox, String.Empty)
            End If
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_EDI_Inbox() As String
        Get
            ' Overridable value
            Dim path As String
            path = Preferencies.DirectoriEdiInbox.Valor.ToString
            If String.IsNullOrEmpty(path) Then
                path = BuildPath(Docs_RootFolder, mDocs_EDI_Inbox, String.Empty)
            End If
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Envios_SEUR(ByVal Filename As String) As String
        Get
            ' Overridable value
            Dim path As String
            'path = Preferencies.DirectoriFitxerSEUR.Valor.ToString
            'If String.IsNullOrEmpty(path) Then
            path = IO.Path.Combine(Docs_RootFolder, mDocs_Envios_Expedicions)
            path = IO.Path.Combine(path, "SEUR")
            path = IO.Path.Combine(path, Filename)
            'End If
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Envios_SOUTO(ByVal Filename As String) As String
        Get
            ' Overridable value
            Dim path As String
            'path = Preferencies.DirectoriFitxerSEUR.Valor.ToString
            'If String.IsNullOrEmpty(path) Then
            path = IO.Path.Combine(Docs_RootFolder, mDocs_Envios_Expedicions)
            path = IO.Path.Combine(path, "SOUTO")
            path = IO.Path.Combine(path, Filename)
            'End If
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Envios_GEFCO(ByVal Filename As String) As String
        Get
            ' Overridable value
            Dim path As String
            'path = Preferencies.DirectoriFitxerSEUR.Valor.ToString
            'If String.IsNullOrEmpty(path) Then
            path = IO.Path.Combine(Docs_RootFolder, mDocs_Envios_Expedicions)
            path = IO.Path.Combine(path, "GEFCO")
            path = IO.Path.Combine(path, Filename)
            'End If
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Envios_TOURLINE(ByVal Filename As String) As String
        Get
            Dim path As String
            path = IO.Path.Combine(Docs_RootFolder, mDocs_Envios_Expedicions)
            path = IO.Path.Combine(path, "TOURLINE")
            path = IO.Path.Combine(path, Filename)

            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_Envios_DHL(ByVal Filename As String) As String
        Get
            Dim path As String
            path = IO.Path.Combine(Docs_RootFolder, mDocs_Envios_Expedicions)
            path = IO.Path.Combine(path, "DHL")
            path = IO.Path.Combine(path, Filename)
            Return path
        End Get
    End Property

    Public Shared ReadOnly Property Docs_CSB(ByVal Filename As String) As String
        Get
            Dim path As String
            path = Preferencies.DirectoriCSB.Valor.ToString
            If String.IsNullOrEmpty(path) Then
                path = BuildPath(Docs_RootFolder, mDocs_CSB, Filename)
            End If
            Return path
        End Get
    End Property


    Private Shared Function BuildPath(ByVal Root As String, ByVal Path As String, ByVal File As String) As String
        Dim fp As String
        Dim fldr As String

        If String.IsNullOrEmpty(Path) Then
            fldr = Root
        Else
            fldr = IO.Path.Combine(Docs_RootFolder, Path)
        End If

        If Not My.Computer.FileSystem.DirectoryExists(fldr) Then
            My.Computer.FileSystem.CreateDirectory(fldr)
        End If

        If String.IsNullOrEmpty(File) Then
            fp = fldr
        Else
            fp = IO.Path.Combine(fldr, File)
        End If
        Return fp

    End Function

    Public Shared ReadOnly Property CurrentUserDocsFolder(ByVal Filename As String) As String
        Get
            Return BuildPath(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "misGestio", Filename)
        End Get
    End Property

    Public Shared ReadOnly Property CurrentUserDocsFolder() As String
        Get
            Return CurrentUserDocsFolder(String.Empty)
        End Get
    End Property
    ' =========================================================================================================
#End Region

    Shared ReadOnly Property SqlConnectionString() As String
        Get
            Dim cs As String
            Dim sn As String
            sn = ServerNameOrIp
            cs = String.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3}",
               sn,
               ServerCatalog,
               SQLAppUserName,
               SQLAppPassword)
            Return cs
        End Get
    End Property

    Shared ReadOnly Property SqlCtbConnectionString() As String
        Get
            Dim cs As String
            Dim sn As String
            sn = ServerNameOrIp
            cs = String.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3}",
               sn,
               ServerCatalogCtb,
               SQLAppUserName,
               SQLAppPassword)
            Return cs
        End Get
    End Property

    Shared ReadOnly Property SqlArsysConnectionString() As String
        Get
            Dim cs As String
            Dim sn As String
            sn = ServerNameOrIp
            cs = String.Format("Data Source={0};Initial Catalog={1};User ID={2};Password={3}",
               "osiris.servidoresdns.net",
               "qc414",
               "qc414",
               "misbd.01")
            Return cs
        End Get
    End Property

    Shared ReadOnly Property FoxConnectionString() As String
        Get
            Dim cs As String

            cs = String.Format("Provider=VFPOLEDB.1;Collating Sequence=MACHINE;Data Source={0}", FoxPath)

            Return cs
        End Get
    End Property

    Shared ReadOnly Property FoxCtbConnectionString() As String
        Get
            Dim cs As String

            ' El mateix. No te sentit que una aplicació nova gestioni dos interficis.
            cs = String.Format("Provider=VFPOLEDB.1;Collating Sequence=MACHINE;Data Source={0}", FoxCtbPath)

            Return cs
        End Get
    End Property

    Shared ReadOnly Property FoxPtmConnectionString() As String
        Get
            Dim cs As String
            ' El mateix. No te sentit que una aplicació nova gestioni dos interficis.
            cs = String.Format("Provider=VFPOLEDB.1;Collating Sequence=MACHINE;Data Source={0}", FoxPtmPath)
            Return cs
        End Get
    End Property

    Shared Property CurrentEmpresaID() As Integer
        Get
            Return mCurrentEmpresaID
        End Get
        Set(ByVal value As Integer)
            mCurrentEmpresaID = value
            AppData.LogoFax = "fax.jpg"
            Select Case value
                Case 1
                    ' Mistral bonsai
                    AppData.CurrentEmpresaForeColor = Drawing.Color.White
                    AppData.TitleForeColor = AppData.CurrentEmpresaForeColor
                    If AppData.DeveloperDB Then
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-16754194)
                    Else
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-16754094)
                    End If
                    AppData.TitleBackColor = AppData.CurrentEmpresaBackColor
                    AppData.CurrentEmpresaLogo = "logfaxmistral.jpg"
                    AppData.CurrentEmpresaLogoFax = "logfaxmistral.jpg"
                Case 2
                    ' Jardin Press
                    AppData.CurrentEmpresaForeColor = Drawing.Color.White
                    AppData.TitleForeColor = AppData.CurrentEmpresaForeColor
                    If AppData.DeveloperDB Then
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-14051134)
                    Else
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-14051034)
                    End If

                    AppData.TitleBackColor = AppData.CurrentEmpresaBackColor
                    AppData.CurrentEmpresaLogo = "logfaxjardin.jpg"
                    AppData.CurrentEmpresaLogoFax = "logfaxjardin.jpg"
                Case 3
                    ' Amazonia
                    AppData.CurrentEmpresaForeColor = Drawing.Color.White
                    AppData.TitleForeColor = AppData.CurrentEmpresaForeColor
                    If AppData.DeveloperDB Then
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-4748643)
                    Else
                        AppData.CurrentEmpresaBackColor = System.Drawing.Color.FromArgb(-14051034)
                    End If

                    AppData.TitleBackColor = AppData.CurrentEmpresaBackColor
                    AppData.CurrentEmpresaLogo = "logfaxamazonia.jpg"
                    AppData.CurrentEmpresaLogoFax = "logfaxamazonia.jpg"
            End Select
        End Set
    End Property

    Public Shared Function GetSerie(ByVal Serie As String) As String
        Dim sInt As Integer = CInt(Serie) + CInt(Math.Pow(10, Serie.Length - 1))
        Return sInt.ToString
    End Function

    Public Shared Sub GetAppDataFromRegistry()
        '
        '==========================================================================================
        '    U L L ! ! !     A C T U A L I T Z A R     f r m P r e r e n c i e s
        '==========================================================================================
        '
        Dim key As Microsoft.Win32.RegistryKey
        Dim regValue As String

        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)

        AppData.FoxPath = Convert.ToString(key.GetValue("FoxPath", ""))
        AppData.FoxCtbPath = Convert.ToString(key.GetValue("FoxCtbPath", ""))
        AppData.FoxPtmPath = Convert.ToString(key.GetValue("FoxPtmPath", ""))
        AppData.Servers = key.GetValue("Servers", "").ToString
        AppData.ServerNameOrIp = Convert.ToString(key.GetValue("ServerNameOrIp", ""))
        AppData.ServerNameOrIpBackup = Convert.ToString(key.GetValue("ServerNameOrIpBackup", ""))
        AppData.UserName = Convert.ToString(key.GetValue("UserName", ""))
        AppData.Developer = (CInt(key.GetValue("Developer", 0)) = 58)
        AppData.DeveloperName = key.GetValue("DeveloperName", "").ToString

        regValue = Convert.ToString(key.GetValue("SQLAppUserName", ""))
        If regValue <> String.Empty Then
            AppData.SQLAppUserName = regValue
        End If
        regValue = Convert.ToString(key.GetValue("SQLAppPassword", ""))
        If regValue <> String.Empty Then
            AppData.SQLAppPassword = regValue
        End If

        'OleDB

        'Printers
        AppData.Prt_Albarans = Convert.ToString(key.GetValue("Prt_Albarans", ""))
        AppData.Prt_Factures = Convert.ToString(key.GetValue("Prt_Factures", ""))
        AppData.Prt_Llistats = Convert.ToString(key.GetValue("Prt_Llistats", ""))
        AppData.Prt_Rebuts = Convert.ToString(key.GetValue("Prt_Rebuts", ""))
        AppData.Prt_LabelProducte = Convert.ToString(key.GetValue("Prt_LabelProducte", ""))
        AppData.Prt_LabelEnvio = Convert.ToString(key.GetValue("Prt_LabelEnvio", ""))
        AppData.Prt_LabelGefco = Convert.ToString(key.GetValue("Prt_LabelGefco", ""))
        AppData.Prt_LabelSEUR = Convert.ToString(key.GetValue("Prt_LabelSEUR", ""))

        'Colors
        'AppData.TitleXXXXColor depenen de la empresa...
        'AppData.TitleBackColor = System.Drawing.Color.FromArgb(CInt(key.GetValue("Color.TitleBackColor", -32768)))
        'AppData.TitleForeColor = System.Drawing.Color.FromArgb(CInt(key.GetValue("Color.TitleForeColor", -1)))

        Dim NumColor As Integer
        NumColor = CInt(key.GetValue("Color.GridEvenRowBackColor", -1379587))
        AppData.GridEvenRowBackColor = System.Drawing.Color.FromArgb(NumColor)

        NumColor = CInt(key.GetValue("Color.GridEvenRowForeColor", -16777216))
        AppData.GridEvenRowForeColor = System.Drawing.Color.FromArgb(NumColor)

        NumColor = CInt(key.GetValue("Color.GridHightLightRowBackColor", -2302756))
        AppData.GridHightLightRowBackColor = System.Drawing.Color.FromArgb(NumColor)

        NumColor = CInt(key.GetValue("Color.GridHightLightRowForeColor", -16777216))
        AppData.GridHightLightRowForeColor = System.Drawing.Color.FromArgb(NumColor)

        'Help
        AppData.HelpNameSpace = Convert.ToString(key.GetValue("HelpFile", ""))

        'Specific
        AppData.ShowAvisosOnStartUp = Convert.ToBoolean(key.GetValue("ShowAvisosOnStartUp"))
        AppData.Prf_LlistatPijama = Convert.ToBoolean(key.GetValue("Prf_LlistatPijama"))

        'e-mail
        'AppData.SMTPMail = Convert.ToString(key.GetValue("SMTPe-Mail", ""))
        'AppData.SMTPServer = Convert.ToString(key.GetValue("SMTPServer", ""))


        'Global vars (gv.XxxxxYyyyyZzzzzz)
        AppData.UltimaEmpresa = CInt(key.GetValue("gv.UltimaEmpresaTreballada", 1))
        AppData.CurrentEmpresaID = CInt(key.GetValue("gv.CurrentEmpresaID", 1))
        AppData.CurrentEmpresaName = key.GetValue("gv.CurrentEmpresaName", "").ToString
        'AppData.TitleBackColor = Drawing.Color.Orange
        'AppData.TitleForeColor = Drawing.Color.White


        ' End
        key.Close()

    End Sub

    Public Shared Sub SaveAppData()

        SetValueIntoRegistry("gv.UltimaEmpresaTreballada", AppData.UltimaEmpresa)
        SetValueIntoRegistry("gv.CurrentEmpresaID", AppData.UltimaEmpresa)

    End Sub

    Public Shared Function GetValueFromRegistry(ByVal KeyName As String, ByVal DefaultValue As Object) As Object
        Dim Value As Object
        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
        Value = key.GetValue(KeyName, DefaultValue)
        key.Close()
        Return Value
    End Function

    Public Shared Sub SetValueIntoRegistry(ByVal KeyName As String, ByVal Value As Object)
        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
        key.SetValue(KeyName, Value)
        key.Close()
    End Sub

End Class
