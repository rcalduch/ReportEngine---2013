Public Class AppPreferencies

  Public ServidorSQL As New Preferencia("Acces a dades", "Servidor SQL", "ServerNameOrIP", "String", "Nom del servidor SQL o IP del mateix")
  Public ServidorSQLalternatiu As New Preferencia("Acces a dades", "Servidor SQL alternatiu", "ServerNameOrIPBackup", "String", "Nom del servidor SQL o IP alternatiu")
  Public DirectoriCust_GST As New Preferencia("Acces a dades", "Dades FoxPro", "FoxPath", "String", "Directori de les dades de la aplicació antiga")
  Public DirectoriCust_CTB As New Preferencia("Acces a dades", "Dades CTB FoxPro", "FoxCtbPath", "String", "Directori de les dades de la comptabilitat antiga")

  Public Usuari As New Preferencia("Identificacio", "Usuari", "UserName", "String", "Usuari ultim accès a la aplicació")

  Public ImpressoraLListats As New Preferencia("Impressores", "Impressora Llistats", "Prt_Llistats", "PrinterName", "Impressora on imprimir els lliestats")
  Public ImpressoraALbarans As New Preferencia("Impressores", "Impressora Albarans", "Prt_Albarans", "PrinterName", "Impressora on imprimir els albarans")
  Public ImpressoraFactures As New Preferencia("Impressores", "Impressora Factures", "Prt_Factures", "PrinterName", "Impressora on imprimir les factures")
  Public ImpressoraRebuts As New Preferencia("Impressores", "Impressora Rebuts", "Prt_Rebuts", "PrinterName", "Impressora on imprimir els rebuts")
  Public ImpressoraDocuments As New Preferencia("Impressores", "Impressora Documents", "Prt_Documents", "PrinterName", "Impressora on imprimir els documents escanejats")
  Public ImpressoraColor As New Preferencia("Impressores", "Impressora Color", "Prt_Color", "PrinterName", "Impressora on imprimir els documents en color")

  Public ImpressoraEtiquetaPetita As New Preferencia("Impressores", "Impressora Etiquetes producte", "Prt_LabelProducte", "PrinterName", "Impressora on imprimir etiquete producte")
  Public ImpressoraEtiquetaGran As New Preferencia("Impressores", "Impressora Etiquetes enviament", "Prt_LabelEnvio", "PrinterName", "Impressora on imprimir etiquetes DIN A5 d'enviament")
  Public ImpressoraEtiquetaMitjana As New Preferencia("Impressores", "Impressora Etiquetes (10x10cm)", "Prt_LabelGefco", "PrinterName", "Impressora on imprimir etiquetes 10x10cm")
  Public ImpressoraEnviamentSEUR As New Preferencia("Impressores", "Impressora enviament SEUR", "Prt_LabelSEUR", "PrinterName", "Impressora on imprimir etiquetes enviament SEUR")

  Public ColorFonsTitol As New Preferencia("Colors titol", "Color de fons del Titol", "Color.TitleBackColor", "Color", "Color de fons de la barra del titol als formularis")
  Public ColorTexteTitol As New Preferencia("Colors titol", "Color texte del Titol", "Color.TitleForeColor", "Color", "Color de la lletra de la barra del titol als formularis")

  Public ColorGridFonsEvenRow As New Preferencia("Colors graella", "Color fons de la fila parell", "Color.GridEvenRowBackColor", "Color", "Color de fons de la fila senàs de les grelles")
  Public ColorGridTexteEvenRow As New Preferencia("Colors graella", "Color text fila Parell", "Color.GridEvenRowForeColor", "Color", "Color de la lletra de la fila senàs de les grelles")
  Public ColorGridFonsCurrentRow As New Preferencia("Colors graella", "Color fons de la fila actual", "Color.GridHightLightRowBackColor", "Color", "Color de fons de la fila actual")
  Public ColorGridTexteCurrentRow As New Preferencia("Colors graella", "Color texte fila actual", "Color.GridHightLightRowForeColor", "Color", "Color de la lletra de la fila actual")

  Public pdfOutputPath As New Preferencia("Opcions PDF", "Directori fitxers PDF", "pdfOutputPath", "FolderName", "Directori fixers generats al imprimir PDFs")
  Public pdfAutoSave As New Preferencia("Opcions PDF", "Dessar PDF al directori per defacte", "pdfAutoSave", "Boolean", "Desar documents PDF al directori per defacte")
  Public pdfOpenFolder As New Preferencia("Opcions PDF", "Obrir carpeta contenedora", "pdfOpenFolder", "Boolean", "Obrir carpeta contenedora al desar documents PDF al directori per defacte")
  Public pdfShowSaveDialog As New Preferencia("Opcions PDF", "Demanar on dessar el PDF generat", "pdfShowSaveDialog", "Boolean", "Desar documents PDF al directori per defacte")

  Public DirectoriArrelDocuments As New Preferencia("Ubicació fitxers", "Directori arrel documents aplicació", "Path_Documents", "FolderName", "Directori arrel on son els documents de la aplicació")
  'Public DirectoriDocumentsEscanejats As New Preferencia("Ubicació fitxers", "Directori Documents escanejats", "Path_Documents", "FolderName", "Directori on son els documents escanejats")
  'Public DirectoriFotos As New Preferencia("Ubicació fitxers", "Directori fotos", "Path_Fotos", "FolderName", "Directori on son les fots dels treballadors")
  'Public DirectoriFitxes As New Preferencia("Ubicació fitxers", "Directori fitxes productes", "Path_DocumentsProductes", "FolderName", "Directori on son les fots dels treballadors")
  Public DirectoriConfiguracio As New Preferencia("Ubicació fitxers", "Directori fitxer configuració", "Path_ConfigFiles", "FolderName", "Directori fitxers de configuració")
  'Public DirectoriFitxerSEUR As New Preferencia("Ubicació fitxers", "Directori fitxers SEUR", "Path_Remesa_SEUR", "FolderName", "Directori ubicació fitxer enviament remeses SEUR")
  Public DirectoriEdiInbox As New Preferencia("Ubicació fitxers", "Directori safata entrada EDI", "Path_EDI_Inbox", "FolderName", "Directori ubicació safata entrada EDI")
  Public DirectoriEdiOutbox As New Preferencia("Ubicació fitxers", "Directori safata sortida EDI", "Path_EDI_Outbox", "FolderName", "Directori ubicació safata sortida EDI")
  Public DirectoriCSB As New Preferencia("Ubicació fitxers", "Directori fitxers remeses bancàries", "Path_CSB", "FolderName", "Directori fitxers remeses bancàries")

  Public Help As New Preferencia("Ubicació fitxers", "Fitxer ajuda", "HelpFile", "FileName", "Fitxer d'ajuda de la aplicació")

  Public EmailServerSMTP As New Preferencia("e-mail", "Servidor SMTP", "SMTPServer", "String", "Servidor de correu sortint SMTP")
  Public EmailAdresa As New Preferencia("e-mail", "Adreça de correu pròpia", "SMTPe-mail", "String", "Adreça de correu electrònic")

  Public ParametresSerieBarCode As New Preferencia("Lector codi de barres", "Paràmetres comunicacions", "BarCodeReader", "SerialSettings", "Paràmetres de comunicació del lector de codi de barres")

  Public MostrarAvisosAlEntrar As New Preferencia("Preferències", "Mostrar avisos al inicar", "ShowAvisosOnStartUp", "Boolean", "Mostrar avisos al iniciar la aplicació")
  Public ImprimirPijamaLlistats As New Preferencia("Preferències", "Imprimir pijama als llistats", "Prf_LlistatPijama", "Boolean", "Imprimir ombra pijama als llistats.")

  Public SerieDocuments As New Preferencia("Numeració documents", "Serie documents", "Sys_SerieDocument", "String", "Serie per defecte a la numeració de documents")

  Public ComandesMagatzem As New Preferencia("Comandes", "Magatzem comandes", "CMD_Magatzem", "Integer", "Magatzem comandes")
  Public ComandesTpvZonaMagatzem As New Preferencia("Comandes", "Zona magatzem comandes", "CMD_TPV_ZonaMagatzem", "Integer", "Zona magatzem per a comandes TPV (Botiga, fires, autovenda)")
  Public ComandesIdiomaDescripcio As New Preferencia("Comandes", "Idioma descripció", "CMD_IDIOMA_IdiomaEnGrid", "Boolean", "Mostrar la descripció traduida al grid)")

  Public DesarPasswordUsuari As New Preferencia("Preferències", "Desar contrasenya usuari", "DesarPasswordUsuari", "Boolean", "Desar contrasenya usuari")

  Public StartUpDirectAppID As New Preferencia("Inici aplicació", "Identificador aplicació", "StartUpAppID", "String", "Identificador de menu de la aplicació a executar")
  Public StartUpDirectAppAcces As New Preferencia("Inici aplicació", "Nivell acces aplicació", "StartUpAppAcces", "enAccesLevel", "Nivell accés de la aplicació a executar")

  Public TerminalID As New Preferencia("Autovenda", "Codi terminal autovenda", "TerminalID", "Integer", "Codi del terminal autovenda")
  Public AgentID As New Preferencia("Autovenda", "Codi agent autovenda", "AgentID", "Integer", "Codi de l'agent d'autovenda")

  Public AVD_VisitaID As New Preferencia("Autovenda", "Número visita activa", "AVD_VisitaID", "Integer", "Número visita activa")
  Public AVD_SortidaID As New Preferencia("Autovenda", "Número sortida activa", "AVD_SortidaID", "Integer", "Número sortida activa")
  Public RunningOnTablet As New Preferencia("Autovenda", "Teminal treballa OFFLINE", "ModeTreballOFFLINE", "Boolean", "Indica si treballa en linea o desconectat de Mistral")
  Public UsuariTerminalServer As New Preferencia("Autovenda", "Teminal de Terminal Server", "UsuariTerminalServer", "Boolean", "Indica si treballa desde una sessió de terminal server")

  Public TerminalBotiga As New Preferencia("Botiga", "Terminal de botiga", "TerminalBotiga", "Boolean", "Indica si estem treballant en un terminal de botiga")

End Class

Public Class Preferencia
  Private mGrup As String
  Private mTitol As String
  Private mClauRegistre As String
  Private mTipusDada As String
  Private mDescripcio As String

  Public Sub New(ByVal Grup As String, ByVal Titol As String, ByVal ClauRegistre As String, ByVal TipusDada As String, ByVal Descripcio As String)
    mGrup = Grup
    mTitol = Titol
    mClauRegistre = ClauRegistre
    mTipusDada = TipusDada
    mDescripcio = Descripcio
  End Sub

  Public ReadOnly Property Grup() As String
    Get
      Return mGrup
    End Get
  End Property

  Public ReadOnly Property Titol() As String
    Get
      Return mTitol
    End Get
  End Property

  Public ReadOnly Property ClauRegistre() As String
    Get
      Return mClauRegistre
    End Get
  End Property

  Public ReadOnly Property TipusDada() As String
    Get
      Return mTipusDada
    End Get
  End Property

  Public ReadOnly Property Descripcio() As String
    Get
      Return mDescripcio
    End Get
  End Property

  Public Property Valor() As Object
    Get
      Dim value As Object
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        Select Case TipusDada
          Case "Integer", "enAccesLevel"
            value = Convert.ToInt32(RegKey.GetValue(ClauRegistre, 0))
          Case "Decimal"
            value = Convert.ToDecimal(RegKey.GetValue(ClauRegistre, 0))
          Case "DateTime"
            value = Convert.ToDateTime(RegKey.GetValue(ClauRegistre, Date.Today))
          Case "String", "Memo", "Char", "FileName", "FolderName", "PrinterName", "SerialSettings"
            value = Convert.ToString(RegKey.GetValue(ClauRegistre, ""))
          Case "Boolean"
            value = Convert.ToBoolean(RegKey.GetValue(ClauRegistre, False))
          Case "Color"
            value = System.Drawing.Color.FromArgb(CInt(RegKey.GetValue(ClauRegistre, -1)))
        End Select
      Catch ex As Exception
        value = Nothing
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As Object)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        Select Case TipusDada
          Case "Integer", "Decimal", "Memo", "Char", "String", "FileName", "FolderName", "PrinterName", "SerialSettings", "enAccesLevel"
            RegKey.SetValue(ClauRegistre, value)
          Case "DateTime"
            RegKey.SetValue(ClauRegistre, CStr(value))
          Case "Boolean"
            RegKey.SetValue(ClauRegistre, CStr(value))
          Case "Color"
            RegKey.SetValue(ClauRegistre, CType(value, System.Drawing.Color).ToArgb)
        End Select
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property


  Public Property ValorBoolean() As Boolean
    Get
      Dim value As Boolean
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        value = Convert.ToBoolean(RegKey.GetValue(ClauRegistre, False))
      Catch ex As Exception
        value = False
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As Boolean)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        RegKey.SetValue(ClauRegistre, CStr(value))
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property

  Public Property ValorString() As String
    Get
      Dim value As String
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        value = Convert.ToString(RegKey.GetValue(ClauRegistre, ""))
      Catch ex As Exception
        value = Nothing
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As String)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        RegKey.SetValue(ClauRegistre, value)
        
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property

  Public Property ValorInteger() As Integer
    Get
      Dim value As Integer
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        value = Convert.ToInt32(RegKey.GetValue(ClauRegistre, 0))
      Catch ex As Exception
        value = Nothing
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As Integer)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        RegKey.SetValue(ClauRegistre, value)
        
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property

  Public Property ValorDecimal() As Decimal
    Get
      Dim value As Decimal
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        value = Convert.ToDecimal(RegKey.GetValue(ClauRegistre, 0))
      Catch ex As Exception
        value = Nothing
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As Decimal)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        RegKey.SetValue(ClauRegistre, value)
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property

  Public Property ValorDateTime(ByVal DefaultDateTime As DateTime) As DateTime
    Get
      Dim value As DateTime
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        value = Convert.ToDateTime(RegKey.GetValue(ClauRegistre, Nothing))
      Catch ex As Exception
        value = Nothing
      Finally
        RegKey.Close()
      End Try
      Return value
    End Get

    Set(ByVal value As DateTime)
      Dim RegKey As Microsoft.Win32.RegistryKey
      RegKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppData.RegKey)
      Try
        RegKey.SetValue(ClauRegistre, CStr(value))
      Catch ex As Exception
      Finally
        RegKey.Close()
      End Try
    End Set
  End Property

End Class
