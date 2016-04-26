Imports System.Drawing
Imports System.Drawing.Printing
Imports Microsoft.Win32
Imports System.Net.Mail
Imports System.Data.SqlClient

Public MustInherit Class csRpt

#Region " Structures & Enums "

  ''' <summary>
  ''' Event per assignar un datareader a la propietat Datasource de la instancia csRpt
  ''' </summary>
  ''' <remarks></remarks>
  Public Event GetDataSource()
  Public Event FilterDatarowOut(ByVal Datarow As IDataReader, ByRef ExcludeRow As Boolean)

  ''' <summary>
  ''' Tipus de capçalera: Plain -> Texte; Image -> amb el logo de la empresa.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum HeaderKindEnum
    Plain
    Image
  End Enum

  ''' <summary>
  ''' Tipus numeració de pàgina. En el cas de PageNofM el llistat es repeteix 2 vegades
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum PageNumberEnum
    PageN
    PageNofM
  End Enum

  Public Enum ReportDestinationEnum
    Printer
    PDF
    FaxAsPDF
    eMailAsPDF
    Preview
    Excel
  End Enum

  Public Enum LayoutOffsetEnum
    Centered
    OneThird
    Custom
  End Enum

  Protected WithEvents pd As New Printing.PrintDocument

  Protected PrtSettings As New Printing.PrinterSettings
  Protected IsPrinterAssigned As Boolean

  Protected CurrentPage As Single
  Protected TotalPages As Single
  Protected DestinationDescription As String

  Protected CurX As Integer
  Protected CurY As Integer
  Protected BottomY As Integer
  Protected BodyLeft As Integer
  Protected BodyWidth As Integer

  Protected PageWidth As Integer
  Protected PageHeight As Integer

  Protected FirstPassReport As Boolean

  Protected LayoutInitialized As Boolean
  Protected EvaluatingTotalPages As Boolean

#End Region

  ''' <summary>Codi intern de la empresa.</summary>
  ''' <remarks></remarks>
  Public EmpresaID As Integer

  ''' <summary>Nom de la empresa, que aparaix a la capçalera del llistat.</summary>
  ''' <remarks></remarks>
  Public EmpresaName As String

  ''' <summary>Titol del llistat. Surt a la cua de impressió de windows. També s'utilitza com a nom de fitxer pdf.</summary>
  ''' <remarks></remarks>
  Public ReportID As String

  ''' <summary>Identificador del llistat que apareixera a cada llistat per identificarlo univocament</summary>
  ''' <remarks></remarks>
  Public ReportName As String

  ''' <summary>Identificador de l'usuari que genera el llistat</summary>
  ''' <remarks></remarks>
  Public UserName As String

  ''' <summary>Nom complet de l'usuari que genera el llistat.</summary>
  ''' <remarks></remarks>
  Public UserFullName As String

  ''' <summary>Adreça de correu de l'usuari que genera el llistat.</summary>
  ''' <remarks></remarks>
  Public UserEMail As String

  ''' <summary>Data del moment en que s'inicia el llistat.</summary>
  ''' <remarks></remarks>
  Public ReportDate As Date

  ''' <summary>Imatge del logo de la empresa per a HeaderKind = HeaderKindEnum.Image</summary>
  ''' <remarks></remarks>
  Public CustomerLogo As Image

  ''' <summary>Si es True el llistat es genera amb el full apaisat.</summary>
  ''' <remarks></remarks>
  Protected mLandscape As Boolean

  ''' <summary>IP del lloc de treball que geera el llistat</summary>
  ''' <remarks></remarks>
  Public WorkstationIP As String

  ''' <summary>Nom del llistat que s'està imprimint</summary>
  ''' <remarks></remarks>
  Public DataSource As IDataReader

  ''' <summary>Disposició horitzontal del llistat.</summary>
  ''' <remarks></remarks>
  Public LayoutOffset As LayoutOffsetEnum

  ''' <summary>Marge esquerre del llistat en 1/100 de polsada.</summary>
  ''' <remarks></remarks>
  Public LeftOffset As Integer

  ''' <summary>Tipus de capçalera del llistat.</summary>
  ''' <remarks></remarks>
  Public HeaderKind As HeaderKindEnum

  ''' <summary>Tipus de numeració de pàgina.</summary>
  ''' <remarks> En el cas de PageNofM el llistat es repeteix 2 vegades</remarks>
  Public PageNumbering As PageNumberEnum

  ''' <summary>Destinació del llistat. Impressiora, fax etc...</summary>
  ''' <remarks></remarks>
  Public Destination As ReportDestinationEnum

  ''' <summary>Nom del fitxer PDF</summary>
  ''' <remarks></remarks>
  Public fmContactes As DataTable

  ''' <summary>Imatge del logo 'fax' empleat al fax</summary>
  ''' <remarks></remarks>
  '''
  Public fmFaxLogo As Image

  ''' <summary>Imatge del logo de la empresa a utilitzar al fax. Normalment en gama de grisos o en BN</summary>
  ''' <remarks></remarks>
  Public fmFaxCustomerLogo As Image

  ''' <summary>Imatge del logo 'fax' empleat al fax</summary>
  ''' <remarks></remarks>
  '''
  Public fmFaxLogoFile As String

  ''' <summary>Imatge del logo de la empresa a utilitzar al fax. Normalment en gama de grisos o en BN</summary>
  ''' <remarks></remarks>
  Public fmFaxCustomerLogoFile As String

  ''' <summary>Nom del Usuari que envia el fax</summary>
  ''' <remarks></remarks>
  Public fmFaxNomUsuari As String

  ''' <summary>NÚMERO DE FAX ON S'ENVIA</summary>
  ''' <remarks></remarks>
  Public fmFaxNumero As String

  ''' <summary>Nom del destinatari</summary>
  ''' <remarks></remarks>
  Public fmFaxAlaAtencio As String

  ''' <summary>Adresa de correu des d'on s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmMailFrom As String

  ''' <summary>Adresa de correu o números de fax on s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmMailTo As String

  ''' <summary>Adresa de correu o números de fax on s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmMailReplyTo As String

  ''' <summary>Adresa de correu CC a qui s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmMailCC As String

  ''' <summary>Mail on s'envia una còpia a manera de bústia d'enviats</summary>
  ''' <remarks></remarks>
  Public fmMailFeedBack As String

  ''' <summary>e-mail del compte del servei de fax</summary>
  ''' <remarks></remarks>
  Public fmMailAccountFax As String

  ''' <summary>Resultat enviament mail</summary>
  ''' <remarks></remarks>
  Public fmMailSentOK As Boolean

  ''' <summary>Subject del correu amb el que s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmSubject As String

  ''' <summary>Body del correu amb el que s'envia el llistat</summary>
  ''' <remarks></remarks>
  Public fmBody As String

  ''' <summary>Servidor SMTP per enviar el correu electrònic</summary>
  ''' <remarks></remarks>
  Public fmSmtpServer As String

  ''' <summary>Login del servidor SMTP per enviar el correu electrònic</summary>
  ''' <remarks></remarks>
  Public fmSmtpLogin As String

  ''' <summary>Password del servidor SMTP per enviar el correu electrònic</summary>
  ''' <remarks></remarks>
  Public fmSmtpPassword As String

  Public fmOFDFileName As String
  Public fmOFDDefaultExension As String
  Public fmOFDInitialDir As String
  Public fmOFDFilter As String
  Public fmOFDShowDialog As Boolean
  Public fmCanAttachFiles As Boolean
  Public fmShowForm As Boolean

  Public pdfShowSaveDialog As Boolean
  Public pdfDirectori As String
  Public pdfNomFitxer As String
  Public pdfPathAndFileName As String
  Public pdfOpenFolder As Boolean

  Private mPdfNumberOfJobs As Integer
  Private pdfMustInitNumberOfJobs As Boolean
  Public Property pdfNumberOfJobs() As Integer
    Get
      Return mPdfNumberOfJobs
    End Get
    Set(ByVal value As Integer)
      mPdfNumberOfJobs = value
      pdfMustInitNumberOfJobs = True
    End Set
  End Property


  ''' <summary>Nom del formulari des d'on es genera el report</summary>
  ''' <remarks></remarks>
  Public Formulari As String

  Protected DataNeeded As Boolean
  Protected DrawingTotalsAndExit As Boolean
  Protected LoadDataSource As Boolean
  Private mShowPrintDialog As Boolean
  Private mShowPDFGenerated As Boolean
  Private mShowMessageError As Boolean
  Protected mStopReport As Boolean = False

  ''' <summary>Nom de la impresora per imprimir el llistat</summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ''' 

  Public Property PrinterName() As String
    Get
      Return PrtSettings.PrinterName
    End Get
    Set(ByVal value As String)
      SetPrinter(value)
    End Set
  End Property

  Public Property ShowPDFGenerated() As Boolean
    Get
      Return mShowPDFGenerated
    End Get
    Set(ByVal value As Boolean)
      mShowPDFGenerated = value
    End Set
  End Property

  ''' <summary>
  ''' Mostrar sempre el quadre de dialeg de sel.lecció d'impresora 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property ShowPrintDialog() As Boolean
    Get
      Return mShowPrintDialog
    End Get
    Set(ByVal value As Boolean)
      mShowPrintDialog = value
    End Set
  End Property

  Public Property ShowMessageError() As Boolean
    Get
      Return mShowMessageError
    End Get
    Set(ByVal value As Boolean)
      mShowMessageError = value
    End Set
  End Property

  Protected mDestinationStr As String
  ''' <summary>
  ''' Destinació del llistat. Es la inicial de destinació: I, E, F, P, V, X
  ''' </summary>
  ''' <value>Clau de destinació de la impressió.</value>
  ''' <remarks></remarks>
  Public WriteOnly Property DestinationStr() As String
    Set(ByVal value As String)
      value = Utils.CNull(value, "I")
      Select Case value.ToUpper '.Replace("&", "").ToUpper.Substring(0, 1)
        Case "E-MAIL", "E"
          Destination = ReportDestinationEnum.eMailAsPDF
          mDestinationStr = value
        Case "FAX", "F"
          Destination = ReportDestinationEnum.FaxAsPDF
          mDestinationStr = value
        Case "ADOBE PDF", "P"
          Destination = ReportDestinationEnum.PDF
          mDestinationStr = value
        Case "VISTA PRÈVIA", "V"
          Destination = ReportDestinationEnum.Preview
          mDestinationStr = value
        Case "IMPRIMIR", "I"
          Destination = ReportDestinationEnum.Printer
          mDestinationStr = value
        Case "EXCEL", "X"
          Destination = ReportDestinationEnum.Excel
          mDestinationStr = value
      End Select
    End Set
  End Property

  ''' <summary>
  ''' Si es True, imprimeix el llistat en apaisat.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Landscape() As Boolean
    Get
      Return mLandscape
    End Get
    Set(ByVal value As Boolean)
      mLandscape = value
      PrtSettings.DefaultPageSettings.Landscape = value
    End Set
  End Property

  Private Sub pd_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles pd.BeginPrint
    TotalPages = 0
    BeginPrint()
  End Sub

  Private Sub pd_EndPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles pd.EndPrint
    If DataSource IsNot Nothing Then
      DataSource.Close()
    End If
  End Sub

  Private Sub pd_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles pd.PrintPage

    If EvaluatingTotalPages Then
      If LoadDataSource Then
        RaiseEvent GetDataSource()
        TotalPages += 1
        FirstPassReport = True
        Do While DrawPage(e.Graphics)
          TotalPages += 1
          ' e.Graphics.Clear(Color.White)
        Loop
        EvaluatingTotalPages = False
        FirstPassReport = True
        If DataSource IsNot Nothing Then
          DataSource.Close()
        End If
        ' e.Graphics.Clear(Color.White)
        BeginPrint()
        'e.HasMorePages = False
        'Return
      End If
    End If

    If FirstPassReport Then
      RaiseEvent GetDataSource()
      LoadDataSource = False
    End If

    e.HasMorePages = DrawPage(e.Graphics)

  End Sub

  ''' <summary>
  ''' Demana la impressora per la que es vol enviar el llistat en cas de ReportDestinationEnum.Printer
  ''' </summary>
  ''' <returns>True si l'usuari ha selecionat una impressora. False en cas contrari</returns>
  ''' <remarks></remarks>
  Private Function GetPrinter() As Boolean
    Dim ReturnValue As Boolean = False
    If Me.Destination = ReportDestinationEnum.eMailAsPDF Or _
      Me.Destination = ReportDestinationEnum.FaxAsPDF Or _
      Me.Destination = ReportDestinationEnum.PDF Or _
      Me.Destination = ReportDestinationEnum.Preview Or _
      Me.Destination = ReportDestinationEnum.Excel Then
      Return True
    End If
    Dim dlg As New Windows.Forms.PrintDialog
    dlg.PrinterSettings = PrtSettings
    If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
      pd.PrinterSettings = dlg.PrinterSettings
      pd.DefaultPageSettings.PrinterSettings = pd.PrinterSettings
      PrtSettings = dlg.PrinterSettings
      ShowPrintDialog = False
      IsPrinterAssigned = True
      ReturnValue = True
    Else
      Return False
    End If
    Return ReturnValue
  End Function

  ''' <summary>
  ''' Assigna la impressora per la que es vol enviar el llistat
  ''' </summary>
  ''' <param name="PrinterName">Nom de la impressora sobre la que es vol imprimir el llistat</param>
  ''' <returns>True si s'aplicat correctament el nom de la impressora. False si l'usuari cancela la sel·lecció.</returns>
  ''' <remarks></remarks>
  Public Function SetPrinter(ByVal PrinterName As String) As Boolean
    If csUtils.Utils.IsEmptyStr(PrinterName) Then
      Return GetPrinter()
    Else
      PrtSettings.PrinterName = PrinterName
      pd.PrinterSettings = PrtSettings
      IsPrinterAssigned = True
    End If
    Return True
  End Function

  Public Function SetDefaultPrinter() As Boolean
    Dim oPS As New System.Drawing.Printing.PrinterSettings
    Try
      PrtSettings.PrinterName = oPS.PrinterName
      pd.PrinterSettings = PrtSettings
      IsPrinterAssigned = True
    Catch ex As Exception
      IsPrinterAssigned = False
    End Try
    Return True
  End Function

  ''' <summary>
  ''' Retorna el total de pagines impresses al finalitzar el llistat
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetTotalPages() As Integer
    Return CInt(TotalPages)
  End Function

  ''' <summary>
  ''' Retorna el nom de la impressora o el nom del fitxer on s'ha generat el llistat.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDestination() As String
    Return DestinationDescription
  End Function

  ''' <summary>
  ''' Inicia el llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Print()
    Dim sfDlg As Windows.Forms.SaveFileDialog

    If Me.Destination = ReportDestinationEnum.Excel Then
      sfDlg = New Windows.Forms.SaveFileDialog
      'sfDlg.DefaultExt = "XLS"
      sfDlg.Filter = "Fitxers Excel (*.xls)|*.xls|Tots els fitxers (*.*)|*.*"
      sfDlg.FilterIndex = 1
      sfDlg.RestoreDirectory = True
      sfDlg.FileName = ReportName + ".xls"

      If sfDlg.ShowDialog() = Windows.Forms.DialogResult.OK Then
        RaiseEvent GetDataSource()
        Print2Excel(sfDlg.FileName)
      End If

      Return

    End If

    If Me.Destination = ReportDestinationEnum.Printer Then
      If Not IsPrinterAssigned Then
        If ShowPrintDialog Then
          If Not GetPrinter() Then
            Return
          End If
        Else
          If Not SetDefaultPrinter() Then
            Return
          End If
        End If
      Else
        If ShowPrintDialog Then
          If Not GetPrinter() Then
            Return
          End If
        End If
      End If
    End If

    pd.PrinterSettings = PrtSettings
    pd.DefaultPageSettings.Margins = New System.Drawing.Printing.Margins(0, 0, 0, 0)
    pd.DocumentName = ReportName

    If Me.PageNumbering = PageNumberEnum.PageNofM Then
      Me.EvaluatingTotalPages = True
    End If

    Select Case Me.Destination
      Case ReportDestinationEnum.Printer
        DestinationDescription = pd.PrinterSettings.PrinterName
        pd.Print()
        'END Case ReportDestinationEnum.Printer

      Case ReportDestinationEnum.Preview
        DestinationDescription = "Vista prèvia"
        Dim ppdlg As New System.Windows.Forms.PrintPreviewDialog

        ppdlg.Document = pd
        ppdlg.ShowDialog()
        'End Case ReportDestinationEnum.Preview

      Case ReportDestinationEnum.PDF
        ' Dim TmpPageNumbering As PageNumberEnum = Me.PageNumbering
        ' Me.PageNumbering = PageNumberEnum.PageN
        Me.PrintPDF()

        'recuperem l'estat de pagenumbering
        ' Me.PageNumbering = TmpPageNumbering

      Case ReportDestinationEnum.eMailAsPDF
        'Dim TmpPageNumbering As PageNumberEnum = Me.PageNumbering
        'Me.PageNumbering = PageNumberEnum.PageN

        pdfShowSaveDialog = False

        If Me.PrintPDF() Then
          Dim sm As New csFaxMail

          sm.fmDestination = FaxEmailFormActorEmum.actorEmail
          sm.fmContactes = fmContactes

          sm.fmSubject = fmSubject
          sm.fmBody = fmBody

          sm.fmShowForm = fmShowForm
          sm.fmAttachment = pdfPathAndFileName
          sm.fmCanAttachFiles = fmCanAttachFiles

          sm.fmOFDDefaultExtension = fmOFDDefaultExension
          sm.fmOFDFilter = fmOFDFilter
          sm.fmOFDInitialDir = fmOFDInitialDir

          sm.fmMailFeedBack = fmMailFeedBack
          sm.fmMailFrom = fmMailFrom
          sm.fmMailTo = fmMailTo
          sm.fmMailReplyTo = fmMailReplyTo

          sm.fmSmtpLogin = fmSmtpLogin
          sm.fmSmtpPassword = fmSmtpPassword
          sm.fmSmtpServer = fmSmtpServer

          sm.fmDeleteSentFiles = True

          sm.Send()

          fmMailSentOK = sm.fmMailSentOK

          sm = Nothing

        End If


      Case ReportDestinationEnum.FaxAsPDF

        pdfShowSaveDialog = False

        If Me.PrintPDF() Then

          Dim sm As New csFaxMail

          sm.fmDestination = FaxEmailFormActorEmum.actorFax
          sm.fmContactes = fmContactes

          sm.fmFaxNumero = fmFaxNumero
          sm.fmFaxNomUsuari = fmFaxNomUsuari
          sm.fmFaxAlaAtencio = fmFaxAlaAtencio
          sm.fmMailAccountFax = fmMailAccountFax

          sm.fmFaxLogo = fmFaxLogo
          sm.fmFaxCustomerLogo = fmFaxCustomerLogo
          sm.fmFaxLogoFile = fmFaxLogoFile
          sm.fmFaxCustomerLogoFile = fmFaxCustomerLogoFile

          sm.fmSubject = fmSubject
          sm.fmBody = fmBody

          sm.fmFaxPaginesDocument = CInt(TotalPages)

          sm.fmAttachment = pdfPathAndFileName
          sm.fmCanAttachFiles = fmCanAttachFiles
          sm.fmOFDDefaultExtension = fmOFDDefaultExension
          sm.fmOFDFilter = fmOFDFilter
          sm.fmOFDInitialDir = fmOFDInitialDir

          sm.fmMailFeedBack = fmMailFeedBack
          sm.fmMailFrom = fmMailFrom
          sm.fmMailReplyTo = fmMailReplyTo

          sm.fmSmtpLogin = fmSmtpLogin
          sm.fmSmtpPassword = fmSmtpPassword
          sm.fmSmtpServer = fmSmtpServer

          sm.fmShowForm = True
          sm.fmDeleteSentFiles = True

          sm.Send()

          fmMailSentOK = sm.fmMailSentOK

          sm = Nothing

        End If


    End Select

  End Sub

  Public Sub StopReport()
    mStopReport = True
  End Sub

  Private Function PrintPDF() As Boolean
    Try
      Dim reg As RegistryKey
      Dim reg2 As RegistryKey
      Dim OldPrinterName As String = pd.PrinterSettings.PrinterName
      Dim dlg As New System.Windows.Forms.SaveFileDialog
      Dim PrintTheDocument As Boolean = True
      Dim CurrentJob As Integer

      If Not String.IsNullOrEmpty(pdfPathAndFileName) Then
        ' ens pasan el nom sencer
        pdfDirectori = IO.Path.GetDirectoryName(pdfPathAndFileName)
        pdfNomFitxer = IO.Path.GetFileName(pdfPathAndFileName)
      End If

      If String.IsNullOrEmpty(pdfDirectori) Then
        If pdfShowSaveDialog Then
          pdfDirectori = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Else
          pdfDirectori = IO.Path.GetTempPath
        End If
      End If

      If String.IsNullOrEmpty(pdfNomFitxer) Or pdfShowSaveDialog Then

        dlg.CheckPathExists = True
        dlg.DefaultExt = "pdf"
        dlg.Filter = "Fitxers pdf (*.pdf)|*.pdf"
        dlg.FileName = pdfNomFitxer
        If String.IsNullOrEmpty(fmOFDInitialDir) Then
          dlg.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        Else
          dlg.InitialDirectory = fmOFDInitialDir
        End If
        dlg.RestoreDirectory = True

        If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
          pdfPathAndFileName = dlg.FileName
        Else
          PrintTheDocument = False
        End If

      Else

        pdfPathAndFileName = IO.Path.Combine(pdfDirectori, pdfNomFitxer)

      End If

      pdfDirectori = IO.Path.GetDirectoryName(pdfPathAndFileName)
      pdfNomFitxer = IO.Path.GetFileName(pdfPathAndFileName)

      pd.PrinterSettings.PrinterName = "pdfFactory"

      If PrintTheDocument Then

        'configurem el registre per poder imprimir fitxer en PDF

        reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory6\FinePrinters\pdfFactory", True)
        If reg Is Nothing Then
          reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory5\FinePrinters\pdfFactory", True)
          If reg Is Nothing Then
            reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory4\FinePrinters\pdfFactory", True)
            If reg Is Nothing Then
              reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3\FinePrinters\pdfFactory", True)
              reg2 = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3", True)
            Else
              reg2 = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory4", True)
            End If
          Else
            reg2 = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory5", True)
          End If
        Else
          reg2 = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory6", True)
        End If

        reg.SetValue("ShowDlg", 2)
        If mShowPDFGenerated Then
          reg.SetValue("PdfAction", 1)
        Else
          reg.SetValue("PdfAction", 0)
        End If
        reg2.SetValue("OutputFile", pdfPathAndFileName, RegistryValueKind.String)

        If pdfMustInitNumberOfJobs Then
          pdfMustInitNumberOfJobs = False
          reg.SetValue("CollectJobs", pdfNumberOfJobs)
        End If

        CurrentJob = Utils.CNull(reg.GetValue("CollectJobs", 0), 0)

        pd.Print()

        ' no continuarem fins que la impressió no alliberi el fitxer

        If CurrentJob > 0 Then
          Do While Utils.CNull(reg.GetValue("CollectJobs", 0), 0) = CurrentJob
            Threading.Thread.Sleep(500)
          Loop
        Else
          Do While reg2.GetValue("OutputFile") IsNot Nothing
            Threading.Thread.Sleep(500)
          Loop
        End If

        'Reiniciem el registre
        reg.SetValue("ShowDlg", 1)
        reg.SetValue("PdfAction", 0)
        reg2.DeleteValue("OutputFilePerm", False)
        reg.Close()
        reg2.Close()
      End If

      If pdfOpenFolder Then
        Process.Start("explorer.exe", pdfDirectori)
      End If

    Catch ex As Exception
      If ShowMessageError Then
        MsgBox("No s'ha pogut generar el fitxer PDF.", MsgBoxStyle.Exclamation, "ERROR")
      End If
      Return False
    End Try

    Return True

  End Function

  Public MustOverride Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean

  Public MustOverride Sub BeginPrint()

  Protected Overridable Function DrawHeader(ByVal Canvas As System.Drawing.Graphics) As Integer

  End Function

  Protected Overridable Function DrawFooter(ByVal Canvas As System.Drawing.Graphics) As Integer

  End Function

  Protected Overridable Function FilterRowOut(ByVal Row As IDataReader) As Boolean
    Dim ExcludeRow As Boolean
    ExcludeRow = False
    RaiseEvent FilterDatarowOut(Row, ExcludeRow)
    Return ExcludeRow
  End Function

  Public Function PrintStandardHeader(ByVal Canvas As Graphics) As Integer
    Dim sf As New StringFormat
    Dim PenThick As New Pen(Color.Black, 2)
    Dim PenThin As New Pen(Color.Black, 1)
    Dim HeaderFont As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point)
    Dim HeaderBrush As Brush = Brushes.Black

    PageHeight = CInt(Canvas.VisibleClipBounds.Height)
    PageWidth = CInt(Canvas.VisibleClipBounds.Width)

    sf.Alignment = StringAlignment.Far
    Me.CurY = 5

    CurrentPage += 1

    Me.DrawLine(Canvas, PenThick, 0, CurY, PageWidth, CurY)

    Me.CurY += 2
    Me.DrawString(Canvas, EmpresaName, HeaderFont, HeaderBrush, 0, Me.CurY)
    Me.DrawString(Canvas, String.Format(ReportName, CurrentPage), HeaderFont, HeaderBrush, PageWidth, CurY, sf)
    Me.CurY += CInt(HeaderFont.GetHeight(Canvas))
    Me.DrawLine(Canvas, PenThin, 0, CurY, PageWidth, CurY)
    Me.CurY += 20

    HeaderFont.Dispose()
    PenThick.Dispose()
    PenThin.Dispose()

    Return Me.CurY
  End Function

  Public Function PrintStandardFooter(ByVal Canvas As Graphics) As Integer
    Dim PenThick As New Pen(Color.Black, 2)
    Dim PenThin As New Pen(Color.Black, 1)
    Dim FooterFont As New Font("Arial", 6, FontStyle.Regular, GraphicsUnit.Point)
    Dim FooterBrush As Brush = Brushes.Black

    Dim y As Integer
    Dim sf As New StringFormat

    PageHeight = CInt(Canvas.VisibleClipBounds.Height)
    PageWidth = CInt(Canvas.VisibleClipBounds.Width)

    y = CInt(PageHeight - FooterFont.GetHeight(Canvas)) - 2
    BottomY = y - 7

    Me.DrawLine(Canvas, PenThin, 0, y, PageWidth, y)
    y += 1
    Me.DrawString(Canvas, String.Format("Usuari: {0} LT: {1} FM: {2}", Me.UserName, Me.WorkstationIP, Me.ReportID), FooterFont, FooterBrush, 0, y)
    sf.Alignment = StringAlignment.Center
    Me.DrawString(Canvas, String.Format("Data: {0:dd/MM/yyyy HH:mm}", Date.Now), FooterFont, FooterBrush, PageWidth \ 2, y, sf)
    sf.Alignment = StringAlignment.Far
    If PageNumbering = PageNumberEnum.PageN Then
      Me.DrawString(Canvas, String.Format("Pàgina: {0}", CurrentPage), FooterFont, FooterBrush, PageWidth, y, sf)
    Else
      Me.DrawString(Canvas, String.Format("Pàgina: {0} de {1}", CurrentPage, TotalPages), FooterFont, FooterBrush, PageWidth, y, sf)
    End If

    FooterFont.Dispose()
    PenThick.Dispose()
    PenThin.Dispose()

    Return BottomY
  End Function

  Protected MustOverride Sub Print2Excel(ByVal FileName As String)


#Region "Primitives"

  Public Sub DrawString(ByVal gr As Graphics, ByVal s As String, ByVal f As Font, ByVal b As Brush, ByVal x As Single, ByVal y As Single)
    If Not EvaluatingTotalPages Then
      gr.DrawString(s, f, b, x, y)
    End If
  End Sub

  Public Sub DrawString(ByVal gr As Graphics, ByVal s As String, ByVal f As Font, ByVal b As Brush, ByVal x As Single, ByVal y As Single, ByVal sf As StringFormat)
    If Not EvaluatingTotalPages Then
      gr.DrawString(s, f, b, x, y, sf)
    End If
  End Sub

  Public Sub DrawString(ByVal gr As Graphics, ByVal s As String, ByVal f As Font, ByVal b As Brush, ByVal r As RectangleF, ByVal sf As StringFormat)
    If Not EvaluatingTotalPages Then
      gr.DrawString(s, f, b, r, sf)
    End If
  End Sub

  Public Sub DrawLine(ByVal gr As Graphics, ByVal p As Pen, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single)
    If Not EvaluatingTotalPages Then
      gr.DrawLine(p, x1, y1, x2, y2)
    End If
  End Sub

  Public Sub DrawLine(ByVal gr As Graphics, ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, stroke As Single, clr As Color)
    If Not EvaluatingTotalPages Then
      gr.DrawLine(New Pen(clr, stroke), x1, y1, x2, y2)
    End If
  End Sub

  Public Sub DrawDashedLine(ByVal gr As Graphics, _
    ByVal x As Integer, _
    ByVal y As Integer, _
    ByVal width As Integer, _
    ByVal p As Pen, _
    Optional ByVal length As Integer = 10, _
    Optional ByVal gap As Integer = 2)

    Dim LineFrom As Integer

    LineFrom = x

    While LineFrom < x + width
      DrawLine(gr, p, LineFrom, y, Math.Min(LineFrom + length, x + width), y)
      LineFrom += length + gap
    End While

  End Sub

  Public Sub DrawLines(ByVal gr As Graphics, ByVal p As Pen, ByVal points() As Point)
    If Not EvaluatingTotalPages Then
      gr.DrawLines(p, points)
    End If
  End Sub

  Public Sub DrawLines(ByVal gr As Graphics, ByVal p As Pen, ByVal points() As PointF)
    If Not EvaluatingTotalPages Then
      gr.DrawLines(p, points)
    End If
  End Sub

  Public Sub DrawRectangle(ByVal gr As Graphics, ByVal p As Pen, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single)
    If Not EvaluatingTotalPages Then
      gr.DrawRectangle(p, x, y, width, height)
    End If
  End Sub

  Public Sub DrawImage(ByVal gr As Graphics, ByVal i As Image, ByVal desRect As Rectangle, ByVal srcRec As Rectangle, ByVal u As GraphicsUnit)
    If Not EvaluatingTotalPages Then
      gr.DrawImage(i, desRect, srcRec, u)
    End If
  End Sub

  Public Sub DrawImage(ByVal gr As Graphics, ByVal i As Image, ByVal desRect As RectangleF, ByVal srcRec As RectangleF, ByVal u As GraphicsUnit)
    If Not EvaluatingTotalPages Then
      gr.DrawImage(i, desRect, srcRec, u)
    End If
  End Sub

  Public Sub DrawImage(ByVal gr As Graphics, ByVal i As Image, ByVal x As Integer, ByVal y As Integer)
    If Not EvaluatingTotalPages Then
      gr.DrawImage(i, x, y)
    End If
  End Sub

  Public Sub DrawImage(ByVal gr As Graphics, ByVal i As Image, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
    If Not EvaluatingTotalPages Then
      gr.DrawImage(i, x, y, width, height)
    End If
  End Sub

  Public Sub FillRectangle(ByVal gr As Graphics, ByVal b As Brush, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single)
    If Not EvaluatingTotalPages Then
      gr.FillRectangle(b, x, y, width, height)
    End If
  End Sub

  Public Sub DrawRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Diameter As Integer, ByVal Stroke As Integer)
    If Not EvaluatingTotalPages Then
      Utils.DrawRoundedRectangle(objGraphics, x, y, Width, Height, Diameter, Stroke)
    End If
  End Sub

  Public Sub DrawRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Diameter As Integer, ByVal Stroke As Integer, PenColor As Color)
    If Not EvaluatingTotalPages Then
      Utils.DrawRoundedRectangle(objGraphics, x, y, Width, Height, Diameter, Stroke, PenColor)
    End If
  End Sub

  Public Sub FillRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics, ByVal brush As System.Drawing.Brush, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Diameter As Integer)
    If Not EvaluatingTotalPages Then
      Utils.FillRoundedRectangle(objGraphics, brush, x, y, Width, Height, Diameter)
    End If
  End Sub

  Public Sub FillTopRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics, ByVal brush As System.Drawing.Brush, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Diameter As Integer)
    If Not EvaluatingTotalPages Then
      Utils.FillTopRoundedRectangle(objGraphics, brush, x, y, Width, Height, Diameter)
    End If
  End Sub

  Public Sub DrawRotateText(ByVal objGraphics As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal Angle As Integer, ByVal Text As String, ByVal Fnt As System.Drawing.Font, ByVal brsh As System.Drawing.Brush)
    If Not EvaluatingTotalPages Then
      Utils.DrawRotateText(objGraphics, x, y, Angle, Text, Fnt, brsh)
    End If
  End Sub

  Public Sub DrawRotateImage(ByVal gr As System.Drawing.Graphics, ByVal bmp As System.Drawing.Bitmap, ByVal x As Integer, ByVal y As Integer, ByVal angle As Single)
    If Not EvaluatingTotalPages Then
      Utils.DrawRotateImage(gr, bmp, x, y, angle)
    End If
  End Sub

  Public Sub DrawRotateImage(ByVal objGraphics As System.Drawing.Graphics, ByVal Image As System.Drawing.Image, ByVal Angle As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer)
    If Not EvaluatingTotalPages Then
      Utils.DrawRotateImage(objGraphics, Image, Angle, x, y, width, height)
    End If
  End Sub

  Public Sub DrawFittedText(ByVal objGraphics As System.Drawing.Graphics, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal Text As String, ByVal Fnt As System.Drawing.Font, ByVal foreBrush As System.Drawing.Brush, ByVal backBrush As System.Drawing.Brush, ByVal Stroke As Integer, ByVal Diameter As Integer)
    If Not EvaluatingTotalPages Then
      Utils.DrawFittedText(objGraphics, x, y, width, height, Text, Fnt, foreBrush, backBrush, Stroke, Diameter)
    End If
  End Sub

  Public Sub DrawBoxedText(ByVal Canvas As Graphics, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal s As String, ByVal fnt As Font, ByVal Aligment As StringFormat, ByVal BackColor As Brush, ByVal ForeColor As Brush)
    If Not EvaluatingTotalPages Then
      FillRectangle(Canvas, BackColor, x, y, width, height)
      DrawRectangle(Canvas, Pens.Black, x, y, width, height)
      DrawString(Canvas, s, fnt, ForeColor, New RectangleF(x, y, width, height), Aligment)
    End If
  End Sub

#End Region

End Class

Public Class csPortadaFax
  Inherits csRpt
  Private WithEvents pdPortadaFax As New Printing.PrintDocument

  Public fmFaxPaginesDocument As Integer

  Private Sub pdPortadaFax_BeginPrint(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs) Handles pdPortadaFax.BeginPrint

  End Sub

  Private Sub pdPortadaFax_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles pdPortadaFax.PrintPage

    Dim pageHeight As Integer
    Dim pageWidth As Integer

    pageHeight = CInt(e.Graphics.VisibleClipBounds.Height)
    pageWidth = CInt(e.Graphics.VisibleClipBounds.Width)

    If Not IsNothing(fmFaxCustomerLogo) Then
      Me.DrawImage(e.Graphics, fmFaxCustomerLogo, 30, 30)

    End If
    If Not IsNothing(fmFaxLogo) Then
      Me.DrawImage(e.Graphics, fmFaxLogo, pageWidth - fmFaxLogo.Width - 30, 30, fmFaxLogo.Width, fmFaxLogo.Height)
      Me.DrawImage(e.Graphics, fmFaxLogo, 30, pageHeight - fmFaxLogo.Height - 30)
    End If

    Dim r1, r2, r3, r4, r5 As Integer

    CurY = 275

    ' Layout

    Using fntCaption As New Font("Arial", 10, FontStyle.Bold)

      Dim rh As Integer = CInt(fntCaption.GetHeight(e.Graphics))

      r1 = CurY
      Me.DrawString(e.Graphics, "Para / Pour:", fntCaption, Brushes.Black, 30, CurY)
      Me.DrawString(e.Graphics, "De:", fntCaption, Brushes.Black, 480, CurY)
      Me.DrawLine(e.Graphics, Pens.Black, 30, CurY + rh + 2, pageWidth - 30, CurY + rh + 2)
      CurY += 50

      r2 = CurY
      Me.DrawString(e.Graphics, "Fax:", fntCaption, Brushes.Black, 30, CurY)
      Me.DrawString(e.Graphics, "Páginas / Pages:", fntCaption, Brushes.Black, 480, CurY)
      Me.DrawLine(e.Graphics, Pens.Black, 30, CurY + rh + 2, pageWidth - 30, CurY + rh + 2)
      CurY += 50

      r3 = CurY
      Me.DrawString(e.Graphics, "Tel / Tél:", fntCaption, Brushes.Black, 30, CurY)
      Me.DrawString(e.Graphics, "Fecha / Date:", fntCaption, Brushes.Black, 480, CurY)
      Me.DrawLine(e.Graphics, Pens.Black, 30, CurY + rh + 2, pageWidth - 30, CurY + rh + 2)
      CurY += 50

      r4 = CurY
      Me.DrawString(e.Graphics, "Asunto / Sujet:", fntCaption, Brushes.Black, 30, CurY)
      Me.DrawLine(e.Graphics, Pens.Black, 30, CurY + rh + 2, pageWidth - 30, CurY + rh + 2)
      CurY += 70

      r5 = CurY
      Me.DrawString(e.Graphics, "Comentario / Comentaire:", fntCaption, Brushes.Black, 30, CurY)
      Me.DrawLine(e.Graphics, Pens.Black, 230, CurY + rh \ 2, pageWidth - 30, CurY + rh \ 2)
      Me.DrawLine(e.Graphics, Pens.Black, pageWidth - 30, CurY + rh \ 2, pageWidth - 30, pageHeight - 50)
      Me.DrawLine(e.Graphics, Pens.Black, 190, pageHeight - 50, pageWidth - 30, pageHeight - 50)

      'Data

    End Using

    Using fntText As New Font("Arial", 10, FontStyle.Regular)

      Dim rh As Integer = CInt(fntText.GetHeight(e.Graphics))

      Dim sf As New StringFormat
      sf.Alignment = StringAlignment.Near
      sf.Trimming = StringTrimming.EllipsisCharacter
      Dim rl As Rectangle

      rl = New Rectangle(120, r1, 300, rh)
      Me.DrawString(e.Graphics, fmFaxAlaAtencio, fntText, Brushes.Black, rl, sf)
      rl = New Rectangle(530, r1, 220, rh)
      Me.DrawString(e.Graphics, fmFaxNomUsuari, fntText, Brushes.Black, rl, sf)

      rl = New Rectangle(120, r2, 300, rh)
      Me.DrawString(e.Graphics, fmFaxNumero, fntText, Brushes.Black, rl, sf)
      rl = New Rectangle(600, r2, 200, rh)
      Me.DrawString(e.Graphics, CStr(fmFaxPaginesDocument + 1).ToString + " (inc.p.)", fntText, Brushes.Black, rl, sf)

      rl = New Rectangle(600, r3, 300, rh)
      Me.DrawString(e.Graphics, Date.Now.ToShortDateString, fntText, Brushes.Black, rl, sf)

      rl = New Rectangle(150, r4, 550, rh)
      Me.DrawString(e.Graphics, fmSubject, fntText, Brushes.Black, rl, sf)

      rl = New Rectangle(130, r5 + 50, 650, 500)
      sf.Trimming = StringTrimming.None
      Me.DrawString(e.Graphics, fmBody, fntText, Brushes.Black, rl, sf)
    End Using

    e.HasMorePages = False

  End Sub

  Public Function PrintPortadaFax() As String
    Try
      Dim reg As RegistryKey
      Dim reg2 As RegistryKey
      Dim OldPrinterName As String = pd.PrinterSettings.PrinterName
      Dim dlg As New System.Windows.Forms.SaveFileDialog
      Dim PortadaFaxFileName As String

      pdPortadaFax.PrinterSettings.PrinterName = "pdfFactory"
      pdPortadaFax.PrinterSettings.DefaultPageSettings.Landscape = False
      PortadaFaxFileName = IO.Path.Combine(IO.Path.GetTempPath, IO.Path.GetFileNameWithoutExtension(IO.Path.GetTempFileName) + ".PDF")
      pdPortadaFax.DocumentName = PortadaFaxFileName

      'configurem el registre per poder imprimir fitxer en PDF
      'reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3\FinePrinters\pdfFactory\PrinterDriverData", True)
      reg = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3\FinePrinters\pdfFactory", True)
      reg.SetValue("ShowDlg", 2)
      reg.SetValue("PdfAction", 0)

      reg2 = Registry.CurrentUser.OpenSubKey("Software\FinePrint Software\pdfFactory3", True)
      reg2.SetValue("OutputFile", PortadaFaxFileName, RegistryValueKind.String)

      pdPortadaFax.Print()

      'Esperem a que el document estigui creat

      Do While reg2.GetValue("OutputFile") IsNot Nothing
        Threading.Thread.Sleep(500)
      Loop

      'Reiniciem el registre
      reg.SetValue("ShowDlg", 1)
      reg.SetValue("PdfAction", 0)
      'reg2.DeleteValue("OutputFilePerm", False)
      reg.Close()
      reg2.Close()

      Return PortadaFaxFileName
    Catch ex As Exception
      If ShowMessageError Then
        MsgBox("No se ha podido generar el fichero PDF.", MsgBoxStyle.Exclamation, "ERROR")
      End If
      Return Nothing
    End Try
  End Function

  Public Overrides Sub BeginPrint()

  End Sub

  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean

  End Function

  Protected Overrides Sub Print2Excel(ByVal FileName As String)

  End Sub

End Class


Public Class csTabularRpt
  Inherits csRpt

  Public Event GetTotalColumn(ByVal ColumnFieldName As String, ByRef TotalValue As Decimal)
  Public Event GetSubTotalColumn(ByVal ColumnFieldName As String, ByRef TotalValue As Decimal)
  Public Event Export2Excel(ByVal FileName As String, ByRef Handled As Boolean)

  Public Enum TotalColumnEnum
    None
    Sum
    Count
    Evaluated
  End Enum

  Public Enum ColumnDataKindEnum
    Normal
    MultipleLines
    IsBoolean
    IsImage
    CheckBox
    BoxToWriteIn
    BarCode
    FormLayout
    MultipleFields
    IndexedValue
  End Enum

  Public Enum FieldNameKindEnum
    Field
    Value
  End Enum

  Public Enum FormatingEnum
    StringFormat
    Custom
  End Enum

  Public Class ColumnInfo
    ''' <summary>
    ''' Titol de la columna
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Nom del camp al Datareader. El contingut del camp es el valor que se imprimirà.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldName As String
    Private mFieldFormat As String
    ''' <summary>
    ''' Format que s'aplicarà al valro del camp. Pot ser estandard del .NET o be custom.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormating As FormatingEnum
    ''' <summary>
    ''' Ens indica si el FieldName es el nom de un camp o es un literal.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldNameKind As FieldNameKindEnum
    ''' <summary>
    ''' Alineació del camp. Aplica a capçalera i valor. Pot ser Near, Center, Far
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment
    ''' <summary>
    ''' Font utilitzat a la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderBrush As Brush
    ''' <summary>
    ''' Font per a imprimir la linea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DetailRowFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la línea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DetailRowBrush As Brush
    ''' <summary>
    ''' Tipus de columna.
    ''' </summary>
    ''' <remarks></remarks>
    Public ColumnDataKind As ColumnDataKindEnum
    ''' <summary>
    ''' Si la columna es ColumnDataKindEnum.FormLayout, amplada dels titols dels camps.
    ''' </summary>
    ''' <remarks></remarks>
    Public FormFieldCaptionMaxWidth As Integer
    ''' <summary>
    ''' Directori on es troban les imatges a imprimir. El valor del camp serà el nom del fitxer de la imatge que es troba en aquest directori.
    ''' </summary>
    ''' <remarks></remarks>
    Public PathImages As String
    ''' <summary>
    ''' Tipus de d'agregació de columna: Suma / Contador
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalColumn As TotalColumnEnum
    ''' <summary>
    ''' Posició de la Columna. Es calcual automàticament.
    ''' </summary>
    ''' <remarks></remarks>
    Public PosX As Integer
    ''' <summary>
    ''' Ample de la columna en 1/100 de polsada.
    ''' </summary>
    ''' <remarks></remarks>
    Public Width As Integer
    ''' <summary>
    ''' Alsada disponible per a imprimir una imatge io un codi de barres.
    ''' </summary>
    ''' <remarks></remarks>
    Public ImageHeight As Integer ' per as Image i barCode
    '''' <summary>
    '''' Simbologia de codi de barres a imprimir
    '''' </summary>
    '''' <remarks></remarks>
    'Public BarCodeSymbol As csBarcode.csBarCode.BarcodeSymbologies

    ''' <summary>
    ''' Indica si s'ha de imprimir el codi sota el codi de barres
    ''' </summary>
    ''' <remarks></remarks>
    Public BarCodeDrawData As Boolean
    ''' <summary>
    ''' Indica el texte que ha quedat pendent de imprimir a una columna. Us intern
    ''' </summary>
    ''' <remarks></remarks>
    Public TextLeft As String
    ''' <summary>
    ''' Definició dels camps que composan una columna del tipus FormField
    ''' </summary>
    ''' <remarks></remarks>
    Public FormFields As System.Collections.Generic.List(Of FormFieldInfo)
    ''' <summary>
    ''' Array d'strings el index del qual es el valor retornat per datareader.
    ''' </summary>
    ''' <remarks></remarks>
    Public IndexedValue() As String
    ''' <summary>
    ''' Indica si imprimeix un valor si te el mateix valor que el registre anterior
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintRepeatedValues As Boolean
    ''' <summary>
    ''' Darrer valor impres. Utilitzat per PrintRepeatedValues.
    ''' </summary>
    ''' <remarks></remarks>
    Friend LastValuePrinted As String
    ''' <summary>
    ''' Valor del format .NET a aplicar al camp.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FieldFormat() As String
      Get
        Return mFieldFormat
      End Get
      Set(ByVal value As String)
        If String.IsNullOrEmpty(value) Then
          mFieldFormat = "{0}"
        Else
          mFieldFormat = "{0:" + value + "}"
        End If
      End Set
    End Property

    ''' <summary>
    ''' Afegir informació d'un camp dins dels camps que formen una columna del tipus FormField
    ''' </summary>
    ''' <param name="Caption">Literal que identifica el camp al llistat.</param>
    ''' <param name="FieldName">Nom del camp que s'imprimeix</param>
    ''' <param name="FormatValue">Format a aplicar</param>
    ''' <param name="FieldFormating">Tipus de format a que correspon el FormatValue</param>
    ''' <remarks></remarks>
    Public Sub AddFormField(ByVal Caption As String, ByVal FieldName As String, ByVal FormatValue As String, ByVal FieldFormating As FormatingEnum)
      Dim ff As New FormFieldInfo
      ff.Caption = Caption
      ff.FieldName = FieldName
      ff.FieldFormat = FormatValue
      ff.FieldFormating = FieldFormating
      If IsNothing(FormFields) Then
        FormFields = New System.Collections.Generic.List(Of FormFieldInfo)
      End If
      FormFields.Add(ff)
    End Sub

    ''' <summary>
    ''' Afegir informació d'un camp dins dels camps que formen una columna del tipus FormField. Aplica el format per defecte.
    ''' </summary>
    ''' <param name="Caption">Literal que identifica la agrupació</param>
    ''' <param name="FieldName">Nom del camp sobre el que s'agrupa.</param>
    ''' <param name="FieldFormat">Format que s'aplica al camp per la agrupació.</param>
    ''' <remarks></remarks>
    Public Sub AddFormField(ByVal Caption As String, ByVal FieldName As String, ByVal FieldFormat As String)
      AddFormField(Caption, FieldName, FieldFormat, FormatingEnum.StringFormat)
    End Sub

    ''' <summary>
    ''' Aplica de forma explicita el format a la columna
    ''' </summary>
    ''' <param name="Format">Format que s'aplica.</param>
    ''' <param name="Formating">Tipus de format al que correspon el Fromat</param>
    ''' <remarks></remarks>
    Public Sub SetFormat(ByVal Format As String, ByVal Formating As FormatingEnum)
      Me.FieldFormating = Formating
      If Formating = FormatingEnum.StringFormat Then
        ' Format de .Net
        Me.FieldFormat = Format
      Else
        ' Format Transform Custom
        Me.mFieldFormat = Format
      End If
    End Sub

  End Class

  Public Class FormLayoutRowInfo

    ''' <summary>
    ''' Font utilitzat a la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionBrush As Brush
    ''' <summary>
    ''' Font per a imprimir la linea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la línea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataBrush As Brush

    ''' <summary>
    ''' Definició dels camps que composan una columna del tipus FormField
    ''' </summary>
    ''' <remarks></remarks>
    Public FormFields As System.Collections.Generic.List(Of FormLayoutFieldInfo)

    ''' <summary>
    ''' Afegir informació d'un camp dins dels camps que formen una columna del tipus FormField
    ''' </summary>
    ''' <param name="Caption">Literal que identifica el camp al llistat.</param>
    ''' <param name="FieldName">Nom del camp que s'imprimeix</param>
    ''' <param name="FormatValue">Format a aplicar</param>
    ''' <param name="FieldFormating">Tipus de format a que correspon el FormatValue</param>
    ''' <remarks></remarks>
    Public Function AddField(ByVal Caption As String, ByVal CaptionWidth As Integer, ByVal FieldName As String, ByVal FieldWidth As Integer, ByVal FormatValue As String, ByVal FieldFormating As FormatingEnum, ByVal Alignment As StringAlignment) As FormLayoutFieldInfo
      Dim ff As New FormLayoutFieldInfo
      ff.Caption = Caption
      ff.FieldName = FieldName
      ff.FieldFormat = FormatValue
      ff.FieldFormating = FieldFormating

      ff.FieldWidth = FieldWidth
      ff.FieldNameKind = FieldNameKindEnum.Field
      ff.Aligment = Alignment
      ff.FieldDataKind = ColumnDataKindEnum.Normal

      ff.CaptionFont = CaptionFont
      ff.CaptionBrush = CaptionBrush

      ff.DataFont = DataFont
      ff.DataBrush = DataBrush

      If IsNothing(FormFields) Then
        FormFields = New System.Collections.Generic.List(Of FormLayoutFieldInfo)
      End If
      FormFields.Add(ff)
      Return ff
    End Function

  End Class

  Public Class FormLayoutFieldInfo

    ''' <summary>
    ''' Font utilitzat a la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionBrush As Brush
    ''' <summary>
    ''' Font per a imprimir la linea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la línea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataBrush As Brush

    ''' <summary>
    ''' Titol del camp
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Nom del camp al Datareader. El contingut del camp es el valor que se imprimirà.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldName As String
    Private mFieldFormat As String
    ''' <summary>
    ''' Format que s'aplicarà al valro del camp. Pot ser estandard del .NET o be custom.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormating As FormatingEnum
    ''' <summary>
    ''' Ens indica si el FieldName es el nom de un camp o es un literal.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldNameKind As FieldNameKindEnum
    ''' <summary>
    ''' Alineació del camp. Aplica a capçalera i valor. Pot ser Near, Center, Far
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment

    ''' <summary>
    ''' Tipus de columna.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldDataKind As ColumnDataKindEnum


    ''' <summary>
    ''' Si la columna es ColumnDataKindEnum.FormLayout, amplada dels titols dels camps.
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionWidth As Integer
    ''' <summary>
    ''' Directori on es troban les imatges a imprimir. El valor del camp serà el nom del fitxer de la imatge que es troba en aquest directori.
    ''' </summary>
    ''' <remarks></remarks>
    Public PathImages As String
    ''' <summary>
    ''' Tipus de d'agregació de columna: Suma / Contador
    ''' </summary>
    ''' <remarks></remarks>
    Public PosX As Integer
    ''' <summary>
    ''' Ample de la columna en 1/100 de polsada.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldWidth As Integer
    ''' <summary>
    ''' Alsada disponible per a imprimir una imatge io un codi de barres.
    ''' </summary>
    ''' <remarks></remarks>
    Public ImageHeight As Integer ' per as Image i barCode
    '''' <summary>
    '''' Simbologia de codi de barres a imprimir
    '''' </summary>
    '''' <remarks></remarks>
    'Public BarCodeSymbol As csBarcode.csBarCode.BarcodeSymbologies
    ''' <summary>
    ''' Indica si s'ha de imprimir el codi sota el codi de barres
    ''' </summary>
    ''' <remarks></remarks>
    Public BarCodeDrawData As Boolean
    ''' <summary>
    ''' Indica el texte que ha quedat pendent de imprimir a una columna. Us intern
    ''' </summary>
    ''' <remarks></remarks>
    Public TextLeft As String
    ''' <summary>
    ''' Array d'strings el index del qual es el valor retornat per datareader.
    ''' </summary>
    ''' <remarks></remarks>
    Public IndexedValue() As String
    Friend LastValuePrinted As String
    ''' <summary>
    ''' Valor del format .NET a aplicar al camp.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FieldFormat() As String
      Get
        Return mFieldFormat
      End Get
      Set(ByVal value As String)
        If String.IsNullOrEmpty(value) Then
          mFieldFormat = "{0}"
        Else
          mFieldFormat = "{0:" + value + "}"
        End If
      End Set
    End Property

    ''' <summary>
    ''' Aplica de forma explicita el format a la columna
    ''' </summary>
    ''' <param name="Format">Format que s'aplica.</param>
    ''' <param name="Formating">Tipus de format al que correspon el Fromat</param>
    ''' <remarks></remarks>
    Public Sub SetFormat(ByVal Format As String, ByVal Formating As FormatingEnum)
      Me.FieldFormating = Formating
      If Formating = FormatingEnum.StringFormat Then
        ' Format de .Net
        Me.FieldFormat = Format
      Else
        ' Format Transform Custom
        Me.mFieldFormat = Format
      End If
    End Sub

  End Class

  ''' <summary>
  ''' Clase per gestionar les agrupacions del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public Class GroupInfo
    ''' <summary>
    ''' Text fixe al titol de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderCaption As String
    ''' <summary>
    ''' Text fixe al peu de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FooterCaption As String
    ''' <summary>
    ''' Camp sobre el que es fa la agrupació. Cal que el datareader estigui ordenat per aquest camp.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldData As String
    ''' <summary>
    ''' Format a aplicar al camp sobre el que es fa la agrupació.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormat As String
    ''' <summary>
    ''' Camp descripció sobre el que es fa la agrupació.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldDescription As String
    ''' <summary>
    ''' Valor actual sobre el que s'esta agrupant. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public LastValue As String
    ''' <summary>
    ''' Valor actual sobre el que s'esta agrupant. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public LastDescription As String
    ''' <summary>
    ''' Indica si al començar un nou grup cal fer-ho a un nova pàgina.
    ''' </summary>
    ''' <remarks></remarks>
    Public StartOnNewPage As Boolean
    ''' <summary>
    ''' Estat en el que estoba una agruapció en un moment determinat durant la impressió. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public State As GroupStateEnum
    ''' <summary>
    ''' Col·lecció de TotalInfo de les columnes sobre les que cal fer agregació per a aquesta agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Totals As New System.Collections.Generic.List(Of TotalInfo)
    ''' <summary>
    ''' Indica si imprimeix la capçalera de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintHeaderCaption As Boolean
    ''' <summary>
    ''' Indica si imprimeix el peu de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintFooterCaption As Boolean
    ''' <summary>
    ''' Espai en 1/100" a deixar despres de imprimir el peu de grup. Se afegirà encara que PrintFooterCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceAfterFooter As Integer
    ''' <summary>
    ''' Espai en 1/100" a deixar abans de imprimir la capçalera de grup. Se afegirà encara que PrintHeaderCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceBeforeCaption As Integer

    ''' <summary>
    ''' Indica si es produeix un canvi de valor al camp sobre el que es fa la agrupació. Us Intern
    ''' </summary>
    ''' <param name="dr">Datareader sobre el que comprobar si hi ha un canvi d'agruapció</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBreak(ByVal dr As IDataReader) As Boolean
      Return (Me.LastValue <> String.Format("{0}", dr(FieldData)))
    End Function

    ''' <summary>
    ''' Inicialització al fer el canvi de grup. Us intern.
    ''' </summary>
    ''' <param name="Value">Nou valor de la agrupació</param>
    ''' <remarks></remarks>
    Public Sub Init(ByVal Value As String)
      Me.LastValue = Value
      Me.LastDescription = String.Empty
    End Sub

    Public Sub Init(ByVal Value As String, ByVal Description As String)
      Me.LastValue = Value
      Me.LastDescription = Description
    End Sub

    ''' <summary>
    ''' Reseteja els valors del totals al canviar de grup. Us intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
      For Each t As TotalInfo In Totals
        t.Reset()
      Next
    End Sub

    ''' <summary>
    ''' Actualitza els totals generals. Us intern.
    ''' </summary>
    ''' <param name="dr">DataReader que conte els valors dels camp a totalitzar.</param>
    ''' <remarks></remarks>
    Public Sub UpdateTotals(ByVal dr As IDataReader)
      For Each t As TotalInfo In Totals
        Select Case t.Col.TotalColumn
          Case TotalColumnEnum.Count
            t.Total += 1
          Case TotalColumnEnum.Sum
            If Not (IsDBNull(dr(t.Col.FieldName)) OrElse IsNothing(dr(t.Col.FieldName))) Then
              t.Total += CDec(dr(t.Col.FieldName))
            End If
        End Select
      Next
      State = GroupStateEnum.AddingRow
    End Sub

  End Class

  ''' <summary>
  ''' Definició del total de una columna.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class TotalInfo
    ''' <summary>
    ''' ColumnInfo de la columna sobre la que s'aplica un total.
    ''' </summary>
    ''' <remarks></remarks>
    Public Col As ColumnInfo
    ''' <summary>
    ''' Variable en la que es magatzema el total acumulat. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Total As Decimal
    ''' <summary>
    ''' Font amb el que s'imprimira el total
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalFont As Font
    ''' <summary>
    ''' 'Brush amb el que s'imprimira el total.
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalBrush As Brush

    ''' <summary>
    ''' Inicialitza la variable sobre la que s'acumula el total. Us intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
      Me.Total = 0
    End Sub

  End Class

  ''' <summary>
  ''' Clase empreada per magatzemar informació sobre els criteris de filtre que s'han aplicat per obtenir les dades.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class CriteriaInfo
    ''' <summary>
    ''' Texte identificatiu del filtre
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Valor del filtre.
    ''' </summary>
    ''' <remarks></remarks>
    Public Value As String
  End Class

  ''' <summary>
  ''' Informació sobre els camps que composen un FormField.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class FormFieldInfo
    ''' <summary>
    ''' Texte identificatiu del camp. Etiqueta.
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Nom del camp.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldName As String
    ''' <summary>
    ''' Fromat a aplicar al valor del camp.
    ''' </summary>
    ''' <remarks></remarks>
    Private mFieldFormat As String
    ''' <summary>
    ''' tipus de format a que correspon la propietat FieldFormat.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormating As FormatingEnum
    ''' <summary>
    ''' Valor a del camp a imprimir. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Value As String

    Public Property FieldFormat() As String
      Get
        Return mFieldFormat
      End Get
      Set(ByVal value As String)
        If String.IsNullOrEmpty(value) Then
          mFieldFormat = "{0}"
        Else
          mFieldFormat = "{0:" + value + "}"
        End If
      End Set
    End Property

    Public Sub New()

    End Sub
  End Class

  ''' <summary>
  ''' Agrupació de capçaleres. Agrupa en un nivell superior diverses capçaleres per donar coherencia a un grup de columnes. Amén.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class GroupCaptionInfo
    ''' <summary>
    ''' Columna inical de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FromColumn As Integer
    ''' <summary>
    ''' Columna final de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public ToColumn As Integer
    ''' <summary>
    ''' Texte de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Alineació de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment
    ''' <summary>
    ''' Si cal subratllar la agruapció
    ''' </summary>
    ''' <remarks></remarks>
    Public Underline As Boolean
  End Class

  ''' <summary>
  ''' Controla l'estat al imprimir un grup quan es produeix un salt de pàgina. Us Intern.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum GroupStateEnum
    ''' <summary>
    ''' Indica si s'ha impres la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    HeaderPrinted
    ''' <summary>
    ''' Indica si s'estan imprimint les linies del detall.
    ''' </summary>
    ''' <remarks></remarks>
    AddingRow
    ''' <summary>
    ''' Indica si ja s'ha impress el peu del grup.
    ''' </summary>
    ''' <remarks></remarks>
    FooterPrinted
  End Enum

#Region " Variables "

  Private ColumnsReport As New System.Collections.Generic.List(Of ColumnInfo)
  Private GroupsReport As New System.Collections.Generic.List(Of GroupInfo)
  Private Totals As New GroupInfo
  Private Criteria As New System.Collections.Generic.List(Of CriteriaInfo)
  Private CaptionGroups As New System.Collections.Generic.List(Of GroupCaptionInfo)

  Private DefaultHeaderFont As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultHeaderBrush As Brush = Brushes.Black

  Private DefaultFooterFont As New Font("Arial", 6, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultFooterBrush As Brush = Brushes.Black

  Private DefaultColumnCaptionFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultColumnCaptionBrush As Brush = Brushes.Black

  Private DefaultDetailRowFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultDetailRowBrush As Brush = Brushes.Black

  Private DefaultGroupHeaderFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultGroupHeaderBrush As Brush = Brushes.Black

  Private DefaultGroupFooterFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultGroupFooterBrush As Brush = Brushes.Black

  Private DefaultTotalFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultTotalBrush As Brush = Brushes.Black

  Private PenThick As New Pen(Color.Black, 2)
  Private PenThin As New Pen(Color.Black, 1)

  Private SubGroupLevel As Integer
  ''' <summary>
  ''' font amb el que s'imprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà el peu de pàgina del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà els encolumnas del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les línies de detall
  ''' </summary>
  ''' <remarks></remarks>
  Public RowFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les capçaleres de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran els peus de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran els totals de grup
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalFont As Font

  ''' <summary>
  ''' Brush amb el que simprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el peu de pàgian del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterBrush As Brush
  ''' <summary>
  ''' Pincell amb el que simprimiran les capçaleres de columna.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimiran les linies de detall del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public RowBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà la capçalera de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el preu de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterBrush As Brush
  ''' <summary>
  ''' Brush amb el que s
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalBrush As Brush
  ''' <summary>
  ''' Indica si cal imprimir un total general. Cal indicar quenes columnes cal totalitzar i el el tipus de total a aplicar.
  ''' </summary>
  ''' <remarks></remarks>
  Public TeTotalGeneral As Boolean
  ''' <summary>
  ''' Literal del total.
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalGeneralCaption As String
  ''' <summary>
  ''' columna sobre la que s'imprimeix el literal del total. Normalment alineat a la dreta.
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalCaptionColumn As Integer
  ''' <summary>
  ''' Separació en 1/100 de polsada entre les columnes del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnGap As Integer

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les files del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public RowGap As Integer

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les files del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRows As Boolean

  ''' <summary>
  ''' Indica si s'imprimeix linea obbrejada per facilitar la lectura.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PaperPijama As Boolean
  ''' <summary>
  ''' Indica cada cuantes linies impreses cal dibuijar la línea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaPeriode As Integer
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada pàgina. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaResetOnNewPage As Boolean
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada grup. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaResetOnNewGroup As Boolean
  ''' <summary>
  ''' Contador intern de linea.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaRowCount As Integer
  ''' <summary>
  ''' Brush utilitzar per dibuijar la linea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaBrush As Brush

#End Region

  Public Sub New()

    HeaderFont = DefaultHeaderFont
    FooterFont = DefaultFooterFont
    ColumnCaptionFont = DefaultColumnCaptionFont
    RowFont = DefaultDetailRowFont
    GroupHeaderFont = DefaultGroupHeaderFont
    GroupFooterFont = DefaultGroupFooterFont
    TotalFont = DefaultTotalFont

    HeaderBrush = DefaultHeaderBrush
    FooterBrush = DefaultFooterBrush
    ColumnCaptionBrush = ColumnCaptionBrush
    RowBrush = DefaultDetailRowBrush
    GroupHeaderBrush = DefaultGroupHeaderBrush
    GroupFooterBrush = DefaultGroupFooterBrush
    TotalBrush = DefaultTotalBrush

    ColumnGap = 5
    RowGap = 0
    HeaderKind = HeaderKindEnum.Plain
    LayoutOffset = LayoutOffsetEnum.Centered

    TeTotalGeneral = False
    PaperPijama = False
    PijamaPeriode = 3
    PijamaResetOnNewPage = True
    PijamaResetOnNewGroup = True
    PijamaBrush = New SolidBrush(Color.FromArgb(128, 220, 220, 220))

  End Sub

  ''' <summary>
  ''' Afegir columna al llistat.
  ''' </summary>
  ''' <param name="Caption">Titol de la columna</param>
  ''' <param name="Width">Ample de la columna en 1/100"</param>
  ''' <param name="FieldName">Nom del camp que s'imprimirà a la columna.</param>
  ''' <param name="Aligment">Alineació a aplicar a la columna</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function AddColumn(ByVal Caption As String, ByVal Width As Integer, ByVal FieldName As String, ByVal Aligment As StringAlignment) As ColumnInfo
    Dim Column As New ColumnInfo
    With Column
      .Caption = Caption
      .Width = Width
      .FieldFormat = ""
      .FieldFormating = FormatingEnum.StringFormat
      .FieldName = FieldName
      .FieldNameKind = FieldNameKindEnum.Field
      .Aligment = Aligment
      .ColumnDataKind = ColumnDataKindEnum.Normal
      .DetailRowFont = RowFont
      .DetailRowBrush = RowBrush
      .HeaderFont = HeaderFont
      .HeaderBrush = HeaderBrush
      .PrintRepeatedValues = True
      .LastValuePrinted = String.Empty
    End With
    ColumnsReport.Add(Column)
    Return Column
  End Function

  Public Function AddColumn(ByVal Width As Integer, ByVal Aligment As StringAlignment) As ColumnInfo
    Return AddColumn("", Width, "", Aligment)
  End Function

  Public Function AddFormLayoutRow() As FormLayoutRowInfo
    Dim row As New FormLayoutRowInfo
    With row
      .CaptionBrush = DefaultHeaderBrush
      .CaptionFont = DefaultHeaderFont
      .DataBrush = Me.DefaultDetailRowBrush
      .DataFont = Me.DefaultDetailRowFont
    End With
    Return row
  End Function

  Private Function IndexOfColumn(ByVal ColumnField As String) As Integer
    Dim Index As Integer = 0
    For Each c As ColumnInfo In ColumnsReport
      If c.FieldName.ToUpper = ColumnField.ToUpper Then
        Exit For
      End If
      Index += 1
    Next
    If Index = ColumnsReport.Count Then
      Index = -1
    End If
    Return Index
  End Function

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String)
    Dim Group As New GroupInfo
    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = True
      .PrintHeaderCaption = True
      .SpaceAfterFooter = 0
      .SpaceBeforeCaption = 0

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String, ByVal FieldDescription As String)
    Dim Group As New GroupInfo
    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .FieldDescription = FieldDescription
      .PrintFooterCaption = True
      .PrintHeaderCaption = True
      .SpaceAfterFooter = 0
      .SpaceBeforeCaption = 0

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <param name="PrintHeaderCaption">Indica si imprimeix la capçalera del grup</param>
  ''' <param name="PrintFooterCaption">Indica si imprimeix el peu de grup</param>
  ''' <param name="SpaceBeforeCaption">Espai en 1/100" anabs de imprimir la capçalera. S'afegeig sempre encara que no s'imprimeixi el grup </param>
  ''' <param name="SpaceAfterFooter">Espai en 1/100" despres de imprimir el peu del grup. S'afegeig sempre encara que no s'imprimeixi el grup</param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String, ByVal PrintHeaderCaption As Boolean, ByVal PrintFooterCaption As Boolean, ByVal SpaceBeforeCaption As Integer, ByVal SpaceAfterFooter As Integer)
    Dim Group As New GroupInfo

    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = PrintFooterCaption
      .PrintHeaderCaption = PrintHeaderCaption
      .SpaceAfterFooter = SpaceAfterFooter
      .SpaceBeforeCaption = SpaceBeforeCaption

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String, ByVal FieldDescription As String, ByVal PrintHeaderCaption As Boolean, ByVal PrintFooterCaption As Boolean, ByVal SpaceBeforeCaption As Integer, ByVal SpaceAfterFooter As Integer)
    Dim Group As New GroupInfo

    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .FieldDescription = FieldDescription
      .PrintFooterCaption = PrintFooterCaption
      .PrintHeaderCaption = PrintHeaderCaption
      .SpaceAfterFooter = SpaceAfterFooter
      .SpaceBeforeCaption = SpaceBeforeCaption

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegeix un criteri o filtre aplicat a les dades. S'imprimeix al principi del llistat.
  ''' </summary>
  ''' <param name="Caption">Titol del vaolr filtrat.</param>
  ''' <param name="Value">Valor aplicat al filtre.</param>
  ''' <remarks></remarks>
  Public Sub AddCriteria(ByVal Caption As String, ByVal Value As String)
    Dim c As New CriteriaInfo
    c.Caption = Caption
    c.Value = Value
    Criteria.Add(c)
  End Sub

  ''' <summary>
  ''' Afegeig una agrupació de columnes. al imprimir els titols de les columnes.
  ''' </summary>
  ''' <param name="FromColumn">Columna inicial</param>
  ''' <param name="ToColumn">Columan final</param>
  ''' <param name="Caption">Titol de la agrupació</param>
  ''' <param name="Aligment">Alineació a aplicar.</param>
  ''' <param name="Underlined">Subratllat.</param>
  ''' <remarks></remarks>
  Public Sub AddGroupCaption(ByVal FromColumn As Integer, ByVal ToColumn As Integer, ByVal Caption As String, ByVal Aligment As StringAlignment, ByVal Underlined As Boolean)
    Dim gc As New GroupCaptionInfo
    gc.FromColumn = FromColumn
    gc.ToColumn = ToColumn
    gc.Caption = Caption
    gc.Aligment = Aligment
    gc.Underline = Underlined
    CaptionGroups.Add(gc)
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="Value">True si es vol paper pijama.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal Value As Boolean)
    Me.PaperPijama = Value
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="ResetOnNewPage">Reinicialitza el contador de paper pijama al canviar de pàgina.</param>
  ''' <param name="ResetOnNewGroup">Reinicialitza el contador de paper pijama al canviar de grup.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal ResetOnNewPage As Boolean, ByVal ResetOnNewGroup As Boolean)
    Me.PaperPijama = True
    Me.PijamaResetOnNewGroup = ResetOnNewGroup
    Me.PijamaResetOnNewPage = ResetOnNewPage
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="ResetOnNewPage">Reinicialitza el contador de paper pijama al canviar de pàgina.</param>
  ''' <param name="ResetOnNewGroup">Reinicialitza el contador de paper pijama al canviar de grup.</param>
  ''' <param name="PijamaPeriode">Indica cada cuantes linies es dibuija la linea ombrejada.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal ResetOnNewPage As Boolean, ByVal ResetOnNewGroup As Boolean, ByVal PijamaPeriode As Integer)
    Me.PaperPijama = True
    Me.PijamaResetOnNewGroup = ResetOnNewGroup
    Me.PijamaResetOnNewPage = ResetOnNewPage
    Me.PijamaPeriode = PijamaPeriode
  End Sub

  Private Sub GroupsInit()
    For Each g As GroupInfo In GroupsReport
      If String.IsNullOrEmpty(g.FieldDescription) Then
        g.Init(String.Format(g.FieldFormat, DataSource(g.FieldData)))
      Else
        g.Init(String.Format(g.FieldFormat, DataSource(g.FieldData)), DataSource(g.FieldDescription).ToString)
      End If
      g.State = GroupStateEnum.FooterPrinted
      g.Reset()
    Next
  End Sub

  Overridable Sub DrawPageHeader(ByVal Canvas As Graphics)
    Dim sf As New StringFormat

    sf.Alignment = StringAlignment.Far
    Me.CurY = 5

    CurrentPage += 1

    Me.DrawLine(Canvas, Me.PenThick, 0, CurY, PageWidth, CurY)

    Me.CurY += 2
    Me.DrawString(Canvas, Me.EmpresaName, HeaderFont, HeaderBrush, 0, Me.CurY)
    Me.DrawString(Canvas, String.Format(Me.ReportName, CurrentPage), HeaderFont, HeaderBrush, PageWidth, CurY, sf)
    Me.CurY += CInt(HeaderFont.GetHeight(Canvas))
    Me.DrawLine(Canvas, Me.PenThin, 0, CurY, PageWidth, CurY)
    Me.CurY += 5

  End Sub

  Overridable Sub DrawPageFooter(ByVal Canvas As Graphics)
    Dim y As Integer
    Dim sf As New StringFormat
    y = CInt(PageHeight - FooterFont.GetHeight(Canvas)) - 2
    BottomY = y - 5

    Me.DrawLine(Canvas, Me.PenThin, 0, y, PageWidth, y)
    y += 1
    Me.DrawString(Canvas, String.Format("Usuari: {0} LT: {1} FM: {2}", Me.UserName, Me.WorkstationIP, Me.ReportID), FooterFont, HeaderBrush, 0, y)
    sf.Alignment = StringAlignment.Center
    Me.DrawString(Canvas, String.Format("Data: {0:dd/MM/yyyy HH:mm}", Date.Now), FooterFont, HeaderBrush, PageWidth \ 2, y, sf)
    sf.Alignment = StringAlignment.Far
    If PageNumbering = PageNumberEnum.PageN Then
      Me.DrawString(Canvas, String.Format("Pàgina: {0}", CurrentPage), FooterFont, HeaderBrush, PageWidth, y, sf)
    Else
      Me.DrawString(Canvas, String.Format("Pàgina: {0} de {1}", CurrentPage, TotalPages), FooterFont, HeaderBrush, PageWidth, y, sf)
    End If

  End Sub

  Overridable Sub DrawCriteria(ByVal Canvas As Graphics)
    If Criteria.Count = 0 Then
      Return
    End If

    Dim height As Integer = 0
    Dim rowHeight As Integer
    Dim MaxCaptionLen As Integer = 0
    Dim MaxValueLen As Integer = 0

    rowHeight = CInt(RowFont.GetHeight(Canvas))

    For Each c As CriteriaInfo In Criteria
      MaxCaptionLen = Math.Max(MaxCaptionLen, CInt(Canvas.MeasureString(c.Caption, RowFont).Width))
      MaxValueLen = Math.Max(MaxValueLen, CInt(Canvas.MeasureString(c.Value, RowFont).Width))
    Next

    CurY += 5

    Me.DrawRectangle(Canvas, Pens.Black, BodyLeft, CurY, 3 + MaxCaptionLen + 5 + MaxValueLen + 3, 3 + rowHeight * Criteria.Count + 3)
    CurY += 3

    For Each c As CriteriaInfo In Criteria
      Me.DrawString(Canvas, c.Caption, RowFont, RowBrush, BodyLeft + 3, CurY)
      Me.DrawString(Canvas, c.Value, RowFont, RowBrush, BodyLeft + 3 + MaxCaptionLen + 5, CurY)
      CurY += rowHeight
    Next

    CurY += 3 ' linea baixa del rectangle

  End Sub

  Overridable Function DrawGroupHeader(ByVal Canvas As System.Drawing.Graphics, ByVal group As GroupInfo) As Boolean
    '
    ' Calcular la alçada del text a imprimir
    Dim alsadaText, alsadaGrup As Integer

    If group.State = GroupStateEnum.HeaderPrinted Then
      Return True
    End If

    If Not group.PrintHeaderCaption Then
      CurY += group.SpaceBeforeCaption
      group.State = GroupStateEnum.HeaderPrinted
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaGrup = 5 + alsadaText + 5 + 3

    If CurY + alsadaGrup >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    CurY += 5


    Me.DrawLine(Canvas, Me.PenThick, BodyLeft, CurY, BodyLeft, CurY + alsadaText + 2)

    Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY + alsadaText + 2, BodyLeft + BodyWidth, CurY + alsadaText + 2)
    Me.DrawString(Canvas, String.Format("{0} {1} {2}", group.HeaderCaption, group.LastValue, group.LastDescription), Me.GroupHeaderFont, Me.GroupHeaderBrush, BodyLeft + 5, CurY)

    CurY += 5 + alsadaText + 3

    group.State = GroupStateEnum.HeaderPrinted

    Return True

  End Function

  Overridable Function DrawGroupFooter(ByVal Canvas As Graphics, ByVal gf As GroupInfo) As Boolean
    Dim alsadaText, alsadaGrup As Integer
    Dim totalValue As Decimal

    If gf.State = GroupStateEnum.FooterPrinted Then
      Return True
    End If

    If Not gf.PrintFooterCaption Then
      CurY += gf.SpaceAfterFooter
      gf.State = GroupStateEnum.FooterPrinted
      Return True
    End If

    If gf.FooterCaption = "-" Then
      ' Nomes imprimeix una linea de separació
      alsadaGrup = 2 + 1 + 2

      If CurY + alsadaGrup >= BottomY Then
        ' Si retorna False es que no hi cabia. 
        ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
        Return False
      End If

      CurY += 2

      Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY, BodyLeft + BodyWidth, CurY)

      CurY += 3

      gf.State = GroupStateEnum.FooterPrinted

      Return True

    End If

    If gf.Totals.Count = 0 Then
      ' No hi han columnes de totals
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaGrup = 2 + 2 + alsadaText + 3

    If CurY + alsadaGrup >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    Dim sf As New StringFormat
    sf.Alignment = StringAlignment.Far

    ' Imprimir les linies de total

    CurY += 2
    For Each t As TotalInfo In gf.Totals
      Me.DrawLine(Canvas, Me.PenThin, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next
    CurY += 2

    ' Imprimir el caption
    Dim pos As Integer
    pos = ColumnsReport(TotalCaptionColumn).PosX + ColumnsReport(TotalCaptionColumn).Width

    Me.DrawString(Canvas, String.Format("{0} {1}:", gf.FooterCaption, gf.LastValue), GroupFooterFont, GroupFooterBrush, pos, CurY, sf)

    For Each t As TotalInfo In gf.Totals
      pos = t.Col.PosX + t.Col.Width
      totalValue = t.Total
      If t.Col.TotalColumn = TotalColumnEnum.Evaluated Then
        RaiseEvent GetSubTotalColumn(t.Col.FieldName, totalValue)
      End If
      Me.DrawString(Canvas, String.Format(t.Col.FieldFormat, totalValue), GroupFooterFont, GroupFooterBrush, pos, CurY, sf)
    Next

    CurY += alsadaText + 3

    gf.State = GroupStateEnum.FooterPrinted

    Return True

  End Function

  Overridable Function DrawSummary(ByVal Canvas As Graphics) As Boolean
    Dim alsadaSummary As Integer
    Dim alsadaText As Integer
    Dim totalValue As Decimal

    If Totals.Totals.Count = 0 Then
      ' No hi han columnes de totals
      Return True
    End If

    If Totals.State = GroupStateEnum.FooterPrinted Then
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaSummary = 2 + 2 + alsadaText + 2

    If CurY + alsadaSummary >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    Dim sf As New StringFormat
    sf.Alignment = StringAlignment.Far

    ' Imprimir les linies de total

    CurY += 2
    For Each t As TotalInfo In Totals.Totals
      Me.DrawLine(Canvas, Me.PenThin, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next
    CurY += 2

    ' Imprimir el caption
    Dim pos As Integer
    pos = ColumnsReport(TotalCaptionColumn).PosX + ColumnsReport(TotalCaptionColumn).Width
    Me.DrawString(Canvas, TotalGeneralCaption, TotalFont, TotalBrush, pos, CurY, sf)

    For Each t As TotalInfo In Totals.Totals
      pos = t.Col.PosX + t.Col.Width
      totalValue = t.Total
      If t.Col.TotalColumn = Me.TotalColumnEnum.Evaluated Then
        RaiseEvent GetTotalColumn(t.Col.FieldName, totalValue)
      End If
      Me.DrawString(Canvas, String.Format(t.Col.FieldFormat, totalValue), TotalFont, TotalBrush, pos, CurY, sf)
    Next

    CurY += alsadaText + 2
    For Each t As TotalInfo In Totals.Totals
      Me.DrawLine(Canvas, Me.PenThick, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next

    Totals.Reset()
    Totals.State = GroupStateEnum.FooterPrinted

    Return True

  End Function

  Private Sub DrawColumnCaptions(ByVal Canvas As Graphics)

    Dim sf As New StringFormat
    Dim rh As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim layoutRec As RectangleF
    Dim layoutMultipleRec As RectangleF
    Dim MaxH As Integer

    CurY += 20

    If CaptionGroups.Count > 0 Then

      For Each gc As GroupCaptionInfo In CaptionGroups

        Dim FromX As Integer
        Dim Width As Integer

        FromX = ColumnsReport(gc.FromColumn).PosX
        Width = ColumnsReport(gc.ToColumn).PosX + ColumnsReport(gc.ToColumn).Width - FromX

        layoutRec = New RectangleF(FromX, CurY, Width, rh)
        sf.Alignment = gc.Aligment
        Me.DrawString(Canvas, gc.Caption, RowFont, RowBrush, layoutRec, sf)

        Me.DrawLine(Canvas, Me.PenThin, FromX, CurY + rh + 2, FromX + Width, CurY + rh + 2)

      Next

      CurY += rh + 2

    End If

    For Each c As ColumnInfo In ColumnsReport

      sf.Alignment = c.Aligment

      Select Case c.ColumnDataKind
        Case ColumnDataKindEnum.BarCode, ColumnDataKindEnum.BoxToWriteIn, ColumnDataKindEnum.CheckBox, ColumnDataKindEnum.IsBoolean, ColumnDataKindEnum.IsImage, ColumnDataKindEnum.MultipleLines, ColumnDataKindEnum.Normal
          layoutRec = New RectangleF(c.PosX, CurY, c.Width, rh)
          Me.DrawString(Canvas, c.Caption, RowFont, RowBrush, layoutRec, sf)
          MaxH = Math.Max(MaxH, rh)
        Case ColumnDataKindEnum.MultipleFields
          layoutMultipleRec = New RectangleF(c.PosX, CurY, c.Width, rh)
          For Each f As FormFieldInfo In c.FormFields
            Me.DrawString(Canvas, f.Caption, RowFont, RowBrush, layoutMultipleRec, sf)
            layoutMultipleRec.Y += rh
          Next
          MaxH = Math.Max(MaxH, CInt(layoutMultipleRec.Y) - CurY)
        Case ColumnDataKindEnum.FormLayout
          ' Nothing
          MaxH = 0
      End Select

    Next

    CurY += MaxH + 2

    For Each c As ColumnInfo In ColumnsReport

      Me.DrawLine(Canvas, Me.PenThin, c.PosX, CurY, c.PosX + c.Width, CurY)

    Next

    CurY += 2

  End Sub

  Private Function DrawColumn(ByVal Canvas As Graphics, ByVal col As ColumnInfo, ByVal Value As String) As Integer
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim rowHeight As Integer = height

    Select Case col.ColumnDataKind
      Case ColumnDataKindEnum.IsBoolean
        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case col.Aligment
          Case StringAlignment.Center
            x = col.PosX + col.Width \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = col.PosX + col.Width - height + offset * 2
          Case StringAlignment.Near
            x = col.PosX
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

        If Value.ToLower = "true" OrElse Value = "1" Then
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + offset * 2, x + costat - offset, CurY + costat)
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + costat, x + costat - offset, CurY + offset * 2)
        End If

      Case ColumnDataKindEnum.MultipleLines
        Dim alsadaText As Integer
        alsadaText = CInt(Canvas.MeasureString(Value, RowFont, col.Width).Height)

        Dim sf As New StringFormat
        sf.Alignment = col.Aligment
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        If alsadaText > Me.BottomY - CurY Then
          ' no hi cap
          Dim charsIn, linFilled As Integer

          Canvas.MeasureString(Value, RowFont, New Size(col.Width, Me.BottomY - CurY), sf, charsIn, linFilled)
          col.TextLeft = Value.Substring(charsIn)

          Me.DrawString(Canvas, Value.Substring(0, charsIn), RowFont, RowBrush, col.PosX, CurY)

          rowHeight = Me.BottomY

        Else

          Dim r As New Rectangle(col.PosX, CurY, col.Width, alsadaText)
          Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)

          rowHeight = alsadaText

        End If

      Case ColumnDataKindEnum.Normal, ColumnDataKindEnum.IndexedValue

        Dim sf As New StringFormat
        sf.Alignment = col.Aligment
        sf.Trimming = StringTrimming.EllipsisCharacter
        Dim r As New Rectangle(col.PosX, CurY, col.Width, height)

        If col.PrintRepeatedValues Then
          Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)
        Else
          If Value <> col.LastValuePrinted Then
            Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)
            col.LastValuePrinted = Value
          End If
        End If


      Case ColumnDataKindEnum.CheckBox

        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case col.Aligment
          Case StringAlignment.Center
            x = col.PosX + col.Width \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = col.PosX + col.Width - height + offset * 2
          Case StringAlignment.Near
            x = col.PosX
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

      Case ColumnDataKindEnum.BoxToWriteIn

        Dim offset As Integer = height \ 4

        Dim Calaix As PointF() = {New PointF(col.PosX, CurY + offset), _
          New PointF(col.PosX, CurY + height), _
          New PointF(col.PosX + col.Width, CurY + height), _
          New PointF(col.PosX + col.Width, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Calaix)

      Case ColumnDataKindEnum.IsImage
        ' el camp data del DataReader es el nom del fitxer si esta en un directori este esta al PathImages
        Dim Filename As String
        Dim img As Image
        Filename = IO.Path.Combine(col.PathImages, Value.Trim)
        If IO.File.Exists(Filename) Then
          img = Image.FromFile(Filename)
          Dim imgH As Integer
          Dim imgW As Integer
          Dim Ratio As Double
          imgH = CInt(img.Height \ CInt(img.VerticalResolution) \ 100)
          imgW = CInt(img.Width \ CInt(img.HorizontalResolution) \ 100)
          ' Faig cuadrar l'alsada
          Ratio = col.ImageHeight / imgH
          imgH = CInt(CDbl(imgH) / Ratio)
          imgW = CInt(CDbl(imgW) / Ratio)
          If imgW > col.Width Then
            'calcular el ratio de reduccio
            Ratio = col.Width / imgW
            imgH = CInt(CDbl(imgH) / Ratio)
            imgW = CInt(CDbl(imgW) / Ratio)
          End If

          ' Calculem la situació Horitzontal
          Dim imgPosX As Integer

          Select Case col.Aligment
            Case StringAlignment.Center
              imgPosX = col.PosX + (col.Width - imgW) \ 2
            Case StringAlignment.Far
              imgPosX = col.PosX + col.Width - imgW
            Case StringAlignment.Near
              imgPosX = col.PosX
          End Select

          Me.DrawImage(Canvas, img, New RectangleF(imgPosX, CurY, imgW, imgH), New RectangleF(0, 0, img.Width, img.Height), GraphicsUnit.Pixel)

          img.Dispose()

        End If

        rowHeight = col.ImageHeight

      Case ColumnDataKindEnum.BarCode

        'Dim bc As New csBarcode.csBarCode
        'bc.Data = Value
        'bc.Symbology = col.BarCodeSymbol
        'bc.DrawReadableData = col.BarCodeDrawData

        ''bc.DrawBarcode(Canvas, CurY, col.PosX, col.Width, col.ImageHeight)

        'bc = Nothing
        'rowHeight = col.ImageHeight

      Case ColumnDataKindEnum.MultipleFields

        Dim sf As New StringFormat
        Dim rh As Integer = CInt(RowFont.GetHeight(Canvas))
        Dim layoutValueRec As RectangleF
        Dim y As Integer

        sf.Alignment = StringAlignment.Near
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.EllipsisCharacter

        y = CurY

        For Each ff As FormFieldInfo In col.FormFields
          layoutValueRec = New RectangleF(col.PosX, y, col.Width, rh)

          Me.DrawString(Canvas, ff.Value, RowFont, RowBrush, layoutValueRec, sf)
          y += rh
        Next

        rowHeight = y - CurY

      Case ColumnDataKindEnum.FormLayout

        Dim sf As New StringFormat
        Dim rh As Integer = CInt(RowFont.GetHeight(Canvas))
        Dim colon As Integer = CInt(Canvas.MeasureString(": ", RowFont).Width)
        Dim layoutValueRec As RectangleF
        Dim y As Integer

        sf.Alignment = StringAlignment.Near
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.EllipsisCharacter

        y = CurY

        For Each ff As FormFieldInfo In col.FormFields
          Me.DrawString(Canvas, ff.Caption, RowFont, RowBrush, col.PosX, y)
          Me.DrawString(Canvas, ": ", RowFont, RowBrush, col.PosX + col.FormFieldCaptionMaxWidth, y)
          layoutValueRec = New RectangleF(col.PosX + col.FormFieldCaptionMaxWidth + colon, y, col.Width - col.FormFieldCaptionMaxWidth - colon, rh)

          Me.DrawString(Canvas, ff.Value, RowFont, RowBrush, layoutValueRec, sf)
          y += rh
        Next

        rowHeight = y - CurY

    End Select

    Return rowHeight

  End Function

  Private Function DrawColumnLeft(ByVal Canvas As Graphics, ByVal col As ColumnInfo) As Integer
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim rowHeight As Integer = height

    Dim alsadaText As Integer
    alsadaText = CInt(Canvas.MeasureString(col.TextLeft, RowFont, col.Width).Height)

    Dim sf As New StringFormat
    sf.Alignment = col.Aligment
    sf.FormatFlags = StringFormatFlags.LineLimit
    sf.Trimming = StringTrimming.Word

    If alsadaText > Me.BottomY - CurY Then
      ' no hi cap
      Dim charsIn, linFilled As Integer

      Canvas.MeasureString(col.TextLeft, RowFont, New Size(col.Width, Me.BottomY - CurY), sf, charsIn, linFilled)
      col.TextLeft = col.TextLeft.Substring(charsIn)

      Me.DrawString(Canvas, col.TextLeft.Substring(0, charsIn), RowFont, RowBrush, col.PosX, CurY)

      rowHeight = Me.BottomY

    Else

      Dim r As New Rectangle(col.PosX, CurY, col.Width, alsadaText)
      Me.DrawString(Canvas, col.TextLeft, RowFont, RowBrush, r, sf)

      col.TextLeft = String.Empty
      rowHeight = alsadaText

    End If

    Return rowHeight

  End Function

  Overridable Sub DrawRow(ByVal Canvas As Graphics)
    Dim rowHeight, maxRowHeight As Integer
    Dim Value As String

    maxRowHeight = 0

    If Me.PaperPijama Then
      If Me.PijamaRowCount Mod Me.PijamaPeriode = 0 Then
        Me.FillRectangle(Canvas, PijamaBrush, BodyLeft, CurY, BodyWidth, RowFont.GetHeight(Canvas))
      End If
      Me.PijamaRowCount += 1
    End If

    For Each c As ColumnInfo In ColumnsReport

      Select Case c.ColumnDataKind
        Case _
          ColumnDataKindEnum.BarCode, _
          ColumnDataKindEnum.IsBoolean, _
          ColumnDataKindEnum.Normal, _
          ColumnDataKindEnum.MultipleLines, _
          ColumnDataKindEnum.IsImage
          If c.FieldNameKind = FieldNameKindEnum.Field Then
            If c.FieldFormating = FormatingEnum.StringFormat Then
              Value = String.Format(c.FieldFormat, DataSource(c.FieldName))
            Else
              ' Custom
              Value = Utils.Transform(String.Format("{0}", DataSource(c.FieldName)), c.FieldFormat)
            End If
          Else
            Value = c.FieldName
          End If
        Case ColumnDataKindEnum.IndexedValue
          Try
            Value = c.IndexedValue(CInt(DataSource(c.FieldName)))
          Catch ex As Exception
            Value = String.Format(c.FieldFormat, DataSource(c.FieldName))
          End Try
        Case _
          ColumnDataKindEnum.MultipleFields, _
          ColumnDataKindEnum.FormLayout

          For Each ff As FormFieldInfo In c.FormFields
            If ff.FieldFormating = FormatingEnum.StringFormat Then
              ff.Value = String.Format(ff.FieldFormat, DataSource(ff.FieldName))
            Else
              ' Custom
              ff.Value = Utils.Transform(String.Format("{0}", DataSource(c.FieldName)), c.FieldFormat)
            End If
          Next

          Value = ""
      End Select

      rowHeight = DrawColumn(Canvas, c, Value)
      maxRowHeight = Math.Max(maxRowHeight, rowHeight)

    Next

    CurY += maxRowHeight

  End Sub

  Private Sub DrawRowLeft(ByVal Canvas As Graphics)
    Dim rowHeight, maxRowHeight As Integer
    Dim Value As String

    maxRowHeight = 0

    For Each c As ColumnInfo In ColumnsReport
      If String.IsNullOrEmpty(c.TextLeft) Then
        Continue For
      End If

      Value = String.Format(c.FieldFormat, DataSource(c.FieldName))
      rowHeight = DrawColumnLeft(Canvas, c)
      maxRowHeight = Math.Max(maxRowHeight, rowHeight)

    Next

    CurY += maxRowHeight

  End Sub

  Public Function TestGroupBreak(ByVal Canvas As Graphics) As Boolean

    If GroupsReport.Count = 0 Then
      Return True
    End If

    For i As Integer = 0 To GroupsReport.Count - 1

      If GroupsReport(i).LastValue <> String.Format(GroupsReport(i).FieldFormat, DataSource(GroupsReport(i).FieldData)) Then

        If Me.PaperPijama Then
          If Me.PijamaResetOnNewGroup Then
            Me.PijamaRowCount = 0
          End If
        End If

        For j As Integer = GroupsReport.Count - 1 To i Step -1
          SubGroupLevel = j
          If GroupsReport(j).State <> GroupStateEnum.FooterPrinted Then
            If Not DrawGroupFooter(Canvas, GroupsReport(j)) Then
              Return False
            End If
            GroupsReport(j).Reset()
            If String.IsNullOrEmpty(GroupsReport(j).FieldDescription) Then
              GroupsReport(j).Init(String.Format(GroupsReport(j).FieldFormat, DataSource(GroupsReport(j).FieldData)))
            Else
              GroupsReport(j).Init(String.Format(GroupsReport(j).FieldFormat, DataSource(GroupsReport(j).FieldData)), DataSource(GroupsReport(j).FieldDescription).ToString)
            End If

            GroupsReport(j).State = GroupStateEnum.FooterPrinted
          End If
        Next

        For j As Integer = i To GroupsReport.Count - 1
          If GroupsReport(j).State <> GroupStateEnum.HeaderPrinted Then
            If Not DrawGroupHeader(Canvas, GroupsReport(j)) Then
              Return False
            End If
          End If
        Next

        Exit For

      End If

    Next

    Return True

  End Function

  Private Function TestColumnsMultiField(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    For Each c As ColumnInfo In Me.ColumnsReport
      If c.FormFields IsNot Nothing Then
        If CurY + c.ImageHeight + RowGap > BottomY Then
          Return False
        End If
      End If
    Next
    Return True
  End Function

  Protected Sub InitLayout(ByVal Canvas As System.Drawing.Graphics)

    If LayoutInitialized Then
      Return
    End If

    LayoutInitialized = True

    'Calcul de posicions de les(columnes)
    ' Ample total de les columnes
    Dim AmpleColumnes As Integer = 0

    For Each c As ColumnInfo In ColumnsReport
      AmpleColumnes += c.Width
    Next

    AmpleColumnes += (ColumnsReport.Count - 1) * ColumnGap

    Select Case LayoutOffset
      Case LayoutOffsetEnum.Centered
        LeftOffset = (PageWidth - AmpleColumnes) \ 2
      Case LayoutOffsetEnum.OneThird
        LeftOffset = (PageWidth - AmpleColumnes) \ 3
      Case LayoutOffsetEnum.Custom
        ' Nothing to do
    End Select

    BodyLeft = LeftOffset
    BodyWidth = AmpleColumnes

    For Each c As ColumnInfo In ColumnsReport
      c.PosX = LeftOffset
      LeftOffset += ColumnGap + c.Width
    Next

    'Configurar Columnes FormFieldLayout 
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    For Each c As ColumnInfo In ColumnsReport
      If c.FormFields IsNot Nothing Then
        c.ImageHeight += height
      End If
    Next

    ' Configurar totals
    For Each c As ColumnInfo In ColumnsReport
      If c.TotalColumn > TotalColumnEnum.None Then
        For Each g As GroupInfo In GroupsReport
          Dim t As New TotalInfo
          t.Col = c
          t.Reset()
          g.Totals.Add(t)
        Next

        If TeTotalGeneral Then
          Dim tg As New TotalInfo
          tg.Col = c
          tg.Reset()
          Totals.Totals.Add(tg)
        End If

      End If
    Next

  End Sub

  ''' <summary>
  ''' Imprimeix una pàgina del llistat. Us intern.
  ''' </summary>
  ''' <param name="Canvas">Objecte graphcs sobre el que es dibuixa la pàgina.</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean

    If FirstPassReport Then

      If DataNeeded Then
        DataNeeded = False
        Do While True
          If Not DataSource.Read Then
            Return False
          End If
          If Not FilterRowOut(DataSource) Then
            Exit Do
          End If
        Loop
      End If

      If Me.Destination = ReportDestinationEnum.Preview Then
        PageHeight = CInt(pd.DefaultPageSettings.Bounds.Height)
        PageWidth = CInt(pd.DefaultPageSettings.Bounds.Width)
      Else
        PageHeight = CInt(Canvas.VisibleClipBounds.Height)
        PageWidth = CInt(Canvas.VisibleClipBounds.Width)
      End If

      InitLayout(Canvas)

      GroupsInit()
      Totals.Reset()

    End If

    DrawPageHeader(Canvas)
    DrawPageFooter(Canvas)

    If FirstPassReport Then
      DrawCriteria(Canvas)
      FirstPassReport = False
    End If

    DrawColumnCaptions(Canvas)

    DrawRowLeft(Canvas) ' potser falta llògica de control si la segona pasada contnua ocupant + de 1 pàgina

    For Each g As GroupInfo In GroupsReport
      If g.State = GroupStateEnum.HeaderPrinted Then
        Continue For
      End If
      If Not DrawGroupHeader(Canvas, g) Then
        ' Ull. que es repeteix el FirstPass
        Return True
      End If
    Next

    If Me.PaperPijama Then
      If Me.PijamaResetOnNewPage Then
        Me.PijamaRowCount = 0
      End If
    End If

    For Each col As ColumnInfo In ColumnsReport
      If Not col.PrintRepeatedValues Then
        col.LastValuePrinted = String.Empty
      End If
    Next

    Do While True

      If DrawingTotalsAndExit Then
        Exit Do
      End If

      If DataNeeded Then
        DataNeeded = False
        If Not DataSource.Read Then
          Exit Do
        End If
        If FilterRowOut(DataSource) Then
          DataNeeded = True
          Continue Do
        End If
      End If

      If Not TestGroupBreak(Canvas) Then
        Return True
      End If

      If Not TestColumnsMultiField(Canvas) Then
        Return True
      End If

      DrawRow(Canvas)

      ' Actualitza subtotals
      For i As Integer = 0 To GroupsReport.Count - 1
        GroupsReport(i).UpdateTotals(DataSource)
      Next
      'Actualitza Total General
      Totals.UpdateTotals(DataSource)

      DataNeeded = True

      DrawLineRow(Canvas)

      If CurY + RowFont.GetHeight(Canvas) + RowGap > BottomY Then
        ' no hi ha espai. 
        Return True
      End If

      CurY += RowGap

    Loop

    DrawingTotalsAndExit = True

    For i As Integer = GroupsReport.Count - 1 To 0 Step -1
      If Not DrawGroupFooter(Canvas, GroupsReport(i)) Then
        Return True
      End If
    Next

    If Not DrawSummary(Canvas) Then
      Return True
    End If

    Return False

  End Function

  Private Sub DrawLineRow(ByVal Canvas As System.Drawing.Graphics)
    If DrawLineBetweenRows Then
      If CurY + 2 >= BottomY Then
        ' si esta a final de pàgina ja no cal dibuixar la línea
        Return
      End If
      CurY += 2
      Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY, BodyLeft + BodyWidth, CurY)
    End If
  End Sub

  ''' <summary>
  ''' Inicialitza el llistat. Us intern.
  ''' </summary>
  ''' <remarks></remarks>
  Public Overrides Sub BeginPrint()
    FirstPassReport = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False

  End Sub

  Protected Overrides Sub Finalize()
    DefaultHeaderFont.Dispose()
    DefaultFooterFont.Dispose()
    DefaultColumnCaptionFont.Dispose()
    DefaultDetailRowFont.Dispose()
    DefaultGroupHeaderFont.Dispose()
    DefaultGroupFooterFont.Dispose()
    DefaultTotalFont.Dispose()
    PijamaBrush.Dispose()
    MyBase.Finalize()
  End Sub

  Public Function GetSubTotalValue(ByVal ColumnFieldName As String) As Decimal
    Dim Value As Decimal
    For Each t As TotalInfo In GroupsReport(SubGroupLevel).Totals
      If t.Col.FieldName.ToLower = ColumnFieldName.ToLower Then
        Value = t.Total
        Exit For
      End If
    Next
    Return Value

  End Function

  Public Function GetTotalValue(ByVal ColumnFieldName As String) As Decimal
    Dim Value As Decimal
    For Each t As TotalInfo In Me.Totals.Totals
      If t.Col.FieldName.ToLower = ColumnFieldName.ToLower Then
        Value = t.Total
        Exit For
      End If
    Next
    Return Value
  End Function

  Protected Overrides Sub Print2Excel(ByVal FileName As String)
    Dim Handled As Boolean
    Handled = False
    RaiseEvent Export2Excel(FileName, Handled)
    If Not Handled Then
      DefaultPrint2Excel(FileName)
    End If
  End Sub

  Public Sub DefaultPrint2Excel(ByVal FileName As String)
    Dim xls As New C1.C1Excel.C1XLBook()
    Dim sheet As C1.C1Excel.XLSheet = xls.Sheets("Sheet1")
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim value As String
    colCount = 0

    For Each c As ColumnInfo In ColumnsReport
      sheet(0, colCount).Value = c.Caption
      colCount += 1
    Next
    rowCount = 0
    Do While DataSource.Read
      rowCount += 1
      colCount = 0
      For Each c As ColumnInfo In ColumnsReport
        Select Case c.ColumnDataKind
          Case _
            ColumnDataKindEnum.BarCode, _
            ColumnDataKindEnum.IsBoolean, _
            ColumnDataKindEnum.Normal, _
            ColumnDataKindEnum.MultipleLines, _
            ColumnDataKindEnum.IsImage
            If c.FieldNameKind = FieldNameKindEnum.Field Then
              If c.FieldFormating = FormatingEnum.StringFormat Then
                value = String.Format(c.FieldFormat, DataSource(c.FieldName))
              Else
                ' Custom
                value = Utils.Transform(String.Format("{0}", DataSource(c.FieldName)), c.FieldFormat)
              End If
            Else
              value = c.FieldName
            End If

            Select Case DataSource.GetDataTypeName(DataSource.GetOrdinal(c.FieldName)).ToLower
              Case "int", "integer"
                If Not c.FieldName.ToUpper.EndsWith("ID") Then
                  If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                    sheet(rowCount, colCount).Value = CInt(DataSource(c.FieldName))
                  Else
                    sheet(rowCount, colCount).Value = 0
                  End If
                Else
                  sheet(rowCount, colCount).Value = value
                End If
              Case "decimal"
                If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                  sheet(rowCount, colCount).Value = CDec(DataSource(c.FieldName))
                Else
                  sheet(rowCount, colCount).Value = 0D
                End If
              Case "datetime"
                If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                  sheet(rowCount, colCount).Value = CStr(Date.Parse(String.Format("{0}", DataSource(c.FieldName))))
                Else
                  sheet(rowCount, colCount).Value = value
                End If
              Case Else
                sheet(rowCount, colCount).Value = value
            End Select

          Case _
            ColumnDataKindEnum.MultipleFields, _
            ColumnDataKindEnum.FormLayout

            value = "Format N.V."
        End Select

        colCount += 1

      Next

    Loop

    DataSource.Close()

    xls.Save(FileName)

  End Sub

End Class

Public Class csFormLayoutRpt
  Inherits csRpt

  Public Enum ColumnDataKindEnum
    Normal
    MultipleLines
    IsBoolean
    IsImage
    CheckBox
    BoxToWriteIn
    BarCode
    IndexedValue
  End Enum

  Public Enum FieldNameKindEnum
    Field
    Value
  End Enum

  Public Enum FormatingEnum
    StringFormat
    Custom
  End Enum

  Public Class FormLayoutRowInfo

    Public RowID As Integer
    Public CaptionFont As Font
    Public CaptionBrush As Brush
    Public DataFont As Font
    Public DataBrush As Brush
    Public Width As Integer

    Public FormFields As System.Collections.Generic.List(Of FormLayoutFieldInfo)

    Public Function AddField(ByVal Caption As String, ByVal CaptionWidth As Integer, ByVal FieldName As String, ByVal FieldWidth As Integer, ByVal FormatValue As String, ByVal FieldFormating As FormatingEnum, ByVal Alignment As StringAlignment) As FormLayoutFieldInfo

      Dim ff As New FormLayoutFieldInfo
      ff.Caption = Caption
      ff.FieldName = FieldName
      ff.FieldFormat = FormatValue
      ff.FieldFormating = FieldFormating

      ff.FieldWidth = FieldWidth
      ff.FieldNameKind = FieldNameKindEnum.Field
      ff.Aligment = Alignment
      ff.FieldDataKind = ColumnDataKindEnum.Normal

      ff.CaptionWidth = CaptionWidth
      ff.CaptionFont = CaptionFont
      ff.CaptionBrush = CaptionBrush

      ff.DataFont = DataFont
      ff.DataBrush = DataBrush

      If IsNothing(FormFields) Then
        FormFields = New System.Collections.Generic.List(Of FormLayoutFieldInfo)
      End If

      FormFields.Add(ff)

      Return ff
    End Function

  End Class

  Public Class FormLayoutFieldInfo

    ''' <summary>
    ''' Font utilitzat a la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionBrush As Brush
    ''' <summary>
    ''' Font per a imprimir la linea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la línea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DataBrush As Brush

    ''' <summary>
    ''' Titol del camp
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String

    ''' <summary>
    ''' Ample del text del caption
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionTextWidth As Integer

    ''' <summary>
    ''' Nom del camp al Datareader. El contingut del camp es el valor que se imprimirà.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldName As String
    Private mFieldFormat As String
    ''' <summary>
    ''' Format que s'aplicarà al valro del camp. Pot ser estandard del .NET o be custom.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormating As FormatingEnum
    ''' <summary>
    ''' Ens indica si el FieldName es el nom de un camp o es un literal.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldNameKind As FieldNameKindEnum
    ''' <summary>
    ''' Alineació del camp. Aplica a capçalera i valor. Pot ser Near, Center, Far
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment

    ''' <summary>
    ''' Tipus de columna.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldDataKind As ColumnDataKindEnum


    ''' <summary>
    ''' Si la columna es ColumnDataKindEnum.FormLayout, amplada dels titols dels camps.
    ''' </summary>
    ''' <remarks></remarks>
    Public CaptionWidth As Integer
    ''' <summary>
    ''' Directori on es troban les imatges a imprimir. El valor del camp serà el nom del fitxer de la imatge que es troba en aquest directori.
    ''' </summary>
    ''' <remarks></remarks>
    Public PathImages As String
    ''' <summary>
    ''' Tipus de d'agregació de columna: Suma / Contador
    ''' </summary>
    ''' <remarks></remarks>
    Public PosX As Integer
    ''' <summary>
    ''' Ample de la columna en 1/100 de polsada.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldWidth As Integer
    ''' <summary>
    ''' Alsada disponible per a imprimir una imatge io un codi de barres.
    ''' </summary>
    ''' <remarks></remarks>
    Public ImageHeight As Integer ' per as Image i barCode
    '''' <summary>
    '''' Simbologia de codi de barres a imprimir
    '''' </summary>
    '''' <remarks></remarks>
    'Public BarCodeSymbol As csBarcode.csBarCode.BarcodeSymbologies
    ''' <summary>
    ''' Indica si s'ha de imprimir el codi sota el codi de barres
    ''' </summary>
    ''' <remarks></remarks>
    Public BarCodeDrawData As Boolean
    ''' <summary>
    ''' Indica el texte que ha quedat pendent de imprimir a una columna. Us intern
    ''' </summary>
    ''' <remarks></remarks>
    Public TextLeft As String
    ''' <summary>
    ''' Array d'strings el index del qual es el valor retornat per datareader.
    ''' </summary>
    ''' <remarks></remarks>
    Public IndexedValue() As String
    Friend LastValuePrinted As String
    ''' <summary>
    ''' Valor del format .NET a aplicar al camp.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FieldFormat() As String
      Get
        Return mFieldFormat
      End Get
      Set(ByVal value As String)
        If String.IsNullOrEmpty(value) Then
          mFieldFormat = "{0}"
        Else
          mFieldFormat = "{0:" + value + "}"
        End If
      End Set
    End Property

    ''' <summary>
    ''' Aplica de forma explicita el format a la columna
    ''' </summary>
    ''' <param name="Format">Format que s'aplica.</param>
    ''' <param name="Formating">Tipus de format al que correspon el Fromat</param>
    ''' <remarks></remarks>
    Public Sub SetFormat(ByVal Format As String, ByVal Formating As FormatingEnum)
      Me.FieldFormating = Formating
      If Formating = FormatingEnum.StringFormat Then
        ' Format de .Net
        Me.FieldFormat = Format
      Else
        ' Format Transform Custom
        Me.mFieldFormat = Format
      End If
    End Sub

  End Class

  ''' <summary>
  ''' Clase empreada per magatzemar informació sobre els criteris de filtre que s'han aplicat per obtenir les dades.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class CriteriaInfo
    ''' <summary>
    ''' Texte identificatiu del filtre
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Valor del filtre.
    ''' </summary>
    ''' <remarks></remarks>
    Public Value As String
  End Class

  ''' <summary>
  ''' Clase per gestionar les agrupacions del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class GroupInfo
    ''' <summary>
    ''' Text fixe al titol de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderCaption As String
    ''' <summary>
    ''' Text fixe al peu de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FooterCaption As String
    ''' <summary>
    ''' Camp sobre el que es fa la agrupació. Cal que el datareader estigui ordenat per aquest camp.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldData As String
    ''' <summary>
    ''' Format a aplicar al camp sobre el que es fa la agrupació.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormat As String
    ''' <summary>
    ''' Camp de la descripció sobre el que es fa la agrupació. Cal que el datareader estigui ordenat per aquest camp.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldDescription As String
    ''' <summary>
    ''' Valor actual sobre el que s'esta agrupant. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public LastValue As String
    ''' <summary>
    ''' Descripció del valor actual sobre el que s'esta agrupant. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public LastDescription As String
    ''' <summary>
    ''' Indica si al començar un nou grup cal fer-ho a un nova pàgina.
    ''' </summary>
    ''' <remarks></remarks>
    Public StartOnNewPage As Boolean
    ''' <summary>
    ''' Estat en el que estoba una agruapció en un moment determinat durant la impressió. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public State As GroupStateEnum
    ''' <summary>
    ''' Indica si imprimeix la capçalera de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintHeaderCaption As Boolean
    ''' <summary>
    ''' Indica si imprimeix el peu de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintFooterCaption As Boolean

    ''' <summary>
    ''' Espai en 1/100" a deixar despres de imprimir el peu de grup. Se afegirà encara que PrintFooterCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceAfterFooter As Integer
    ''' <summary>
    ''' Espai en 1/100" a deixar abans de imprimir la capçalera de grup. Se afegirà encara que PrintHeaderCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceBeforeCaption As Integer

    ''' <summary>
    ''' Indica si es produeix un canvi de valor al camp sobre el que es fa la agrupació. Us Intern
    ''' </summary>
    ''' <param name="dr">Datareader sobre el que comprobar si hi ha un canvi d'agruapció</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBreak(ByVal dr As IDataReader) As Boolean
      Return (Me.LastValue <> String.Format("{0}", dr(FieldData)))
    End Function

    ''' <summary>
    ''' Inicialització al fer el canvi de grup. Us intern.
    ''' </summary>
    ''' <param name="Value">Nou valor de la agrupació</param>
    ''' <remarks></remarks>
    Public Sub Init(ByVal Value As String)
      Me.LastValue = Value
      Me.LastDescription = String.Empty
    End Sub

    ''' <summary>
    ''' Inicialització al fer el canvi de grup. Us intern.
    ''' </summary>
    ''' <param name="Value">Nou valor de la agrupació</param>
    ''' <remarks></remarks>
    Public Sub Init(ByVal Value As String, ByVal Description As String)
      Me.LastValue = Value
      Me.LastDescription = Description
    End Sub

    ''' <summary>
    ''' Reseteja els valors del totals al canviar de grup. Us intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
    End Sub

  End Class

  ''' <summary>
  ''' Agrupació de capçaleres. Agrupa en un nivell superior diverses capçaleres per donar coherencia a un grup de columnes. Amén.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class GroupCaptionInfo
    ''' <summary>
    ''' Columna inical de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FromColumn As Integer
    ''' <summary>
    ''' Columna final de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public ToColumn As Integer
    ''' <summary>
    ''' Texte de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Alineació de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment
    ''' <summary>
    ''' Si cal subratllar la agruapció
    ''' </summary>
    ''' <remarks></remarks>
    Public Underline As Boolean
  End Class

  ''' <summary>
  ''' Controla l'estat al imprimir un grup quan es produeix un salt de pàgina. Us Intern.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum GroupStateEnum
    ''' <summary>
    ''' Indica si s'ha impres la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    HeaderPrinted
    ''' <summary>
    ''' Indica si s'estan imprimint les linies del detall.
    ''' </summary>
    ''' <remarks></remarks>
    AddingRow
    ''' <summary>
    ''' Indica si ja s'ha impress el peu del grup.
    ''' </summary>
    ''' <remarks></remarks>
    FooterPrinted
  End Enum

#Region " Variables "

  Private FormLayoutRows As New System.Collections.Generic.List(Of FormLayoutRowInfo)
  Private GroupsReport As New System.Collections.Generic.List(Of GroupInfo)
  Private Totals As New GroupInfo
  Private Criteria As New System.Collections.Generic.List(Of CriteriaInfo)
  Private CaptionGroups As New System.Collections.Generic.List(Of GroupCaptionInfo)

  Private DefaultHeaderFont As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultHeaderBrush As Brush = Brushes.Black

  Private DefaultFooterFont As New Font("Arial", 6, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultFooterBrush As Brush = Brushes.Black

  Private DefaultColumnCaptionFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultColumnCaptionBrush As Brush = Brushes.Black

  Private DefaultDetailRowFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultDetailRowBrush As Brush = Brushes.Black

  Private DefaultGroupHeaderFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultGroupHeaderBrush As Brush = Brushes.Black

  Private DefaultGroupFooterFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultGroupFooterBrush As Brush = Brushes.Black

  Private DefaultTotalFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultTotalBrush As Brush = Brushes.Black

  Private PenThick As New Pen(Color.Black, 2)
  Private PenThin As New Pen(Color.Black, 1)
  Private LastRowIDPrinted As Integer

  ''' <summary>
  ''' font amb el que s'imprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà el peu de pàgina del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà els encolumnas del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les línies de detall
  ''' </summary>
  ''' <remarks></remarks>
  Public RowFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les capçaleres de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran els peus de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterFont As Font

  ''' <summary>
  ''' Si imprimeix linea de punts al imprimir el caption.
  ''' </summary>
  ''' <remarks></remarks>
  Public PrintCaptionDots As Boolean

  ''' <summary>
  ''' Si imprimeix linea de punts al imprimir el caption.
  ''' </summary>
  ''' <remarks></remarks>
  Public GapBetweenDots As Integer

  ''' <summary>
  ''' Brush amb el que simprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el peu de pàgian del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterBrush As Brush
  ''' <summary>
  ''' Pincell amb el que simprimiran les capçaleres de columna.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimiran les linies de detall del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public RowBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà la capçalera de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el preu de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterBrush As Brush

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les columnes del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public CaptionFieldGap As Integer

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les columnes del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FieldsGap As Integer


  ''' <summary>
  ''' Separació en 1/100 de polsada entre les files del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public RowGap As Integer

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les files del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRows As Boolean


#End Region

  Public Sub New()

    HeaderFont = DefaultHeaderFont
    FooterFont = DefaultFooterFont
    ColumnCaptionFont = DefaultColumnCaptionFont
    RowFont = DefaultDetailRowFont
    GroupHeaderFont = DefaultGroupHeaderFont
    GroupFooterFont = DefaultGroupFooterFont

    HeaderBrush = DefaultHeaderBrush
    FooterBrush = DefaultFooterBrush
    ColumnCaptionBrush = ColumnCaptionBrush
    RowBrush = DefaultDetailRowBrush
    GroupHeaderBrush = DefaultGroupHeaderBrush
    GroupFooterBrush = DefaultGroupFooterBrush

    CaptionFieldGap = 8
    FieldsGap = 25
    RowGap = 10
    DrawLineBetweenRows = True
    PrintCaptionDots = True
    GapBetweenDots = 5

    HeaderKind = HeaderKindEnum.Plain
    LayoutOffset = LayoutOffsetEnum.Centered

  End Sub

  Public Function AddRow() As FormLayoutRowInfo
    Dim row As New FormLayoutRowInfo
    With row
      .RowID = FormLayoutRows.Count
      .CaptionBrush = DefaultHeaderBrush
      .CaptionFont = DefaultHeaderFont
      .DataBrush = Me.DefaultDetailRowBrush
      .DataFont = Me.DefaultDetailRowFont
    End With
    FormLayoutRows.Add(row)
    Return row
  End Function

  Private Function IndexOfRow(ByVal RowID As Integer) As Integer
    Dim Index As Integer = 0
    For Each r As FormLayoutRowInfo In Me.FormLayoutRows
      If r.RowID = RowID Then
        Exit For
      End If
      Index += 1
    Next
    If Index = FormLayoutRows.Count Then
      Index = -1
    End If
    Return Index
  End Function

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String)
    Dim Group As New GroupInfo
    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = True
      .PrintHeaderCaption = True
      .SpaceAfterFooter = 0
      .SpaceBeforeCaption = 0

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <param name="PrintHeaderCaption">Indica si imprimeix la capçalera del grup</param>
  ''' <param name="PrintFooterCaption">Indica si imprimeix el peu de grup</param>
  ''' <param name="SpaceBeforeCaption">Espai en 1/100" anabs de imprimir la capçalera. S'afegeig sempre encara que no s'imprimeixi el grup </param>
  ''' <param name="SpaceAfterFooter">Espai en 1/100" despres de imprimir el peu del grup. S'afegeig sempre encara que no s'imprimeixi el grup</param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String, ByVal PrintHeaderCaption As Boolean, ByVal PrintFooterCaption As Boolean, ByVal SpaceBeforeCaption As Integer, ByVal SpaceAfterFooter As Integer)
    Dim Group As New GroupInfo

    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = PrintFooterCaption
      .PrintHeaderCaption = PrintHeaderCaption
      .SpaceAfterFooter = SpaceAfterFooter
      .SpaceBeforeCaption = SpaceBeforeCaption

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegeix un criteri o filtre aplicat a les dades. S'imprimeix al principi del llistat.
  ''' </summary>
  ''' <param name="Caption">Titol del vaolr filtrat.</param>
  ''' <param name="Value">Valor aplicat al filtre.</param>
  ''' <remarks></remarks>
  Public Sub AddCriteria(ByVal Caption As String, ByVal Value As String)
    Dim c As New CriteriaInfo
    c.Caption = Caption
    c.Value = Value
    Criteria.Add(c)
  End Sub

  ''' <summary>
  ''' Afegeig una agrupació de columnes. al imprimir els titols de les columnes.
  ''' </summary>
  ''' <param name="FromColumn">Columna inicial</param>
  ''' <param name="ToColumn">Columan final</param>
  ''' <param name="Caption">Titol de la agrupació</param>
  ''' <param name="Aligment">Alineació a aplicar.</param>
  ''' <param name="Underlined">Subratllat.</param>
  ''' <remarks></remarks>
  Public Sub AddGroupCaption(ByVal FromColumn As Integer, ByVal ToColumn As Integer, ByVal Caption As String, ByVal Aligment As StringAlignment, ByVal Underlined As Boolean)
    Dim gc As New GroupCaptionInfo
    gc.FromColumn = FromColumn
    gc.ToColumn = ToColumn
    gc.Caption = Caption
    gc.Aligment = Aligment
    gc.Underline = Underlined
    CaptionGroups.Add(gc)
  End Sub

  Private Sub GroupsInit()
    For Each g As GroupInfo In GroupsReport
      g.Init(String.Format(g.FieldFormat, (DataSource(g.FieldData))))
      g.State = GroupStateEnum.FooterPrinted
      g.Reset()
    Next
  End Sub

  Private Sub DrawPageHeader(ByVal Canvas As Graphics)
    Dim sf As New StringFormat

    sf.Alignment = StringAlignment.Far
    Me.CurY = 5

    CurrentPage += 1

    Me.DrawLine(Canvas, Me.PenThick, 0, CurY, PageWidth, CurY)
    Me.CurY += 2
    Me.DrawString(Canvas, Me.EmpresaName, HeaderFont, HeaderBrush, 0, Me.CurY)
    Me.DrawString(Canvas, String.Format(Me.ReportName, CurrentPage), HeaderFont, HeaderBrush, PageWidth, CurY, sf)
    Me.CurY += CInt(HeaderFont.GetHeight(Canvas))
    Me.DrawLine(Canvas, Me.PenThin, 0, CurY, PageWidth, CurY)
    Me.CurY += 5

  End Sub

  Private Sub DrawPageFooter(ByVal Canvas As Graphics)
    Dim y As Integer
    Dim sf As New StringFormat
    y = CInt(PageHeight - FooterFont.GetHeight(Canvas)) - 2
    BottomY = y - 5

    Me.DrawLine(Canvas, Me.PenThin, 0, y, PageWidth, y)
    y += 1
    Me.DrawString(Canvas, String.Format("Usuari: {0} LT: {1} FM: {2}", Me.UserName, Me.WorkstationIP, Me.ReportID), FooterFont, HeaderBrush, 0, y)
    sf.Alignment = StringAlignment.Center
    Me.DrawString(Canvas, String.Format("Data: {0:dd/MM/yyyy HH:mm}", Date.Now), FooterFont, HeaderBrush, PageWidth \ 2, y, sf)
    sf.Alignment = StringAlignment.Far
    If PageNumbering = PageNumberEnum.PageN Then
      Me.DrawString(Canvas, String.Format("Pàgina: {0}", CurrentPage), FooterFont, HeaderBrush, PageWidth, y, sf)
    Else
      Me.DrawString(Canvas, String.Format("Pàgina: {0} de {1}", CurrentPage, TotalPages), FooterFont, HeaderBrush, PageWidth, y, sf)
    End If

  End Sub

  Private Sub DrawCriteria(ByVal Canvas As Graphics)
    If Criteria.Count = 0 Then
      Return
    End If

    Dim height As Integer = 0
    Dim rowHeight As Integer
    Dim MaxCaptionLen As Integer = 0
    Dim MaxValueLen As Integer = 0

    rowHeight = CInt(RowFont.GetHeight(Canvas))

    For Each c As CriteriaInfo In Criteria
      MaxCaptionLen = Math.Max(MaxCaptionLen, CInt(Canvas.MeasureString(c.Caption, RowFont).Width))
      MaxValueLen = Math.Max(MaxValueLen, CInt(Canvas.MeasureString(c.Value, RowFont).Width))
    Next

    CurY += 5

    Me.DrawRectangle(Canvas, Pens.Black, BodyLeft, CurY, 3 + MaxCaptionLen + 5 + MaxValueLen + 3, 3 + rowHeight * Criteria.Count + 3)

    CurY += 3

    For Each c As CriteriaInfo In Criteria
      Me.DrawString(Canvas, c.Caption, RowFont, RowBrush, BodyLeft + 3, CurY)
      Me.DrawString(Canvas, c.Value, RowFont, RowBrush, BodyLeft + 3 + MaxCaptionLen + 5, CurY)
      CurY += rowHeight
    Next

    CurY += 3 ' linea baixa del rectangle

  End Sub

  Private Function DrawGroupHeader(ByVal Canvas As System.Drawing.Graphics, ByVal group As GroupInfo) As Boolean
    '
    ' Calcular la alçada del text a imprimir
    Dim alsadaText, alsadaGrup As Integer

    If group.State = GroupStateEnum.HeaderPrinted Then
      Return True
    End If

    If Not group.PrintHeaderCaption Then
      CurY += group.SpaceBeforeCaption
      group.State = GroupStateEnum.HeaderPrinted
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaGrup = 5 + alsadaText + 5 + 3

    If CurY + alsadaGrup >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    CurY += 5

    Me.DrawLine(Canvas, Me.PenThick, BodyLeft, CurY, BodyLeft, CurY + alsadaText + 2)
    Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY + alsadaText + 2, BodyLeft + BodyWidth, CurY + alsadaText + 2)

    Me.DrawString(Canvas, String.Format("{0} {1}", group.HeaderCaption, group.LastValue), Me.GroupHeaderFont, Me.GroupHeaderBrush, BodyLeft + 5, CurY)

    CurY += 5 + alsadaText + 3

    group.State = GroupStateEnum.HeaderPrinted

    Return True

  End Function

  Private Function DrawField(ByVal Canvas As Graphics, ByVal fld As FormLayoutFieldInfo, ByVal Value As String) As Integer
    Dim height As Integer = CInt(fld.CaptionFont.GetHeight(Canvas))
    Dim rowHeight As Integer = height

    Me.DrawString(Canvas, fld.Caption, fld.CaptionFont, fld.CaptionBrush, fld.PosX, CurY)

    If Me.PrintCaptionDots Then
      Me.DrawString(Canvas, fld.Caption, fld.CaptionFont, fld.CaptionBrush, fld.PosX, CurY)
      Me.DrawString(Canvas, ":", fld.DataFont, fld.DataBrush, fld.PosX + fld.CaptionWidth, CurY)
      For i As Integer = fld.PosX + fld.CaptionWidth - 5 To fld.PosX + fld.CaptionTextWidth + 2 Step -GapBetweenDots
        Me.DrawString(Canvas, ".", fld.DataFont, fld.DataBrush, i, CurY)
      Next
    Else
      Me.DrawString(Canvas, fld.Caption + ":", fld.CaptionFont, fld.CaptionBrush, fld.PosX, CurY)
    End If


    Select Case fld.FieldDataKind
      Case ColumnDataKindEnum.IsBoolean
        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case fld.Aligment
          Case StringAlignment.Center
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth - height + offset * 2
          Case StringAlignment.Near
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap)
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

        If Value.ToLower = "true" OrElse Value = "1" Then
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + offset * 2, x + costat - offset, CurY + costat)
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + costat, x + costat - offset, CurY + offset * 2)
        End If

      Case ColumnDataKindEnum.Normal, ColumnDataKindEnum.IndexedValue

        Dim sf As New StringFormat
        sf.Alignment = fld.Aligment
        sf.Trimming = StringTrimming.EllipsisCharacter
        Dim r As New Rectangle((fld.PosX + fld.CaptionWidth + CaptionFieldGap), CurY, fld.FieldWidth, height)

        Me.DrawString(Canvas, Value, fld.DataFont, fld.DataBrush, r, sf)

      Case ColumnDataKindEnum.CheckBox

        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case fld.Aligment
          Case StringAlignment.Center
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth - height + offset * 2
          Case StringAlignment.Near
            x = (fld.PosX + fld.CaptionWidth + CaptionFieldGap)
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

      Case ColumnDataKindEnum.BoxToWriteIn
        Dim offset As Integer = height \ 4
        Dim Calaix As PointF() = {New PointF((fld.PosX + fld.CaptionWidth + CaptionFieldGap), CurY + offset), _
          New PointF((fld.PosX + fld.CaptionWidth + CaptionFieldGap), CurY + height), _
          New PointF((fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth, CurY + height), _
          New PointF((fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Calaix)

      Case ColumnDataKindEnum.IsImage
        ' el camp data del DataReader es el nom del fitxer si esta en un directori este esta al PathImages
        Dim Filename As String
        Dim img As Image
        Filename = IO.Path.Combine(fld.PathImages, Value.Trim)
        If IO.File.Exists(Filename) Then
          img = Image.FromFile(Filename)
          Dim imgH As Integer
          Dim imgW As Integer
          Dim Ratio As Double
          imgH = CInt(img.Height \ CInt(img.VerticalResolution) \ 100)
          imgW = CInt(img.Width \ CInt(img.HorizontalResolution) \ 100)
          ' Faig cuadrar l'alsada
          Ratio = fld.ImageHeight / imgH
          imgH = CInt(CDbl(imgH) / Ratio)
          imgW = CInt(CDbl(imgW) / Ratio)
          If imgW > fld.FieldWidth Then
            'calcular el ratio de reduccio
            Ratio = fld.FieldWidth / imgW
            imgH = CInt(CDbl(imgH) / Ratio)
            imgW = CInt(CDbl(imgW) / Ratio)
          End If

          ' Calculem la situació Horitzontal
          Dim imgPosX As Integer

          Select Case fld.Aligment
            Case StringAlignment.Center
              imgPosX = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + (fld.FieldWidth - imgW) \ 2
            Case StringAlignment.Far
              imgPosX = (fld.PosX + fld.CaptionWidth + CaptionFieldGap) + fld.FieldWidth - imgW
            Case StringAlignment.Near
              imgPosX = fld.PosX
          End Select

          Me.DrawImage(Canvas, img, New RectangleF(imgPosX, CurY, imgW, imgH), New RectangleF(0, 0, img.Width, img.Height), GraphicsUnit.Pixel)

          img.Dispose()

        End If

        rowHeight = fld.ImageHeight

      Case ColumnDataKindEnum.BarCode

        'Dim bc As New csBarcode.csBarCode
        'bc.Data = Value
        'bc.Symbology = col.BarCodeSymbol
        'bc.DrawReadableData = col.BarCodeDrawData

        ''bc.DrawBarcode(Canvas, CurY, col.PosX, col.Width, col.ImageHeight)

        'bc = Nothing
        'rowHeight = col.ImageHeight

    End Select

    Return rowHeight

  End Function

  Private Sub DrawRow(ByVal Canvas As Graphics)
    Dim rowHeight, maxRowHeight As Integer
    Dim Value As String

    maxRowHeight = 0

    For Each r As FormLayoutRowInfo In FormLayoutRows
      For Each f As FormLayoutFieldInfo In r.FormFields

        Select Case f.FieldDataKind
          Case _
            ColumnDataKindEnum.BarCode, _
            ColumnDataKindEnum.IsBoolean, _
            ColumnDataKindEnum.Normal, _
            ColumnDataKindEnum.MultipleLines, _
            ColumnDataKindEnum.IsImage
            If f.FieldNameKind = FieldNameKindEnum.Field Then
              If f.FieldFormating = FormatingEnum.StringFormat Then
                Value = String.Format(f.FieldFormat, DataSource(f.FieldName))
              Else
                ' Custom
                Value = Utils.Transform(String.Format("{0}", DataSource(f.FieldName)), f.FieldFormat)
              End If
            Else
              Value = f.FieldName
            End If
          Case ColumnDataKindEnum.IndexedValue
            Try
              Value = f.IndexedValue(CInt(DataSource(f.FieldName)))
            Catch ex As Exception
              Value = String.Format(f.FieldFormat, DataSource(f.FieldName))
            End Try
        End Select

        rowHeight = DrawField(Canvas, f, Value)
        maxRowHeight = Math.Max(maxRowHeight, rowHeight)
      Next
      CurY += maxRowHeight
    Next


  End Sub

  Public Function TestGroupBreak(ByVal Canvas As Graphics) As Boolean

    If GroupsReport.Count = 0 Then
      Return True
    End If

    For i As Integer = 0 To GroupsReport.Count - 1

      If GroupsReport(i).LastValue <> String.Format(GroupsReport(i).FieldFormat, DataSource(GroupsReport(i).FieldData)) Then

        For j As Integer = GroupsReport.Count - 1 To i Step -1
          If GroupsReport(j).State <> GroupStateEnum.FooterPrinted Then
            GroupsReport(j).Reset()
            If String.IsNullOrEmpty(GroupsReport(j).FieldDescription) Then
              GroupsReport(j).Init(String.Format(GroupsReport(j).FieldFormat, DataSource(GroupsReport(j).FieldData)))
            Else
              GroupsReport(j).Init(String.Format(GroupsReport(j).FieldFormat, DataSource(GroupsReport(j).FieldData)), DataSource(GroupsReport(j).FieldDescription).ToString)
            End If
            GroupsReport(j).State = GroupStateEnum.FooterPrinted
          End If
        Next

        For j As Integer = i To GroupsReport.Count - 1
          If GroupsReport(j).State <> GroupStateEnum.HeaderPrinted Then
            If Not DrawGroupHeader(Canvas, GroupsReport(j)) Then
              Return False
            End If
          End If
        Next

        Exit For

      End If

    Next

    Return True

  End Function

  Protected Sub InitLayout(ByVal Canvas As System.Drawing.Graphics)

    If LayoutInitialized Then
      Return
    End If

    LayoutInitialized = True

    'Calcul de posicions de les(columnes)
    ' Ample total de les columnes
    Dim FieldsWidth As Integer
    BodyWidth = 0

    For Each r As FormLayoutRowInfo In FormLayoutRows
      FieldsWidth = 0
      For Each f As FormLayoutFieldInfo In r.FormFields
        f.CaptionTextWidth = CInt(Canvas.MeasureString(f.Caption, f.CaptionFont).Width)
        If f.CaptionWidth = 0 Then
          FieldsWidth += f.CaptionTextWidth + CaptionFieldGap + f.FieldWidth + FieldsGap
        Else
          FieldsWidth += f.CaptionWidth + CaptionFieldGap + f.FieldWidth + FieldsGap
        End If
      Next
      FieldsWidth -= FieldsGap
      If BodyWidth < FieldsWidth Then
        BodyWidth = FieldsWidth
      End If
    Next

    Select Case LayoutOffset
      Case LayoutOffsetEnum.Centered
        LeftOffset = (PageWidth - FieldsWidth) \ 2
      Case LayoutOffsetEnum.OneThird
        LeftOffset = (PageWidth - FieldsWidth) \ 3
      Case LayoutOffsetEnum.Custom
        ' Nothing to do
    End Select

    BodyLeft = LeftOffset

    Dim curx As Integer

    For Each r As FormLayoutRowInfo In FormLayoutRows
      curx = BodyLeft
      For Each f As FormLayoutFieldInfo In r.FormFields
        f.PosX = curx
        curx += f.CaptionWidth + CaptionFieldGap + f.FieldWidth + FieldsGap
      Next
    Next

  End Sub

  ''' <summary>
  ''' Imprimeix una pàgina del llistat. Us intern.
  ''' </summary>
  ''' <param name="Canvas">Objecte graphcs sobre el que es dibuixa la pàgina.</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean

    If FirstPassReport Then

      PageHeight = CInt(Canvas.VisibleClipBounds.Height)
      PageWidth = CInt(Canvas.VisibleClipBounds.Width)

      InitLayout(Canvas)

      FirstPassReport = False

    End If

    DrawPageHeader(Canvas)
    DrawPageFooter(Canvas)
    If FirstPassReport Then
      DrawCriteria(Canvas)
    End If

    DrawLineRow(Canvas)

    For Each g As GroupInfo In GroupsReport
      If g.State = GroupStateEnum.HeaderPrinted Then
        Continue For
      End If
      If Not DrawGroupHeader(Canvas, g) Then
        ' Ull. que es repeteix el FirstPass
        Return True
      End If
    Next

    Do While True

      If DrawingTotalsAndExit Then
        Exit Do
      End If

      If DataNeeded Then
        DataNeeded = False
        Do While True
          If Not DataSource.Read Then
            Return False
          End If
          If Not FilterRowOut(DataSource) Then
            Exit Do
          End If
        Loop
      End If

      If Not TestGroupBreak(Canvas) Then
        Return True
      End If

      DrawRow(Canvas)
      DrawLineRow(Canvas)

      DataNeeded = True

      If CurY + RowFont.GetHeight(Canvas) + RowGap > BottomY Then
        ' no hi ha espai. 
        Return True
      End If

      CurY += RowGap

    Loop

    Return False

  End Function

  Private Sub DrawLineRow(ByVal Canvas As System.Drawing.Graphics)
    If DrawLineBetweenRows Then
      If CurY + 2 > BottomY Then
        ' si esta a final de pàgina ja no cal dibuixar la línea
        Return
      End If
      CurY += 2
      Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY, BodyLeft + BodyWidth, CurY)
    End If
  End Sub

  ''' <summary>
  ''' Inicialitza el llistat. Us intern.
  ''' </summary>
  ''' <remarks></remarks>
  Public Overrides Sub BeginPrint()
    FirstPassReport = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False
  End Sub

  Protected Overrides Sub Finalize()
    DefaultHeaderFont.Dispose()
    DefaultFooterFont.Dispose()
    DefaultColumnCaptionFont.Dispose()
    DefaultDetailRowFont.Dispose()
    DefaultGroupHeaderFont.Dispose()
    DefaultGroupFooterFont.Dispose()
    DefaultTotalFont.Dispose()
    MyBase.Finalize()
  End Sub

  Protected Overrides Sub Print2Excel(ByVal FileName As String)

    Dim xls As New C1.C1Excel.C1XLBook()
    Dim sheet As C1.C1Excel.XLSheet = xls.Sheets("Sheet1")
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim value As String
    colCount = 0

    For Each r As FormLayoutRowInfo In FormLayoutRows
      For Each f As FormLayoutFieldInfo In r.FormFields
        sheet(0, colCount).Value = f.Caption
        colCount += 1
      Next
    Next
    rowCount = 0
    Do While DataSource.Read

      If FilterRowOut(DataSource) Then
        Continue Do
      End If

      rowCount += 1
      colCount = 0
      For Each r As FormLayoutRowInfo In FormLayoutRows
        For Each f As FormLayoutFieldInfo In r.FormFields
          Select Case f.FieldDataKind
            Case _
              ColumnDataKindEnum.BarCode, _
              ColumnDataKindEnum.IsBoolean, _
              ColumnDataKindEnum.Normal, _
              ColumnDataKindEnum.MultipleLines, _
              ColumnDataKindEnum.IsImage
              If f.FieldNameKind = FieldNameKindEnum.Field Then
                If f.FieldFormating = FormatingEnum.StringFormat Then
                  value = String.Format(f.FieldFormat, DataSource(f.FieldName))
                Else
                  ' Custom
                  value = Utils.Transform(String.Format("{0}", DataSource(f.FieldName)), f.FieldFormat)
                End If
              Else
                value = f.FieldName
              End If
          End Select
          sheet(rowCount, colCount).Value = value
          colCount += 1
        Next
      Next


    Loop
    xls.Save(FileName)

  End Sub

End Class

Public Class csFreeLayoutRpt
  Inherits csRpt

  Public Event PaintHeader(ByVal Canvas As Graphics, ByVal CurrentPage As Single, ByVal TotalPages As Single)
  Public Event PaintColumnCaptions(ByVal Canvas As Graphics, ByRef ColumnCaptionsHaveBeenPrinted As Boolean)
  Public Event PaintCriteria(ByVal Canvas As Graphics)
  Public Event PaintGroupHeader(ByVal Canvas As Graphics, ByVal group As GroupInfo, ByRef GroupHeaderHasBeenPrinted As Boolean)
  Public Event PaintRow(ByVal Canvas As Graphics, ByVal Row As IDataReader, ByRef RowHasBeenPrinted As Boolean)
  Public Event PaintGroupFooter(ByVal Canvas As Graphics, ByVal group As GroupInfo, ByRef GroupFooterHasBeenPrinted As Boolean)
  Public Event PaintSummary(ByVal Canvas As Graphics, ByVal group As GroupInfo, ByRef SummaryHasBeenPrinted As Boolean)
  Public Event PaintFooter(ByVal Canvas As Graphics)
  Public Event Export2Excel(ByVal FileName As String, ByRef Handled As Boolean)
  Public Event InitializeReportValues()
  Public Event EventValue(ByVal Row As IDataReader, ByVal Column As ColumnInfo, ByRef NewValue As String)
  Public Event UpdateValuesBeforeRowPrinted(ByVal Row As IDataReader, ByRef ProcessRow As Boolean)
  Public Event UpdateValuesAfterRowPrinted(ByVal Row As IDataReader)
  Public Event GetTotalColumn(ByVal ColumnFieldName As String, ByRef TotalValue As Decimal)
  Public Event GetSubTotalColumn(ByVal ColumnFieldName As String, ByRef TotalValue As Decimal)

#Region " Properties & enums "
  'Fem accessibles el Numero de pàgina i el total de pàgines
  Public ReadOnly Property PaginaActual() As Single
    Get
      Return CurrentPage
    End Get
  End Property

  ReadOnly Property PaginesTotals() As Single
    Get
      Return TotalPages
    End Get
  End Property

  'Fem Accessible la Coordenada y
  Public Property PosY() As Integer
    Get
      Return CurY
    End Get
    Set(ByVal value As Integer)
      CurY = value
    End Set
  End Property

  Public ReadOnly Property Page_BottomY() As Integer
    Get
      Return BottomY
    End Get
  End Property

  Public ReadOnly Property Page_Height() As Integer
    Get
      Return PageHeight
    End Get
  End Property

  Public ReadOnly Property Page_Width() As Integer
    Get
      Return PageWidth
    End Get
  End Property

  Public ReadOnly Property EvaluatingPageCount() As Boolean
    Get
      Return EvaluatingTotalPages
    End Get
  End Property

  Public Enum TotalColumnEnum
    None
    Sum
    Count
    Evaluated
  End Enum

  Public Enum ColumnDataKindEnum
    Normal
    MultipleLines
    IsBoolean
    IsImage
    CheckBox
    BoxToWriteIn
    BarCode
    FormLayout
    MultipleFields
    IndexedValue
    EventValue
  End Enum

  Public Enum FieldNameKindEnum
    Field
    Value
  End Enum

  Public Enum FormatingEnum
    StringFormat
    Custom
  End Enum

  ''' <summary>
  ''' Controla l'estat al imprimir un grup quan es produeix un salt de pàgina. Us Intern.
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum GroupStateEnum
    ''' <summary>
    ''' Indica si s'ha impres la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    HeaderPrinted
    ''' <summary>
    ''' Indica si s'estan imprimint les linies del detall.
    ''' </summary>
    ''' <remarks></remarks>
    AddingRow
    ''' <summary>
    ''' Indica si ja s'ha impress el peu del grup.
    ''' </summary>
    ''' <remarks></remarks>
    FooterPrinted
  End Enum

#End Region

#Region " Helper classes "

  Public Class ColumnInfo
    ''' <summary>
    ''' Titol de la columna
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Nom del camp al Datareader. El contingut del camp es el valor que se imprimirà.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldName As String
    Private mFieldFormat As String
    ''' <summary>
    ''' Format que s'aplicarà al valro del camp. Pot ser estandard del .NET o be custom.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormating As FormatingEnum
    ''' <summary>
    ''' Ens indica si el FieldName es el nom de un camp o es un literal.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldNameKind As FieldNameKindEnum
    ''' <summary>
    ''' Alineació del camp. Aplica a capçalera i valor. Pot ser Near, Center, Far
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment
    ''' <summary>
    ''' Font utilitzat a la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la capçalera
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderBrush As Brush
    ''' <summary>
    ''' Font per a imprimir la linea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DetailRowFont As Font
    ''' <summary>
    ''' Brush utilitzat per imprimir la línea de detall del llistat
    ''' </summary>
    ''' <remarks></remarks>
    Public DetailRowBrush As Brush
    ''' <summary>
    ''' Tipus de columna.
    ''' </summary>
    ''' <remarks></remarks>
    Public ColumnDataKind As ColumnDataKindEnum
    ''' <summary>
    ''' Si la columna es ColumnDataKindEnum.FormLayout, amplada dels titols dels camps.
    ''' </summary>
    ''' <remarks></remarks>
    Public FormFieldCaptionMaxWidth As Integer
    ''' <summary>
    ''' Directori on es troban les imatges a imprimir. El valor del camp serà el nom del fitxer de la imatge que es troba en aquest directori.
    ''' </summary>
    ''' <remarks></remarks>
    Public PathImages As String
    ''' <summary>
    ''' Tipus de d'agregació de columna: Suma / Contador
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalColumn As TotalColumnEnum
    ''' <summary>
    ''' Posició de la Columna. Es calcual automàticament.
    ''' </summary>
    ''' <remarks></remarks>
    Public PosX As Integer
    ''' <summary>
    ''' Ample de la columna en 1/100 de polsada.
    ''' </summary>
    ''' <remarks></remarks>
    Public Width As Integer
    ''' <summary>
    ''' Alsada disponible per a imprimir una imatge io un codi de barres.
    ''' </summary>
    ''' <remarks></remarks>
    Public ImageHeight As Integer ' per as Image i barCode
    '''' <summary>
    '''' Simbologia de codi de barres a imprimir
    '''' </summary>
    '''' <remarks></remarks>
    'Public BarCodeSymbol As csBarcode.csBarCode.BarcodeSymbologies

    ''' <summary>
    ''' Indica si s'ha de imprimir el codi sota el codi de barres
    ''' </summary>
    ''' <remarks></remarks>
    Public BarCodeDrawData As Boolean
    ''' <summary>
    ''' Indica el texte que ha quedat pendent de imprimir a una columna. Us intern
    ''' </summary>
    ''' <remarks></remarks>
    Public TextLeft As String
    ''' <summary>
    ''' Array d'strings el index del qual es el valor retornat per datareader.
    ''' </summary>
    ''' <remarks></remarks>
    Public IndexedValue() As String
    ''' <summary>
    ''' Indica si imprimeix un valor si te el mateix valor que el registre anterior
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintRepeatedValues As Boolean
    ''' <summary>
    ''' Darrer valor impres. Utilitzat per PrintRepeatedValues.
    ''' </summary>
    ''' <remarks></remarks>
    Friend LastValuePrinted As String
    ''' <summary>
    ''' Valor del format .NET a aplicar al camp.
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property FieldFormat() As String
      Get
        Return mFieldFormat
      End Get
      Set(ByVal value As String)
        If String.IsNullOrEmpty(value) Then
          mFieldFormat = "{0}"
        Else
          mFieldFormat = "{0:" + value + "}"
        End If
      End Set
    End Property

    ''' <summary>
    ''' Aplica de forma explicita el format a la columna
    ''' </summary>
    ''' <param name="Format">Format que s'aplica.</param>
    ''' <param name="Formating">Tipus de format al que correspon el Fromat</param>
    ''' <remarks></remarks>
    Public Sub SetFormat(ByVal Format As String, ByVal Formating As FormatingEnum)
      Me.FieldFormating = Formating
      If Formating = FormatingEnum.StringFormat Then
        ' Format de .Net
        Me.FieldFormat = Format
      Else
        ' Format Transform Custom
        Me.mFieldFormat = Format
      End If
    End Sub

  End Class

  ''' <summary>
  ''' Clase per gestionar les agrupacions del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public Class GroupInfo
    ''' <summary>
    ''' Text fixe al titol de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public HeaderCaption As String
    ''' <summary>
    ''' Text fixe al peu de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FooterCaption As String
    ''' <summary>
    ''' Camp sobre el que es fa la agrupació. Cal que el datareader estigui ordenat per aquest camp.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldData As String
    ''' <summary>
    ''' Format a aplicar al camp sobre el que es fa la agrupació.
    ''' </summary>
    ''' <remarks></remarks>
    Public FieldFormat As String
    ''' <summary>
    ''' Valor actual sobre el que s'esta agrupant. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public LastValue As String
    ''' <summary>
    ''' Indica si al començar un nou grup cal fer-ho a un nova pàgina.
    ''' </summary>
    ''' <remarks></remarks>
    Public StartOnNewPage As Boolean
    ''' <summary>
    ''' Estat en el que estoba una agruapció en un moment determinat durant la impressió. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public State As GroupStateEnum
    ''' <summary>
    ''' Col·lecció de TotalInfo de les columnes sobre les que cal fer agregació per a aquesta agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Totals As New System.Collections.Generic.List(Of TotalInfo)
    ''' <summary>
    ''' Indica si imprimeix la capçalera de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintHeaderCaption As Boolean
    ''' <summary>
    ''' Indica si imprimeix el peu de grup
    ''' </summary>
    ''' <remarks></remarks>
    Public PrintFooterCaption As Boolean
    ''' <summary>
    ''' Espai en 1/100" a deixar despres de imprimir el peu de grup. Se afegirà encara que PrintFooterCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceAfterFooter As Integer
    ''' <summary>
    ''' Espai en 1/100" a deixar abans de imprimir la capçalera de grup. Se afegirà encara que PrintHeaderCaption sigui False
    ''' </summary>
    ''' <remarks></remarks>
    Public SpaceBeforeCaption As Integer

    ''' <summary>
    ''' Indica si es produeix un canvi de valor al camp sobre el que es fa la agrupació. Us Intern
    ''' </summary>
    ''' <param name="dr">Datareader sobre el que comprobar si hi ha un canvi d'agruapció</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBreak(ByVal dr As IDataReader) As Boolean
      Return (Me.LastValue <> String.Format("{0}", dr(FieldData)))
    End Function

    ''' <summary>
    ''' Inicialització al fer el canvi de grup. Us intern.
    ''' </summary>
    ''' <param name="Value">Nou valor de la agrupació</param>
    ''' <remarks></remarks>
    Public Sub Init(ByVal Value As String)
      Me.LastValue = Value
    End Sub

    ''' <summary>
    ''' Reseteja els valors del totals al canviar de grup. Us intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
      For Each t As TotalInfo In Totals
        t.Reset()
      Next
    End Sub

    ''' <summary>
    ''' Actualitza els totals generals. Us intern.
    ''' </summary>
    ''' <param name="dr">DataReader que conte els valors dels camp a totalitzar.</param>
    ''' <remarks></remarks>
    Public Sub UpdateTotals(ByVal dr As IDataReader)
      For Each t As TotalInfo In Totals
        Select Case t.Col.TotalColumn
          Case TotalColumnEnum.Count
            t.Total += 1
          Case TotalColumnEnum.Sum
            If Not (IsDBNull(dr(t.Col.FieldName)) OrElse IsNothing(dr(t.Col.FieldName))) Then
              t.Total += CDec(dr(t.Col.FieldName))
            End If
        End Select
      Next
      State = GroupStateEnum.AddingRow
    End Sub

  End Class

  ''' <summary>
  ''' Definició del total de una columna.
  ''' </summary>
  ''' <remarks></remarks>
  Public Class TotalInfo
    ''' <summary>
    ''' ColumnInfo de la columna sobre la que s'aplica un total.
    ''' </summary>
    ''' <remarks></remarks>
    Public Col As ColumnInfo
    ''' <summary>
    ''' Variable en la que es magatzema el total acumulat. Us Intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Total As Decimal
    ''' <summary>
    ''' Font amb el que s'imprimira el total
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalFont As Font
    ''' <summary>
    ''' 'Brush amb el que s'imprimira el total.
    ''' </summary>
    ''' <remarks></remarks>
    Public TotalBrush As Brush

    ''' <summary>
    ''' Inicialitza la variable sobre la que s'acumula el total. Us intern.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Reset()
      Me.Total = 0
    End Sub

  End Class

  ''' <summary>
  ''' Clase empreada per magatzemar informació sobre els criteris de filtre que s'han aplicat per obtenir les dades.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class CriteriaInfo
    ''' <summary>
    ''' Texte identificatiu del filtre
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Valor del filtre.
    ''' </summary>
    ''' <remarks></remarks>
    Public Value As String
  End Class

  ''' <summary>
  ''' Agrupació de capçaleres. Agrupa en un nivell superior diverses capçaleres per donar coherencia a un grup de columnes. Amén.
  ''' </summary>
  ''' <remarks></remarks>
  Protected Class GroupCaptionInfo
    ''' <summary>
    ''' Columna inical de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public FromColumn As Integer
    ''' <summary>
    ''' Columna final de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public ToColumn As Integer
    ''' <summary>
    ''' Texte de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Caption As String
    ''' <summary>
    ''' Alineació de la agrupació
    ''' </summary>
    ''' <remarks></remarks>
    Public Aligment As StringAlignment
    ''' <summary>
    ''' Si cal subratllar la agruapció
    ''' </summary>
    ''' <remarks></remarks>
    Public Underline As Boolean
  End Class


#End Region

#Region " Variables "

  Private ColumnsReport As New System.Collections.Generic.List(Of ColumnInfo)
  Private GroupsReport As New System.Collections.Generic.List(Of GroupInfo)
  Private Totals As New GroupInfo
  Private Criteria As New System.Collections.Generic.List(Of CriteriaInfo)
  Private CaptionGroups As New System.Collections.Generic.List(Of GroupCaptionInfo)
  Private SubGroupLevel As Integer

#Region " Fonts "
  Private DefaultHeaderFont As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultHeaderBrush As Brush = Brushes.Black

  Private DefaultFooterFont As New Font("Arial", 6, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultFooterBrush As Brush = Brushes.Black

  Private DefaultColumnCaptionFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultColumnCaptionBrush As Brush = Brushes.Black

  Private DefaultDetailRowFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultDetailRowBrush As Brush = Brushes.Black

  Private DefaultGroupHeaderFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultGroupHeaderBrush As Brush = Brushes.Black

  Private DefaultGroupFooterFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultGroupFooterBrush As Brush = Brushes.Black

  Private DefaultTotalFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultTotalBrush As Brush = Brushes.Black

  Private PenThick As New Pen(Color.Black, 2)
  Private PenThin As New Pen(Color.Black, 1)

#End Region

  ''' <summary>
  ''' font amb el que s'imprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà el peu de pàgina del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimirà els encolumnas del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les línies de detall
  ''' </summary>
  ''' <remarks></remarks>
  Public RowFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran les capçaleres de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran els peus de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterFont As Font
  ''' <summary>
  ''' Font amb el que s'imprimiran els totals de grup
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalFont As Font

  ''' <summary>
  ''' Brush amb el que simprimirà la capçalera del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public HeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el peu de pàgian del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public FooterBrush As Brush
  ''' <summary>
  ''' Pincell amb el que simprimiran les capçaleres de columna.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnCaptionBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimiran les linies de detall del llistat
  ''' </summary>
  ''' <remarks></remarks>
  Public RowBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà la capçalera de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupHeaderBrush As Brush
  ''' <summary>
  ''' Brush amb el que s'imprimirà el preu de grup.
  ''' </summary>
  ''' <remarks></remarks>
  Public GroupFooterBrush As Brush
  ''' <summary>
  ''' Brush amb el que s
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalBrush As Brush
  ''' <summary>
  ''' Indica si cal imprimir un total general. Cal indicar quenes columnes cal totalitzar i el el tipus de total a aplicar.
  ''' </summary>
  ''' <remarks></remarks>
  Public TeTotalGeneral As Boolean
  ''' <summary>
  ''' Literal del total.
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalGeneralCaption As String
  ''' <summary>
  ''' columna sobre la que s'imprimeix el literal del total. Normalment alineat a la dreta.
  ''' </summary>
  ''' <remarks></remarks>
  Public TotalCaptionColumn As Integer
  ''' <summary>
  ''' Separació en 1/100 de polsada entre les columnes del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public ColumnGap As Integer

  ''' <summary>
  ''' Separació en 1/100 de polsada entre les files del llistat.
  ''' </summary>
  ''' <remarks></remarks>
  Public RowGap As Integer

  ''' <summary>
  ''' Dibija una linea de separació entre línies.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRows As Boolean
  ''' <summary>
  ''' Indica cada cuantes linies impreses cal dibuijar la línea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRowsPeriode As Integer
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada pàgina. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRowsResetOnNewPage As Boolean
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada grup. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Public DrawLineBetweenRowsResetOnNewGroup As Boolean
  ''' <summary>
  ''' Contador intern de linea.
  ''' </summary>
  ''' <remarks></remarks>
  Protected DrawLineBetweenRowsRowCount As Integer
  ''' <summary>
  ''' Brush utilitzar per dibuijar la linea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Protected DrawLineBetweenRowsPen As Pen


  ''' <summary>
  ''' Indica si s'imprimeix linea obbrejada per facilitar la lectura.
  ''' </summary>
  ''' <remarks></remarks>
  Public PaperPijama As Boolean
  ''' <summary>
  ''' Indica cada cuantes linies impreses cal dibuijar la línea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Public PijamaPeriode As Integer
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada pàgina. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Public PijamaResetOnNewPage As Boolean
  ''' <summary>
  ''' Indica si el periode cal reinicairlo a cada grup. Per defecte True.
  ''' </summary>
  ''' <remarks></remarks>
  Public PijamaResetOnNewGroup As Boolean
  ''' <summary>
  ''' Contador intern de linea.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaRowCount As Integer
  ''' <summary>
  ''' Brush utilitzar per dibuijar la linea de pijama.
  ''' </summary>
  ''' <remarks></remarks>
  Protected PijamaBrush As Brush

#End Region

  Public Sub New()

    HeaderFont = DefaultHeaderFont
    FooterFont = DefaultFooterFont
    ColumnCaptionFont = DefaultColumnCaptionFont
    RowFont = DefaultDetailRowFont
    GroupHeaderFont = DefaultGroupHeaderFont
    GroupFooterFont = DefaultGroupFooterFont
    TotalFont = DefaultTotalFont

    HeaderBrush = DefaultHeaderBrush
    FooterBrush = DefaultFooterBrush
    ColumnCaptionBrush = ColumnCaptionBrush
    RowBrush = DefaultDetailRowBrush
    GroupHeaderBrush = DefaultGroupHeaderBrush
    GroupFooterBrush = DefaultGroupFooterBrush
    TotalBrush = DefaultTotalBrush

    ColumnGap = 5
    RowGap = 0
    HeaderKind = HeaderKindEnum.Plain
    LayoutOffset = LayoutOffsetEnum.Centered

    TeTotalGeneral = False
    PaperPijama = False
    PijamaPeriode = 3
    PijamaResetOnNewPage = True
    PijamaResetOnNewGroup = True
    PijamaBrush = New SolidBrush(Color.FromArgb(128, 220, 220, 220))

  End Sub

  ''' <summary>
  ''' Afegir columna al llistat.
  ''' </summary>
  ''' <param name="Caption">Titol de la columna</param>
  ''' <param name="Width">Ample de la columna en 1/100"</param>
  ''' <param name="FieldName">Nom del camp que s'imprimirà a la columna.</param>
  ''' <param name="Aligment">Alineació a aplicar a la columna</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function AddColumn(ByVal Caption As String, ByVal Width As Integer, ByVal FieldName As String, ByVal Aligment As StringAlignment) As ColumnInfo
    Dim Column As New ColumnInfo
    With Column
      .Caption = Caption
      .Width = Width
      .FieldFormat = ""
      .FieldFormating = FormatingEnum.StringFormat
      .FieldName = FieldName
      .FieldNameKind = FieldNameKindEnum.Field
      .Aligment = Aligment
      .ColumnDataKind = ColumnDataKindEnum.Normal
      .DetailRowFont = RowFont
      .DetailRowBrush = RowBrush
      .HeaderFont = HeaderFont
      .HeaderBrush = HeaderBrush
      .PrintRepeatedValues = True
      .LastValuePrinted = String.Empty
    End With
    ColumnsReport.Add(Column)
    Return Column
  End Function

  Public Function AddColumn(ByVal Width As Integer, ByVal Aligment As StringAlignment) As ColumnInfo
    Return AddColumn("", Width, "", Aligment)
  End Function

  Private Function IndexOfColumn(ByVal ColumnField As String) As Integer
    Dim Index As Integer = 0
    For Each c As ColumnInfo In ColumnsReport
      If c.FieldName.ToUpper = ColumnField.ToUpper Then
        Exit For
      End If
      Index += 1
    Next
    If Index = ColumnsReport.Count Then
      Index = -1
    End If
    Return Index
  End Function

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String)
    Dim Group As New GroupInfo
    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = True
      .PrintHeaderCaption = True
      .SpaceAfterFooter = 0
      .SpaceBeforeCaption = 0

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegir una agrupació al llistat. Cal afegirles de manera ordenada.
  ''' </summary>
  ''' <param name="HeaderCaption">Titol del grup</param>
  ''' <param name="FooterCaption">Literal del peu de grup</param>
  ''' <param name="FieldData">Nom del camp sobre el que es fa la agrupació</param>
  ''' <param name="FieldFormat">Format (.NET) a aplicar. </param>
  ''' <param name="PrintHeaderCaption">Indica si imprimeix la capçalera del grup</param>
  ''' <param name="PrintFooterCaption">Indica si imprimeix el peu de grup</param>
  ''' <param name="SpaceBeforeCaption">Espai en 1/100" anabs de imprimir la capçalera. S'afegeig sempre encara que no s'imprimeixi el grup </param>
  ''' <param name="SpaceAfterFooter">Espai en 1/100" despres de imprimir el peu del grup. S'afegeig sempre encara que no s'imprimeixi el grup</param>
  ''' <remarks></remarks>
  Public Sub AddGroup(ByVal HeaderCaption As String, ByVal FooterCaption As String, ByVal FieldData As String, ByVal FieldFormat As String, ByVal PrintHeaderCaption As Boolean, ByVal PrintFooterCaption As Boolean, ByVal SpaceBeforeCaption As Integer, ByVal SpaceAfterFooter As Integer)
    Dim Group As New GroupInfo

    With Group

      .HeaderCaption = HeaderCaption
      .FooterCaption = FooterCaption
      .FieldData = FieldData
      .PrintFooterCaption = PrintFooterCaption
      .PrintHeaderCaption = PrintHeaderCaption
      .SpaceAfterFooter = SpaceAfterFooter
      .SpaceBeforeCaption = SpaceBeforeCaption

      If String.IsNullOrEmpty(FieldFormat) Then
        .FieldFormat = "{0}"
      Else
        .FieldFormat = "{0:" + FieldFormat + "}"
      End If
      Me.GroupsReport.Add(Group)
    End With
  End Sub

  ''' <summary>
  ''' Afegeix un criteri o filtre aplicat a les dades. S'imprimeix al principi del llistat.
  ''' </summary>
  ''' <param name="Caption">Titol del vaolr filtrat.</param>
  ''' <param name="Value">Valor aplicat al filtre.</param>
  ''' <remarks></remarks>
  Public Sub AddCriteria(ByVal Caption As String, ByVal Value As String)
    Dim c As New CriteriaInfo
    c.Caption = Caption
    c.Value = Value
    Criteria.Add(c)
  End Sub

  ''' <summary>
  ''' Afegeig una agrupació de columnes. al imprimir els titols de les columnes.
  ''' </summary>
  ''' <param name="FromColumn">Columna inicial</param>
  ''' <param name="ToColumn">Columan final</param>
  ''' <param name="Caption">Titol de la agrupació</param>
  ''' <param name="Aligment">Alineació a aplicar.</param>
  ''' <param name="Underlined">Subratllat.</param>
  ''' <remarks></remarks>
  Public Sub AddGroupCaption(ByVal FromColumn As Integer, ByVal ToColumn As Integer, ByVal Caption As String, ByVal Aligment As StringAlignment, ByVal Underlined As Boolean)
    Dim gc As New GroupCaptionInfo
    gc.FromColumn = FromColumn
    gc.ToColumn = ToColumn
    gc.Caption = Caption
    gc.Aligment = Aligment
    gc.Underline = Underlined
    CaptionGroups.Add(gc)
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="Value">True si es vol paper pijama.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal Value As Boolean)
    Me.PaperPijama = Value
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="ResetOnNewPage">Reinicialitza el contador de paper pijama al canviar de pàgina.</param>
  ''' <param name="ResetOnNewGroup">Reinicialitza el contador de paper pijama al canviar de grup.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal ResetOnNewPage As Boolean, ByVal ResetOnNewGroup As Boolean)
    Me.PaperPijama = True
    Me.PijamaResetOnNewGroup = ResetOnNewGroup
    Me.PijamaResetOnNewPage = ResetOnNewPage
  End Sub

  ''' <summary>
  ''' Indica si imprimiex paper pijama
  ''' </summary>
  ''' <param name="ResetOnNewPage">Reinicialitza el contador de paper pijama al canviar de pàgina.</param>
  ''' <param name="ResetOnNewGroup">Reinicialitza el contador de paper pijama al canviar de grup.</param>
  ''' <param name="PijamaPeriode">Indica cada cuantes linies es dibuija la linea ombrejada.</param>
  ''' <remarks></remarks>
  Public Sub SetPaperPijama(ByVal ResetOnNewPage As Boolean, ByVal ResetOnNewGroup As Boolean, ByVal PijamaPeriode As Integer)
    Me.PaperPijama = True
    Me.PijamaResetOnNewGroup = ResetOnNewGroup
    Me.PijamaResetOnNewPage = ResetOnNewPage
    Me.PijamaPeriode = PijamaPeriode
  End Sub

  Private Sub GroupsInit()
    For Each g As GroupInfo In GroupsReport
      g.Init(String.Format(g.FieldFormat, (DataSource(g.FieldData))))
      g.State = GroupStateEnum.FooterPrinted
      g.Reset()
    Next
  End Sub

  Private Sub DrawPageHeader(ByVal Canvas As Graphics)

    CurrentPage += 1

    RaiseEvent PaintHeader(Canvas, CurrentPage, TotalPages)

  End Sub

  Public Sub DrawDefaultPageHeader(ByVal Canvas As Graphics)
    Dim sf As New StringFormat

    sf.Alignment = StringAlignment.Far
    Me.CurY = 5

    Me.DrawLine(Canvas, Me.PenThick, 0, CurY, PageWidth, CurY)

    Me.CurY += 2
    Me.DrawString(Canvas, Me.EmpresaName, HeaderFont, HeaderBrush, 0, Me.CurY)
    Me.DrawString(Canvas, String.Format(Me.ReportName, CurrentPage), HeaderFont, HeaderBrush, PageWidth, CurY, sf)
    Me.CurY += CInt(HeaderFont.GetHeight(Canvas))
    Me.DrawLine(Canvas, Me.PenThin, 0, CurY, PageWidth, CurY)
    Me.CurY += 5

  End Sub

  Private Sub DrawPageFooter(ByVal Canvas As Graphics)
    RaiseEvent PaintFooter(Canvas)
  End Sub

  Public Sub DrawDefaultPageFooter(ByVal Canvas As Graphics)
    Dim y As Integer
    Dim sf As New StringFormat
    y = CInt(PageHeight - FooterFont.GetHeight(Canvas)) - 2
    BottomY = y - 5

    Me.DrawLine(Canvas, Me.PenThin, 0, y, PageWidth, y)
    y += 1
    Me.DrawString(Canvas, String.Format("Usuari: {0} LT: {1} FM: {2}", Me.UserName, Me.WorkstationIP, Me.ReportID), FooterFont, HeaderBrush, 0, y)
    sf.Alignment = StringAlignment.Center
    Me.DrawString(Canvas, String.Format("Data: {0:dd/MM/yyyy HH:mm}", Date.Now), FooterFont, HeaderBrush, PageWidth \ 2, y, sf)
    sf.Alignment = StringAlignment.Far
    If PageNumbering = PageNumberEnum.PageN Then
      Me.DrawString(Canvas, String.Format("Pàgina: {0}", CurrentPage), FooterFont, HeaderBrush, PageWidth, y, sf)
    Else
      Me.DrawString(Canvas, String.Format("Pàgina: {0} de {1}", CurrentPage, TotalPages), FooterFont, HeaderBrush, PageWidth, y, sf)
    End If

  End Sub

  Private Sub DrawCriteria(ByVal Canvas As Graphics)
    RaiseEvent PaintCriteria(Canvas)
  End Sub

  Public Sub DrawDefaultCriteria(ByVal Canvas As Graphics)
    If Criteria.Count = 0 Then
      Return
    End If

    Dim height As Integer = 0
    Dim rowHeight As Integer
    Dim MaxCaptionLen As Integer = 0
    Dim MaxValueLen As Integer = 0

    rowHeight = CInt(RowFont.GetHeight(Canvas))

    For Each c As CriteriaInfo In Criteria
      MaxCaptionLen = Math.Max(MaxCaptionLen, CInt(Canvas.MeasureString(c.Caption, RowFont).Width))
      MaxValueLen = Math.Max(MaxValueLen, CInt(Canvas.MeasureString(c.Value, RowFont).Width))
    Next

    CurY += 5

    Me.DrawRectangle(Canvas, Pens.Black, BodyLeft, CurY, 3 + MaxCaptionLen + 5 + MaxValueLen + 3, 3 + rowHeight * Criteria.Count + 3)
    CurY += 3

    For Each c As CriteriaInfo In Criteria
      Me.DrawString(Canvas, c.Caption, RowFont, RowBrush, BodyLeft + 3, CurY)
      Me.DrawString(Canvas, c.Value, RowFont, RowBrush, BodyLeft + 3 + MaxCaptionLen + 5, CurY)
      CurY += rowHeight
    Next

    CurY += 3 ' linea baixa del rectangle

  End Sub

  Private Function DrawGroupHeader(ByVal Canvas As System.Drawing.Graphics, ByVal group As GroupInfo) As Boolean
    Dim GroupHeaderHasBeenPrinted As Boolean
    GroupHeaderHasBeenPrinted = True
    RaiseEvent PaintGroupHeader(Canvas, group, GroupHeaderHasBeenPrinted)
    If GroupHeaderHasBeenPrinted Then
      group.State = GroupStateEnum.HeaderPrinted
    End If
    Return GroupHeaderHasBeenPrinted
  End Function

  Public Function DrawDefaultGroupHeader(ByVal Canvas As System.Drawing.Graphics, ByVal group As GroupInfo) As Boolean
    '
    ' Calcular la alçada del text a imprimir
    Dim alsadaText, alsadaGrup As Integer

    If group.State = GroupStateEnum.HeaderPrinted Then
      Return True
    End If

    If Not group.PrintHeaderCaption Then
      CurY += group.SpaceBeforeCaption
      group.State = GroupStateEnum.HeaderPrinted
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaGrup = 5 + alsadaText + 5 + 3

    If CurY + alsadaGrup >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    CurY += 5


    Me.DrawLine(Canvas, Me.PenThick, BodyLeft, CurY, BodyLeft, CurY + alsadaText + 2)

    Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY + alsadaText + 2, BodyLeft + BodyWidth, CurY + alsadaText + 2)
    Me.DrawString(Canvas, String.Format("{0} {1}", group.HeaderCaption, group.LastValue), Me.GroupHeaderFont, Me.GroupHeaderBrush, BodyLeft + 5, CurY)

    CurY += 5 + alsadaText + 3

    group.State = GroupStateEnum.HeaderPrinted

    Return True

  End Function

  Private Function DrawGroupFooter(ByVal Canvas As Graphics, ByVal group As GroupInfo) As Boolean
    Dim GroupFooterHasBeenPrinted As Boolean
    GroupFooterHasBeenPrinted = True
    RaiseEvent PaintGroupFooter(Canvas, group, GroupFooterHasBeenPrinted)
    If GroupFooterHasBeenPrinted Then
      group.State = GroupStateEnum.FooterPrinted
    End If
    Return GroupFooterHasBeenPrinted
  End Function

  Public Function DrawDefaultGroupFooter(ByVal Canvas As Graphics, ByVal gf As GroupInfo) As Boolean
    Dim alsadaText, alsadaGrup As Integer
    Dim totalValue As Decimal

    If gf.State = GroupStateEnum.FooterPrinted Then
      Return True
    End If

    If Not gf.PrintFooterCaption Then
      CurY += gf.SpaceAfterFooter
      gf.State = GroupStateEnum.FooterPrinted
      Return True
    End If

    If gf.FooterCaption = "-" Then
      ' Nomes imprimeix una linea de separació
      alsadaGrup = 2 + 1 + 2

      If CurY + alsadaGrup >= BottomY Then
        ' Si retorna False es que no hi cabia. 
        ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
        Return False
      End If

      CurY += 2

      Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY, BodyLeft + BodyWidth, CurY)

      CurY += 3

      gf.State = GroupStateEnum.FooterPrinted

      Return True

    End If

    If gf.Totals.Count = 0 Then
      ' No hi han columnes de totals
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaGrup = 2 + 2 + alsadaText + 3

    If CurY + alsadaGrup >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    Dim sf As New StringFormat
    sf.Alignment = StringAlignment.Far

    ' Imprimir les linies de total

    CurY += 2
    For Each t As TotalInfo In gf.Totals
      Me.DrawLine(Canvas, Me.PenThin, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next
    CurY += 2

    ' Imprimir el caption
    Dim pos As Integer
    pos = ColumnsReport(TotalCaptionColumn).PosX + ColumnsReport(TotalCaptionColumn).Width

    Me.DrawString(Canvas, String.Format("{0} {1}:", gf.FooterCaption, gf.LastValue), GroupFooterFont, GroupFooterBrush, pos, CurY, sf)

    For Each t As TotalInfo In gf.Totals
      pos = t.Col.PosX + t.Col.Width
      totalValue = t.Total
      If t.Col.TotalColumn = TotalColumnEnum.Evaluated Then
        RaiseEvent GetSubTotalColumn(t.Col.FieldName, totalValue)
      End If
      Me.DrawString(Canvas, String.Format(t.Col.FieldFormat, totalValue), GroupFooterFont, GroupFooterBrush, pos, CurY, sf)
    Next

    CurY += alsadaText + 3

    gf.State = GroupStateEnum.FooterPrinted

    Return True

  End Function

  Private Function DrawSummary(ByVal Canvas As Graphics) As Boolean
    Dim SummaryHasBeenPrinted As Boolean
    SummaryHasBeenPrinted = True
    RaiseEvent PaintSummary(Canvas, Totals, SummaryHasBeenPrinted)
    If SummaryHasBeenPrinted Then
      Totals.Reset()
      Totals.State = GroupStateEnum.FooterPrinted
    End If
    Return SummaryHasBeenPrinted
  End Function

  Public Function DrawDefaultSummary(ByVal Canvas As Graphics) As Boolean
    Dim alsadaSummary As Integer
    Dim alsadaText As Integer
    Dim totalValue As Decimal

    If Totals.Totals.Count = 0 Then
      ' No hi han columnes de totals
      Return True
    End If

    If Totals.State = GroupStateEnum.FooterPrinted Then
      Return True
    End If

    alsadaText = CInt(GroupHeaderFont.GetHeight(Canvas))
    alsadaSummary = 2 + 2 + alsadaText + 2

    If CurY + alsadaSummary >= BottomY Then
      ' Si retorna False es que no hi cabia. 
      ' Es reposnsabilitat de qui crida de generar pagina nova i tornar a cridar la funció.
      Return False
    End If

    Dim sf As New StringFormat
    sf.Alignment = StringAlignment.Far

    ' Imprimir les linies de total

    CurY += 2
    For Each t As TotalInfo In Totals.Totals
      Me.DrawLine(Canvas, Me.PenThin, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next
    CurY += 2

    ' Imprimir el caption
    Dim pos As Integer
    pos = ColumnsReport(TotalCaptionColumn).PosX + ColumnsReport(TotalCaptionColumn).Width
    Me.DrawString(Canvas, TotalGeneralCaption, TotalFont, TotalBrush, pos, CurY, sf)

    For Each t As TotalInfo In Totals.Totals
      pos = t.Col.PosX + t.Col.Width
      totalValue = t.Total
      If t.Col.TotalColumn = Me.TotalColumnEnum.Evaluated Then
        RaiseEvent GetTotalColumn(t.Col.FieldName, totalValue)
      End If
      Me.DrawString(Canvas, String.Format(t.Col.FieldFormat, totalValue), TotalFont, TotalBrush, pos, CurY, sf)
    Next

    CurY += alsadaText + 2
    For Each t As TotalInfo In Totals.Totals
      Me.DrawLine(Canvas, Me.PenThick, t.Col.PosX, CurY, t.Col.PosX + t.Col.Width, CurY)
    Next

    Totals.Reset()
    Totals.State = GroupStateEnum.FooterPrinted

    Return True

  End Function

  Private Function DrawColumnCaptions(ByVal Canvas As Graphics) As Boolean
    Dim ColumnCaptionsHaveBeenPrinted As Boolean
    ColumnCaptionsHaveBeenPrinted = True
    RaiseEvent PaintColumnCaptions(Canvas, ColumnCaptionsHaveBeenPrinted)
    Return ColumnCaptionsHaveBeenPrinted
  End Function

  Public Sub DrawDefaultColumnCaptions(ByVal Canvas As Graphics)

    Dim sf As New StringFormat
    Dim rh As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim layoutRec As RectangleF
    Dim MaxH As Integer

    CurY += 20

    If CaptionGroups.Count > 0 Then

      For Each gc As GroupCaptionInfo In CaptionGroups

        Dim FromX As Integer
        Dim Width As Integer

        FromX = ColumnsReport(gc.FromColumn).PosX
        Width = ColumnsReport(gc.ToColumn).PosX + ColumnsReport(gc.ToColumn).Width - FromX

        layoutRec = New RectangleF(FromX, CurY, Width, rh)
        sf.Alignment = gc.Aligment
        Me.DrawString(Canvas, gc.Caption, RowFont, RowBrush, layoutRec, sf)

        Me.DrawLine(Canvas, Me.PenThin, FromX, CurY + rh + 2, FromX + Width, CurY + rh + 2)

      Next

      CurY += rh + 2

    End If

    For Each c As ColumnInfo In ColumnsReport

      sf.Alignment = c.Aligment

      Select Case c.ColumnDataKind
        Case ColumnDataKindEnum.BarCode, ColumnDataKindEnum.BoxToWriteIn, ColumnDataKindEnum.CheckBox, ColumnDataKindEnum.IsBoolean, ColumnDataKindEnum.IsImage, ColumnDataKindEnum.MultipleLines, ColumnDataKindEnum.Normal, ColumnDataKindEnum.EventValue
          layoutRec = New RectangleF(c.PosX, CurY, c.Width, rh)
          Me.DrawString(Canvas, c.Caption, RowFont, RowBrush, layoutRec, sf)
          MaxH = Math.Max(MaxH, rh)
        Case ColumnDataKindEnum.FormLayout
          ' Nothing
          MaxH = 0
      End Select

    Next

    CurY += MaxH + 2

    For Each c As ColumnInfo In ColumnsReport

      Me.DrawLine(Canvas, Me.PenThin, c.PosX, CurY, c.PosX + c.Width, CurY)

    Next

    CurY += 2

  End Sub

  Private Function DrawColumn(ByVal Canvas As Graphics, ByVal col As ColumnInfo, ByVal Value As String) As Integer
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim rowHeight As Integer = height

    Value = Value.Trim

    Select Case col.ColumnDataKind
      Case ColumnDataKindEnum.IsBoolean
        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case col.Aligment
          Case StringAlignment.Center
            x = col.PosX + col.Width \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = col.PosX + col.Width - height + offset * 2
          Case StringAlignment.Near
            x = col.PosX
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

        If Value.ToLower = "true" OrElse Value = "1" Then
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + offset * 2, x + costat - offset, CurY + costat)
          Me.DrawLine(Canvas, Pens.Black, x + offset, CurY + costat, x + costat - offset, CurY + offset * 2)
        End If

      Case ColumnDataKindEnum.MultipleLines
        Dim alsadaText As Integer
        alsadaText = CInt(Canvas.MeasureString(Value, RowFont, col.Width).Height)

        Dim sf As New StringFormat
        sf.Alignment = col.Aligment
        sf.FormatFlags = StringFormatFlags.LineLimit
        sf.Trimming = StringTrimming.Word

        If alsadaText > Me.BottomY - CurY Then
          ' no hi cap
          Dim charsIn, linFilled As Integer

          Canvas.MeasureString(Value, RowFont, New Size(col.Width, Me.BottomY - CurY), sf, charsIn, linFilled)
          col.TextLeft = Value.Substring(charsIn)

          Me.DrawString(Canvas, Value.Substring(0, charsIn), RowFont, RowBrush, col.PosX, CurY)

          rowHeight = Me.BottomY

        Else

          Dim r As New Rectangle(col.PosX, CurY, col.Width, alsadaText)
          Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)

          rowHeight = alsadaText

        End If

      Case ColumnDataKindEnum.Normal, ColumnDataKindEnum.IndexedValue, ColumnDataKindEnum.EventValue

        Dim sf As New StringFormat
        sf.Alignment = col.Aligment
        sf.Trimming = StringTrimming.EllipsisCharacter
        Dim r As New Rectangle(col.PosX, CurY, col.Width, height)

        If col.PrintRepeatedValues Then
          Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)
        Else
          If Value <> col.LastValuePrinted Then
            Me.DrawString(Canvas, Value, RowFont, RowBrush, r, sf)
            col.LastValuePrinted = Value
          End If
        End If


      Case ColumnDataKindEnum.CheckBox

        Dim offset As Integer = height \ 10
        Dim costat As Integer = height - offset * 2

        Dim x As Integer
        Select Case col.Aligment
          Case StringAlignment.Center
            x = col.PosX + col.Width \ 2 - height \ 2 + offset
          Case StringAlignment.Far
            x = col.PosX + col.Width - height + offset * 2
          Case StringAlignment.Near
            x = col.PosX
        End Select

        Dim Cuadre As PointF() = {New PointF(x, CurY + offset), _
          New PointF(x + costat, CurY + offset), _
          New PointF(x + costat, CurY + costat + offset), _
          New PointF(x, CurY + offset + costat), _
          New PointF(x, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Cuadre)

      Case ColumnDataKindEnum.BoxToWriteIn

        Dim offset As Integer = height \ 4

        Dim Calaix As PointF() = {New PointF(col.PosX, CurY + offset), _
          New PointF(col.PosX, CurY + height), _
          New PointF(col.PosX + col.Width, CurY + height), _
          New PointF(col.PosX + col.Width, CurY + offset)}

        Me.DrawLines(Canvas, Pens.Black, Calaix)

      Case ColumnDataKindEnum.IsImage
        ' el camp data del DataReader es el nom del fitxer si esta en un directori este esta al PathImages
        Dim Filename As String
        Dim img As Image
        Filename = IO.Path.Combine(col.PathImages, Value.Trim)
        If IO.File.Exists(Filename) Then
          img = Image.FromFile(Filename)
          Dim imgH As Integer
          Dim imgW As Integer
          Dim Ratio As Double
          imgH = CInt(img.Height \ CInt(img.VerticalResolution) \ 100)
          imgW = CInt(img.Width \ CInt(img.HorizontalResolution) \ 100)
          ' Faig cuadrar l'alsada
          Ratio = col.ImageHeight / imgH
          imgH = CInt(CDbl(imgH) / Ratio)
          imgW = CInt(CDbl(imgW) / Ratio)
          If imgW > col.Width Then
            'calcular el ratio de reduccio
            Ratio = col.Width / imgW
            imgH = CInt(CDbl(imgH) / Ratio)
            imgW = CInt(CDbl(imgW) / Ratio)
          End If

          ' Calculem la situació Horitzontal
          Dim imgPosX As Integer

          Select Case col.Aligment
            Case StringAlignment.Center
              imgPosX = col.PosX + (col.Width - imgW) \ 2
            Case StringAlignment.Far
              imgPosX = col.PosX + col.Width - imgW
            Case StringAlignment.Near
              imgPosX = col.PosX
          End Select

          Me.DrawImage(Canvas, img, New RectangleF(imgPosX, CurY, imgW, imgH), New RectangleF(0, 0, img.Width, img.Height), GraphicsUnit.Pixel)

          img.Dispose()

        End If

        rowHeight = col.ImageHeight

      Case ColumnDataKindEnum.BarCode

        'Dim bc As New csBarcode.csBarCode
        'bc.Data = Value
        'bc.Symbology = col.BarCodeSymbol
        'bc.DrawReadableData = col.BarCodeDrawData

        ''bc.DrawBarcode(Canvas, CurY, col.PosX, col.Width, col.ImageHeight)

        'bc = Nothing
        'rowHeight = col.ImageHeight

    End Select

    Return rowHeight

  End Function

  Private Function DrawColumnLeft(ByVal Canvas As Graphics, ByVal col As ColumnInfo) As Integer
    Dim height As Integer = CInt(RowFont.GetHeight(Canvas))
    Dim rowHeight As Integer = height

    Dim alsadaText As Integer
    alsadaText = CInt(Canvas.MeasureString(col.TextLeft, RowFont, col.Width).Height)

    Dim sf As New StringFormat
    sf.Alignment = col.Aligment
    sf.FormatFlags = StringFormatFlags.LineLimit
    sf.Trimming = StringTrimming.Word

    If alsadaText > Me.BottomY - CurY Then
      ' no hi cap
      Dim charsIn, linFilled As Integer

      Canvas.MeasureString(col.TextLeft, RowFont, New Size(col.Width, Me.BottomY - CurY), sf, charsIn, linFilled)
      col.TextLeft = col.TextLeft.Substring(charsIn)

      Me.DrawString(Canvas, col.TextLeft.Substring(0, charsIn), RowFont, RowBrush, col.PosX, CurY)

      rowHeight = Me.BottomY

    Else

      Dim r As New Rectangle(col.PosX, CurY, col.Width, alsadaText)
      Me.DrawString(Canvas, col.TextLeft, RowFont, RowBrush, r, sf)

      col.TextLeft = String.Empty
      rowHeight = alsadaText

    End If

    Return rowHeight

  End Function

  Private Function DrawRow(ByVal Canvas As Graphics) As Boolean
    Dim RowHasBeenPrinted As Boolean
    RowHasBeenPrinted = True
    RaiseEvent PaintRow(Canvas, DataSource, RowHasBeenPrinted)
    Return RowHasBeenPrinted
  End Function

  Public Function DrawDefaultRow(ByVal Canvas As Graphics) As Boolean
    Dim rowHeight, maxRowHeight As Integer
    Dim Value As String

    If CurY + RowFont.GetHeight(Canvas) + RowGap > BottomY Then
      ' no hi ha espai. 
      Return False
    End If

    maxRowHeight = 0

    If Me.PaperPijama Then
      If Me.PijamaRowCount Mod Me.PijamaPeriode = 0 Then
        Me.FillRectangle(Canvas, PijamaBrush, BodyLeft, CurY, BodyWidth, RowFont.GetHeight(Canvas))
      End If
      Me.PijamaRowCount += 1
    End If

    For Each c As ColumnInfo In ColumnsReport

      Select Case c.ColumnDataKind
        Case _
          ColumnDataKindEnum.BarCode, _
          ColumnDataKindEnum.IsBoolean, _
          ColumnDataKindEnum.Normal, _
          ColumnDataKindEnum.MultipleLines, _
          ColumnDataKindEnum.IsImage
          If c.FieldNameKind = FieldNameKindEnum.Field Then
            If c.FieldFormating = FormatingEnum.StringFormat Then
              Value = String.Format(c.FieldFormat, DataSource(c.FieldName))
            Else
              ' Custom
              Value = Utils.Transform(String.Format("{0}", DataSource(c.FieldName)), c.FieldFormat)
            End If
          Else
            Value = c.FieldName
          End If
        Case ColumnDataKindEnum.IndexedValue
          Try
            Value = c.IndexedValue(CInt(DataSource(c.FieldName)))
          Catch ex As Exception
            Value = String.Format(c.FieldFormat, DataSource(c.FieldName))
          End Try
        Case ColumnDataKindEnum.EventValue
          RaiseEvent EventValue(DataSource, c, Value)
      End Select

      rowHeight = DrawColumn(Canvas, c, Value)
      maxRowHeight = Math.Max(maxRowHeight, rowHeight)

    Next

    CurY += maxRowHeight

    If Me.DrawLineBetweenRows Then
      If CurY + 2 < BottomY Then
        Me.DrawLineBetweenRowsRowCount += 1
        If Me.DrawLineBetweenRowsRowCount Mod Me.DrawLineBetweenRowsPeriode = 0 Then
          CurY += 2
          Me.DrawLine(Canvas, Me.PenThin, BodyLeft, CurY, BodyLeft + BodyWidth, CurY)
        End If
      End If
    End If

    Return True

  End Function

  Public Function TestGroupBreak(ByVal Canvas As Graphics) As Boolean

    If GroupsReport.Count = 0 Then
      Return True
    End If

    For i As Integer = 0 To GroupsReport.Count - 1

      If GroupsReport(i).LastValue <> String.Format(GroupsReport(i).FieldFormat, DataSource(GroupsReport(i).FieldData)) Then

        If Me.PaperPijama Then
          If Me.PijamaResetOnNewGroup Then
            Me.PijamaRowCount = 0
          End If
        End If

        For j As Integer = GroupsReport.Count - 1 To i Step -1
          SubGroupLevel = j
          If GroupsReport(j).State <> GroupStateEnum.FooterPrinted Then
            If Not DrawGroupFooter(Canvas, GroupsReport(j)) Then
              Return False
            End If
            GroupsReport(j).Reset()
            GroupsReport(j).Init(String.Format(GroupsReport(j).FieldFormat, DataSource(GroupsReport(j).FieldData)))
            GroupsReport(j).State = GroupStateEnum.FooterPrinted
          End If
        Next

        For j As Integer = i To GroupsReport.Count - 1
          If GroupsReport(j).State <> GroupStateEnum.HeaderPrinted Then
            If Not DrawGroupHeader(Canvas, GroupsReport(j)) Then
              Return False
            End If
          End If
        Next

        Exit For

      End If

    Next

    Return True

  End Function

  Protected Sub InitLayout(ByVal Canvas As System.Drawing.Graphics)

    If LayoutInitialized Then
      Return
    End If

    LayoutInitialized = True

    'Calcul de posicions de les(columnes)
    ' Ample total de les columnes
    Dim AmpleColumnes As Integer = 0

    For Each c As ColumnInfo In ColumnsReport
      AmpleColumnes += c.Width
    Next

    AmpleColumnes += (ColumnsReport.Count - 1) * ColumnGap

    Select Case LayoutOffset
      Case LayoutOffsetEnum.Centered
        LeftOffset = (PageWidth - AmpleColumnes) \ 2
      Case LayoutOffsetEnum.OneThird
        LeftOffset = (PageWidth - AmpleColumnes) \ 3
      Case LayoutOffsetEnum.Custom
        ' Nothing to do
    End Select

    BodyLeft = LeftOffset
    BodyWidth = AmpleColumnes

    For Each c As ColumnInfo In ColumnsReport
      c.PosX = LeftOffset
      LeftOffset += ColumnGap + c.Width
    Next

    ' Configurar totals
    For Each c As ColumnInfo In ColumnsReport
      If c.TotalColumn > TotalColumnEnum.None Then
        For Each g As GroupInfo In GroupsReport
          Dim t As New TotalInfo
          t.Col = c
          t.Reset()
          g.Totals.Add(t)
        Next

        If TeTotalGeneral Then
          Dim tg As New TotalInfo
          tg.Col = c
          tg.Reset()
          Totals.Totals.Add(tg)
        End If

      End If
    Next

  End Sub

  ''' <summary>
  ''' Imprimeix una pàgina del llistat. Us intern.
  ''' </summary>
  ''' <param name="Canvas">Objecte graphcs sobre el que es dibuixa la pàgina.</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  ''' 

  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim ProcessRow As Boolean

    If FirstPassReport Then

      If DataNeeded Then
        DataNeeded = False
        Do While True
          If Not DataSource.Read Then
            Return False
          End If
          If Not FilterRowOut(DataSource) Then
            Exit Do
          End If
        Loop
      End If

      PageHeight = CInt(Canvas.VisibleClipBounds.Height)
      PageWidth = CInt(Canvas.VisibleClipBounds.Width)

      InitLayout(Canvas)
      GroupsInit()
      Totals.Reset()

    End If

    DrawPageHeader(Canvas)
    DrawPageFooter(Canvas)

    If FirstPassReport Then
      DrawCriteria(Canvas)
      FirstPassReport = False
    End If

    DrawColumnCaptions(Canvas)

    For Each g As GroupInfo In GroupsReport
      If g.State = GroupStateEnum.HeaderPrinted Then
        Continue For
      End If
      If Not DrawGroupHeader(Canvas, g) Then
        ' Ull. que es repeteix el FirstPass
        Return True
      End If
    Next

    If Me.PaperPijama Then
      If Me.PijamaResetOnNewPage Then
        Me.PijamaRowCount = 0
      End If
    End If

    If Me.DrawLineBetweenRows Then
      If Me.DrawLineBetweenRowsResetOnNewPage Then
        Me.DrawLineBetweenRowsRowCount = 0
      End If
    End If

    For Each col As ColumnInfo In ColumnsReport
      If Not col.PrintRepeatedValues Then
        col.LastValuePrinted = String.Empty
      End If
    Next

    Do While True

      If DrawingTotalsAndExit Then
        Exit Do
      End If

      If DataNeeded Then
        DataNeeded = False
        If Not DataSource.Read Then
          Exit Do
        End If
        If FilterRowOut(DataSource) Then
          DataNeeded = True
          Continue Do
        End If
      End If

      If Not TestGroupBreak(Canvas) Then
        Return True
      End If

      ProcessRow = True
      RaiseEvent UpdateValuesBeforeRowPrinted(DataSource, ProcessRow)

      If Not ProcessRow Then
        DataNeeded = True
        Continue Do
      End If

      If mStopReport Then
        DrawString(Canvas, "Llistat interrumput ...", New Font("Arial", 10, FontStyle.Bold), Brushes.Red, 0, CurY + 3)
        Return False
      End If

      If Not DrawRow(Canvas) Then
        Return True
      End If

      RaiseEvent UpdateValuesAfterRowPrinted(DataSource)

      If mStopReport Then
        DrawString(Canvas, "Llistat interrumput ...", New Font("Arial", 10, FontStyle.Bold), Brushes.Red, 0, CurY + 3)
        Return False
      End If

      ' Actualitza subtotals
      For i As Integer = 0 To GroupsReport.Count - 1
        GroupsReport(i).UpdateTotals(DataSource)
      Next
      'Actualitza Total General
      Totals.UpdateTotals(DataSource)

      DataNeeded = True

      CurY += RowGap

    Loop

    DrawingTotalsAndExit = True

    For i As Integer = GroupsReport.Count - 1 To 0 Step -1
      If Not DrawGroupFooter(Canvas, GroupsReport(i)) Then
        Return True
      End If
    Next

    If Not DrawSummary(Canvas) Then
      Return True
    End If

    If False Then
      ' test for more datasources
      ' potser llançar un nou event que redefineixi les columnes i el datareader.
      ' i posar el parar compte que la pàgina pot estar a mitat

      ' Tenir en compte que per al càcul de pàgines cal tornar a deixar-ho en posició inicial de columnes i data source

      ' potser millor encadenar els llistats ...

    End If

    Return False

  End Function

  ''' <summary>
  ''' Inicialitza el llistat. Us intern.
  ''' </summary>
  ''' <remarks></remarks>
  ''' 

  Public Overrides Sub BeginPrint()

    FirstPassReport = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False
    mStopReport = False
    PijamaRowCount = 0
    DrawLineBetweenRowsRowCount = 0

    RaiseEvent InitializeReportValues()

  End Sub

  Protected Overrides Sub Finalize()
    DefaultHeaderFont.Dispose()
    DefaultFooterFont.Dispose()
    DefaultColumnCaptionFont.Dispose()
    DefaultDetailRowFont.Dispose()
    DefaultGroupHeaderFont.Dispose()
    DefaultGroupFooterFont.Dispose()
    DefaultTotalFont.Dispose()
    PijamaBrush.Dispose()
    MyBase.Finalize()
  End Sub

  Public Function GetSubTotalValue(ByVal ColumnFieldName As String) As Decimal
    Dim Value As Decimal
    For Each t As TotalInfo In GroupsReport(SubGroupLevel).Totals
      If t.Col.FieldName.ToLower = ColumnFieldName.ToLower Then
        Value = t.Total
        Exit For
      End If
    Next
    Return Value

  End Function

  Public Function GetTotalValue(ByVal ColumnFieldName As String) As Decimal
    Dim Value As Decimal
    For Each t As TotalInfo In Me.Totals.Totals
      If t.Col.FieldName.ToLower = ColumnFieldName.ToLower Then
        Value = t.Total
        Exit For
      End If
    Next
    Return Value
  End Function

  Protected Overrides Sub Print2Excel(ByVal FileName As String)
    Dim Handled As Boolean
    Handled = False
    RaiseEvent Export2Excel(FileName, Handled)
    If Not Handled Then
      DefaultPrint2Excel(FileName)
    End If
  End Sub

  Public Sub DefaultPrint2Excel(ByVal FileName As String)
    Dim xls As New C1.C1Excel.C1XLBook()
    Dim sheet As C1.C1Excel.XLSheet = xls.Sheets("Sheet1")
    Dim colCount As Integer
    Dim rowCount As Integer
    Dim value As String
    colCount = 0

    For Each c As ColumnInfo In ColumnsReport
      sheet(0, colCount).Value = c.Caption
      colCount += 1
    Next
    rowCount = 0
    Do While DataSource.Read
      If FilterRowOut(DataSource) Then
        Continue Do
      End If
      rowCount += 1
      colCount = 0
      For Each c As ColumnInfo In ColumnsReport
        Select Case c.ColumnDataKind
          Case _
            ColumnDataKindEnum.BarCode, _
            ColumnDataKindEnum.IsBoolean, _
            ColumnDataKindEnum.Normal, _
            ColumnDataKindEnum.MultipleLines, _
            ColumnDataKindEnum.IsImage
            If c.FieldNameKind = FieldNameKindEnum.Field Then
              If c.FieldFormating = FormatingEnum.StringFormat Then
                value = String.Format(c.FieldFormat, DataSource(c.FieldName))
              Else
                ' Custom
                value = Utils.Transform(String.Format("{0}", DataSource(c.FieldName)), c.FieldFormat)
              End If
            Else
              value = c.FieldName
            End If

            Select Case DataSource.GetDataTypeName(DataSource.GetOrdinal(c.FieldName)).ToLower
              Case "int"
                If Not c.FieldName.ToUpper.EndsWith("ID") Then
                  If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                    sheet(rowCount, colCount).Value = CInt(DataSource(c.FieldName))
                  Else
                    sheet(rowCount, colCount).Value = 0
                  End If
                Else
                  sheet(rowCount, colCount).Value = value
                End If
              Case "decimal"
                If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                  sheet(rowCount, colCount).Value = CDec(DataSource(c.FieldName))
                Else
                  sheet(rowCount, colCount).Value = 0D
                End If
              Case "datetime"
                If Not Utils.IsNullOrEmptyValue(DataSource(c.FieldName)) Then
                  sheet(rowCount, colCount).Value = CStr(Date.Parse(String.Format("{0}", DataSource(c.FieldName))))
                Else
                  sheet(rowCount, colCount).Value = value
                End If
              Case Else
                sheet(rowCount, colCount).Value = value
            End Select

          Case _
            ColumnDataKindEnum.MultipleFields, _
            ColumnDataKindEnum.FormLayout

            value = "Format N.V."
        End Select

        colCount += 1

      Next

    Loop

    DataSource.Close()

    xls.Save(FileName)
  End Sub

End Class

Public Class csGeneralRpt
  Inherits csRpt

  Public Event InitializeReportValues()
  Public Event PaintPage(ByVal Canvas As Graphics, ByRef HasMorePages As Boolean, ByVal CurrentPage As Single, ByVal TotalPages As Single)
  Public Event Export2Excel(ByVal FileName As String, ByRef Handled As Boolean)

#Region " Fonts "
  Private DefaultHeaderFont As New Font("Arial", 10, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultHeaderBrush As Brush = Brushes.Black

  Private DefaultFooterFont As New Font("Arial", 6, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultFooterBrush As Brush = Brushes.Black

  Private DefaultColumnCaptionFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultColumnCaptionBrush As Brush = Brushes.Black

  Private DefaultDetailRowFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultDetailRowBrush As Brush = Brushes.Black

  Private DefaultGroupHeaderFont As New Font("Arial", 10, FontStyle.Bold, GraphicsUnit.Point)
  Private DefaultGroupHeaderBrush As Brush = Brushes.Black

  Private DefaultGroupFooterFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultGroupFooterBrush As Brush = Brushes.Black

  Private DefaultTotalFont As New Font("Arial Narrow", 9, FontStyle.Regular, GraphicsUnit.Point)
  Private DefaultTotalBrush As Brush = Brushes.Black

  Private PenThick As New Pen(Color.Black, 2)
  Private PenThin As New Pen(Color.Black, 1)

#End Region

  Public Overrides Sub BeginPrint()
    FirstPassReport = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False
    mStopReport = False

    RaiseEvent InitializeReportValues()
  End Sub

  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim HasMorePages As Boolean
    RaiseEvent PaintPage(Canvas, HasMorePages, CurrentPage, TotalPages)
    Return HasMorePages
  End Function

  Protected Overrides Sub Print2Excel(ByVal FileName As String)
    Dim Handled As Boolean
    Handled = False
    RaiseEvent Export2Excel(FileName, Handled)
    If Not Handled Then
      ' Do something
    End If
  End Sub

  Protected Overrides Sub Finalize()
    DefaultHeaderFont.Dispose()
    DefaultFooterFont.Dispose()
    DefaultColumnCaptionFont.Dispose()
    DefaultDetailRowFont.Dispose()
    DefaultGroupHeaderFont.Dispose()
    DefaultGroupFooterFont.Dispose()
    DefaultTotalFont.Dispose()
    MyBase.Finalize()
  End Sub

End Class
