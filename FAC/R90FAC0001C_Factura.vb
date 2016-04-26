Imports csAppData
Imports csUtils
Imports csUtils.Utils
Imports System.Windows
Imports System.Drawing

Public Class R90FAC0001C_Factura
  Inherits csRpt

  Private FirstPage As Boolean
  Private mIdiomaID As Integer
  Private mCustomID As String
  Private Logo As Image
  Private LogoDimmed As Image
  Private Inscrita As String

  Private TraceLocation As String

  Private dbaCap As New C00_gst_fac
  Private dbaLin As New C00_gst_fal
  Private dbaReb As New C00_gst_reb
  ' Private dbaSys As New C00_gst_sys

  Private cap As DataRow
  Private cli As DataRow
  Private lin As csTableReader
  Private reb As DataTable

  Private ColumnWidth() As Integer
  Private TotalsWidth() As Integer
  Private DetailColumns As Integer
  Private TotalsColumns As Integer
  Private DataLoaded As Boolean
  Private sfCenter As StringFormat
  Private sfNear As New StringFormat
  Private sfFar As New StringFormat
  Private sfTotalBox As New StringFormat
  Private TopYLines As Integer
  Private BottomYLines As Integer
  Private WidthLines As Integer
  Private marginBox As Integer
  Private BodyLeftOffset As Integer

  Private TopYVenciments As Integer
  Private TopYTotals As Integer
  Private TopYTotal As Integer
  Private TopYPagament As Integer
  Private TopYComentaris As Integer
  Private widthVto As Integer
  Private TotalFacturaBoxWidth As Integer

  Private TeDteComercial As Boolean
  Private TipusSerie As Integer
  Private mCanviDivisa As Decimal
  Private mFormatValor As String
 
  Private fntTitle As Font
  Private fntValue As Font
  Private fntHdrValue As Font
  Private fntLine As Font
  Private fntItalic As Font
  Private fntAnagrama As Font
  Private fntInscrita As Font

  Private clrDarkAzure As Color
  Private clrLightAzure As Color

  Private brshDarkAzure As SolidBrush

  Private deltaLine As Integer

  Private TextLeft As String = String.Empty
  Private AmpliacioDescripcio As String = String.Empty
  Private fntLeft As Font
  Private OffsetLeft As Integer
  Private OffsetTop As Integer
  Private ColumnLeft As Integer
  Private LinesLeft As Boolean

  Private LastExpedientID As String

  Private NomClientCustom As String

  Private Enum NowPrintingEnum
    LinFactura
    TextClient
    Summary
  End Enum

  Private NowPrinting As NowPrintingEnum

  Public Shadows Property CustomID() As String
    Get
      Return mCustomID
    End Get
    Set(ByVal value As String)
      If mCustomID <> value Then

        mCustomID = value

        Try
          
          NomClientCustom = "VIESTENDALS"
          Inscrita = "Inscrita en el Reg. Merc. Tarragona, Tomo 1536, Sección de Sociedades, Folio 108, Hoja nº T-19613 inscripción primera. NIF: B-43594092"
          Logo = System.Drawing.Bitmap.FromStream(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("csDosReportEngine.viestendals_normal.png"))
          LogoDimmed = System.Drawing.Bitmap.FromStream(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("csDosReportEngine.viestendals_transparent.png"))

        Catch ex As Exception
          If AppData.Debug Then
            MessageBox.Show("Logo no trobat al Resources")
          End If
          Logo = Nothing
        End Try

      End If

    End Set
  End Property

#Region " Properties "

#Region " Propietat factura "

  Private mOrigenDades As String
  Public Property OrigenDades() As String
    Get
      Return mOrigenDades
    End Get
    Set(ByVal value As String)
      mOrigenDades = value
    End Set
  End Property

  Private mSerieFactura As String
  Public Property SerieFactura() As String
    Get
      Return mSerieFactura
    End Get
    Set(ByVal value As String)
      mSerieFactura = value
    End Set
  End Property

  Private mNumeroFactura As Integer
  Public Property NumeroFactura() As Integer
    Get
      Return mNumeroFactura
    End Get
    Set(ByVal value As Integer)
      mNumeroFactura = value
    End Set
  End Property

  Private mAnyFactura As String
  Public Property AnyFactura() As String
    Get
      Return mAnyFactura
    End Get
    Set(ByVal value As String)
      mAnyFactura = value
    End Set
  End Property

#End Region

#Region " Propietats llistat "

  Private mDestinacio As csRpt.ReportDestinationEnum
  Public Property Destinacio() As csRpt.ReportDestinationEnum
    Get
      Return mDestinacio
    End Get
    Set(ByVal value As csRpt.ReportDestinationEnum)
      Destination = value
      mDestinacio = value
    End Set
  End Property

  Private mCopies As Integer
  Public Property Copies() As Integer
    Get
      Return mCopies
    End Get
    Set(ByVal value As Integer)
      mCopies = value
    End Set
  End Property

  Private mFitxerPdf As String
  Public Property FitxerPdf() As String
    Get
      Return mFitxerPdf
    End Get
    Set(ByVal value As String)
      mFitxerPdf = value
    End Set
  End Property

#End Region

#Region " Propietats correo "

  Private mMailFeedback As String
  Public Property MailFeedback() As String
    Get
      Return mMailFeedback
    End Get
    Set(ByVal value As String)
      mMailFeedback = value
      MyBase.fmMailFeedBack = value
    End Set
  End Property

  Private mMailFrom As String
  Public Property MailFrom() As String
    Get
      Return mMailFrom
    End Get
    Set(ByVal value As String)
      mMailFrom = value
      MyBase.fmMailFrom = value
    End Set
  End Property

  Private mMailReplyTo As String
  Public Property MailReplyTo() As String
    Get
      Return mMailReplyTo
    End Get
    Set(ByVal value As String)
      mMailReplyTo = value
      MyBase.fmMailReplyTo = value
    End Set
  End Property

  Private mMailTo As String
  Public Property MailTo() As String
    Get
      Return mMailTo
    End Get
    Set(ByVal value As String)
      mMailTo = value.Trim
      MyBase.fmMailTo = value.Trim
    End Set
  End Property

  Private mMailSentOK As Boolean
  Public Property MailSentOK() As Boolean
    Get
      Return mMailSentOK
    End Get
    Set(ByVal value As Boolean)
      mMailSentOK = value
    End Set
  End Property

  Private mSmtpLogin As String
  Public Property SmtpLogin() As String
    Get
      Return mSmtpLogin
    End Get
    Set(ByVal value As String)
      mSmtpLogin = value
      MyBase.fmSmtpLogin = value
    End Set
  End Property

  Private mSmtpPassword As String
  Public Property SmtpPassword() As String
    Get
      Return mSmtpPassword
    End Get
    Set(ByVal value As String)
      mSmtpPassword = value
      MyBase.fmSmtpPassword = value
    End Set
  End Property

  Private mSmtpServer As String
  Public Property SmtpServer() As String
    Get
      Return mSmtpServer
    End Get
    Set(ByVal value As String)
      mSmtpServer = value
      MyBase.fmSmtpServer = value
    End Set
  End Property

  Private mShowForm As Boolean
  Public Property ShowForm() As Boolean
    Get
      Return mShowForm
    End Get
    Set(ByVal value As Boolean)
      mShowForm = value
    End Set
  End Property

  Private mSubject As String
  Public Property Subject() As String
    Get
      Return mSubject
    End Get
    Set(ByVal value As String)
      mSubject = value
      MyBase.fmSubject = value
    End Set
  End Property

  Private mBody As String
  Public Property Body() As String
    Get
      Return mBody
    End Get
    Set(ByVal value As String)
      mBody = value
      MyBase.fmBody = value
    End Set
  End Property

  Private mNomFitxer As String
  Public Property NomFitxer() As String
    Get
      Return mNomFitxer
    End Get
    Set(ByVal value As String)
      mNomFitxer = value
    End Set
  End Property

#End Region

#End Region

  Public Sub New()
    MyBase.New()

    PrtSettings.DefaultPageSettings.Landscape = False

    DetailColumns = 3
    TotalsColumns = 6
    WidthLines = 740
    BodyLeftOffset = 25

    ReDim ColumnWidth(DetailColumns - 1)
    ReDim TotalsWidth(TotalsColumns - 1)

    BottomYLines = 940

    ColumnWidth(0) = 460 ' si canvia repasar DrawTextLeft()
    ColumnWidth(1) = 130
    ColumnWidth(2) = 140

    TotalsWidth(0) = 145
    TotalsWidth(1) = 160
    TotalsWidth(2) = 150
    TotalsWidth(3) = 50
    TotalsWidth(4) = 100
    TotalsWidth(5) = 140

    clrDarkAzure = Color.FromArgb(42, 66, 106)
    clrLightAzure = Color.FromArgb(178, 195, 224)

    brshDarkAzure = New SolidBrush(clrDarkAzure)

    sfCenter = New StringFormat(StringFormatFlags.NoWrap)
    sfCenter.Alignment = StringAlignment.Center
    sfCenter.LineAlignment = StringAlignment.Center
    sfCenter.Trimming = StringTrimming.Character

    sfNear.Alignment = StringAlignment.Near
    sfNear.LineAlignment = StringAlignment.Center

    sfTotalBox.Alignment = StringAlignment.Far
    sfTotalBox.LineAlignment = StringAlignment.Near

    sfFar.Alignment = StringAlignment.Far
    sfFar.LineAlignment = StringAlignment.Center

    fntTitle = New Font("Arial Narrow", 8, FontStyle.Regular)
    fntValue = New Font("Arial", 10, FontStyle.Regular)
    fntItalic = New Font("Arial", 10, FontStyle.Italic)
    fntHdrValue = New Font("Arial Narrow", 9, FontStyle.Regular)
    fntAnagrama = New Font("HPDXCB", 10, FontStyle.Regular)
    fntInscrita = New Font("Arial Narrow", 6, FontStyle.Regular)
    fntLine = New Font("Arial", 11, FontStyle.Regular)

    marginBox = 5

  End Sub

  Public Overrides Sub BeginPrint()
    FirstPassReport = True
    FirstPage = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False
  End Sub

  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim hasMoreData As Boolean
    Dim printingSummaryData As Boolean

    If Not DataLoaded Then
      '   Return False
    End If

    deltaLine = CInt(Canvas.MeasureString("Íg", fntLine).Height)

    CurrentPage += 1

    Try
      PrintHeader(Canvas)
    Catch ex As Exception
      If AppData.Debug Then
        MessageBox.Show("Error al Imprimir la capçalera:" + TraceLocation)
      End If
    End Try


    If FirstPassReport Then
      ' Inicilització primera pasada del llistat
      FirstPassReport = False
    End If

    Do While True
      Select Case NowPrinting

        Case NowPrintingEnum.LinFactura

          hasMoreData = FillDetail(Canvas)

          If hasMoreData Then
            Exit Do
          End If

          NowPrinting = NowPrintingEnum.Summary

        Case NowPrintingEnum.TextClient

          If TextLeft.Length = 0 Then
            TextLeft = ""
          End If

          If TextLeft.Length = 0 Then
            NowPrinting = NowPrintingEnum.Summary
          Else
            If DrawTextLeft(Canvas) Then
              NowPrinting = NowPrintingEnum.Summary
            Else
              hasMoreData = True
              Exit Do
            End If
          End If


        Case NowPrintingEnum.Summary
          Try
            FillSummary(Canvas)
          Catch ex As Exception
            MessageBox.Show("Error al imprimir el sumari")
          End Try

          printingSummaryData = False

          Exit Do

      End Select

    Loop

    Return hasMoreData Or printingSummaryData

  End Function

  Private Sub PrintHeader(ByVal Canvas As System.Drawing.Graphics)
    ' Imprimeix les parts fixes de l'albara.
    Dim curX As Integer
    Dim width As Integer
    Dim height As Integer
    Dim i As Integer
    Dim wLogo As Integer
    Dim wAnagrama As Integer

    Dim sFac As String = String.Empty
    Dim sEnv As String = String.Empty

    ' Regla de paper per a mesurar
    'For x As Integer = 0 To 800 Step 10
    '  DrawLine(Canvas, Pens.Black, x, 0, x, 10)
    '  If x Mod 100 = 0 And x > 0 Then
    '    DrawLine(Canvas, Pens.Black, x, 12, x, 22)
    '    DrawString(Canvas, x.ToString, fntLine, Brushes.Black, x - 10, 25)
    '  End If
    'Next

    'For y As Integer = 0 To 1150 Step 10
    '  DrawLine(Canvas, Pens.Black, 0, y, 10, y)
    '  If y Mod 100 = 0 And y > 0 Then
    '    DrawLine(Canvas, Pens.Black, 12, y, 22, y)
    '    DrawString(Canvas, y.ToString, fntLine, Brushes.Black, 25, y - 2)
    '  End If
    'Next

    'Return

    ' Imprimir logo
    TraceLocation = "Logo"
    wLogo = 250
    wAnagrama = 300
    curX = BodyLeftOffset
    CurY = 10

    If Not Logo Is Nothing Then
      DrawImage(Canvas, Logo, curX + 5, CurY, wLogo, wLogo * Logo.Height \ Logo.Width)
    End If
    If Not LogoDimmed Is Nothing Then
      DrawImage(Canvas, LogoDimmed, 100, 450, 650, 650 * Logo.Height \ Logo.Width)
    End If

    'curX += 10
    'DrawFittedText(Canvas, curX, CurY, 300, 50, "Gestions.cat", New Font("Verdana", 24, FontStyle.Regular), Brushes.Black, Brushes.White, 0, 0)

    curX = 10
    CurY += wLogo * Logo.Height \ Logo.Width + 5
    DrawString(Canvas, "Polígono Industrial Oriola - Calle Támesis nº 27-51", fntAnagrama, brshDarkAzure, New RectangleF(curX, CurY, wAnagrama, fntAnagrama.Height), sfCenter)
    CurY += fntAnagrama.Height
    DrawString(Canvas, "Tel: 977 706 115   Fax: 977 706 141", fntAnagrama, brshDarkAzure, New RectangleF(curX, CurY, wAnagrama, fntAnagrama.Height), sfCenter)
    CurY += fntAnagrama.Height
    DrawString(Canvas, "43870 - Amposta (Tarragona)", fntAnagrama, brshDarkAzure, New RectangleF(curX, CurY, wAnagrama, fntAnagrama.Height), sfCenter)
    CurY += fntAnagrama.Height
    DrawString(Canvas, "e-mail: viestendals@gmail.com", fntAnagrama, brshDarkAzure, New RectangleF(curX, CurY, wAnagrama, fntAnagrama.Height), sfCenter)

    ' Codi client
    FillTopRoundedRectangle(Canvas, New SolidBrush(clrLightAzure), BodyLeftOffset, 210, 412, 20, 10)
    DrawRoundedRectangle(Canvas, BodyLeftOffset, 210, 412, 65, 10, 1, clrDarkAzure)
    Dim w As Single = 103
    DrawLine(Canvas, BodyLeftOffset, 210 + 20, 437, 210 + 20, 1, clrDarkAzure)
    DrawLine(Canvas, BodyLeftOffset + w, 210, BodyLeftOffset + w, 210 + 65, 1, clrDarkAzure)
    DrawLine(Canvas, BodyLeftOffset + w * 2, 210, BodyLeftOffset + w * 2, 210 + 65, 1, clrDarkAzure)
    DrawLine(Canvas, BodyLeftOffset + w * 3, 210, BodyLeftOffset + w * 3, 210 + 65, 1, clrDarkAzure)
    DrawString(Canvas, "Cód. Cliente", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset, 210, w, 20), sfCenter)
    DrawString(Canvas, "N.I.F.", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset + w, 210, w, 20), sfCenter)
    DrawString(Canvas, "Fecha Factura", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset + w * 2, 210, w, 20), sfCenter)
    DrawString(Canvas, "Número Factura", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset + w * 3, 210, w, 20), sfCenter)

    ' Nom client
    DrawRoundedRectangle(Canvas, 460, 140, 310, 135, 10, 1, clrDarkAzure)

    ' Cos
    FillTopRoundedRectangle(Canvas, New SolidBrush(clrLightAzure), BodyLeftOffset, 290, 745, 20, 10)
    DrawRoundedRectangle(Canvas, BodyLeftOffset, 290, 745, 650, 10, 1, clrDarkAzure)
    DrawLine(Canvas, BodyLeftOffset, 290 + 20, 770, 290 + 20, 1, clrDarkAzure)
    DrawLine(Canvas, 490, 290, 490, 290 + 650, 1, clrDarkAzure)
    DrawLine(Canvas, 620, 290, 620, 290 + 650, 1, clrDarkAzure)
    DrawString(Canvas, "Concepto", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset, 290, 465, 20), sfCenter)
    DrawString(Canvas, "Precio Unitario", fntTitle, brshDarkAzure, New RectangleF(490, 290, 130, 20), sfCenter)
    DrawString(Canvas, "Importe", fntTitle, brshDarkAzure, New RectangleF(620, 290, 150, 20), sfCenter)
    TopYLines = 290 + 20 + 10

    ' Sumary
    FillTopRoundedRectangle(Canvas, New SolidBrush(clrLightAzure), BodyLeftOffset, 955, 745, 20, 10)
    DrawRoundedRectangle(Canvas, BodyLeftOffset, 955, 745, 50, 10, 1, clrDarkAzure)
    DrawLine(Canvas, BodyLeftOffset, 955 + 20, 770, 955 + 20, 1, clrDarkAzure)
    DrawLine(Canvas, 170, 955, 170, 955 + 50, 1, clrDarkAzure)
    DrawLine(Canvas, 330, 955, 330, 955 + 50, 1, clrDarkAzure)
    DrawLine(Canvas, 480, 955, 480, 955 + 50, 1, clrDarkAzure)
    DrawLine(Canvas, 530, 955, 530, 955 + 50, 1, clrDarkAzure)
    DrawLine(Canvas, 630, 955, 630, 955 + 50, 1, clrDarkAzure)
    DrawString(Canvas, "Importe", fntTitle, brshDarkAzure, New RectangleF(BodyLeftOffset, 955, 145, 20), sfCenter)
    DrawString(Canvas, "", fntTitle, brshDarkAzure, New RectangleF(170, 955, 170, 20), sfCenter)
    DrawString(Canvas, "Base Imponible", fntTitle, brshDarkAzure, New RectangleF(330, 955, 150, 20), sfCenter)
    DrawString(Canvas, "% I.V.A.", fntTitle, brshDarkAzure, New RectangleF(480, 955, 50, 20), sfCenter)
    DrawString(Canvas, "Importe", fntTitle, brshDarkAzure, New RectangleF(530, 955, 100, 20), sfCenter)
    DrawString(Canvas, "Importe Total", fntTitle, brshDarkAzure, New RectangleF(630, 955, 140, 20), sfCenter)
    TopYTotals = 980

    ' Pagament
    DrawRoundedRectangle(Canvas, BodyLeftOffset, 1020, 745, 90, 10, 1, clrDarkAzure)
    DrawString(Canvas, "FORMA DE PAGO:", fntTitle, brshDarkAzure, 30, 1020 + 20)
    DrawString(Canvas, "VENCIMIENTO:", fntTitle, brshDarkAzure, 30, 1020 + 40)
    DrawString(Canvas, "IMPORTE:", fntTitle, brshDarkAzure, 30, 1020 + 60)
    TopYVenciments = 1020 + 18

    ' Inscrita
    DrawRotateText(Canvas, 10, 900, 270, Inscrita, fntInscrita, New SolidBrush(clrDarkAzure))


    ' TraceLocation = "Agafant camps de la adreça postal"
    sFac = cap("fc_nomcli").ToString + vbCrLf
    If Not IsNullOrEmptyValue(cap("fc_anagram").ToString) Then sFac += cap("fc_anagram").ToString + vbCrLf
    sFac += cap("fc_adrcli").ToString + vbCrLf
    sFac += cap("fc_cpcli").ToString.Trim + " - "
    sFac += cap("fc_pobcli").ToString + vbCrLf
    sFac += cap("fc_procli").ToString

    curX = 460 : CurY = 140 : width = 300 : height = 130
    DrawString(Canvas, sFac, fntValue, Brushes.Black, New RectangleF(curX + 10, CurY, width - 10, height), sfNear)

    ' Imprimir Recuadres Info Factura -------------------------------------------------------
    ' Primer recuadre --------------------------
    CurY = 210 + 20

    curX = BodyLeftOffset : DrawString(Canvas, cap("fc_codcli").ToString, fntValue, Brushes.Black, New RectangleF(curX, CurY, 103, 45), sfCenter)
    curX += 103 : DrawString(Canvas, cap("fc_nifcli").ToString, fntValue, Brushes.Black, New RectangleF(curX, CurY, 103, 45), sfCenter)
    curX += 103 : DrawString(Canvas, String.Format("{0:dd/MM/yyyy}", cap("fc_data")), fntValue, Brushes.Black, New RectangleF(curX, CurY, 103, 45), sfCenter)
    curX += 103 : DrawString(Canvas, String.Format("{0}/{1}", cap("fc_any"), cap("fc_numero").ToString.Trim), fntValue, Brushes.Black, New RectangleF(curX, CurY, 103, 45), sfCenter)

    ' Cleanup

    CurY = TopYLines

  End Sub

  Private Function DrawItemBoxHeader(ByVal Canvas As System.Drawing.Graphics, ByVal curX As Integer, ByVal curY As Integer, ByVal width As Integer, ByVal height As Integer, ByVal delta As Integer, ByVal Title As String, ByVal TitleAlign As StringFormat, ByVal Value As String, ByVal ValueAlign As StringFormat, ByVal LastItem As Boolean) As Integer

    DrawString(Canvas, Title, fntTitle, Brushes.Black, New RectangleF(curX, curY, width, delta), TitleAlign)
    DrawString(Canvas, Value, fntHdrValue, Brushes.Black, New RectangleF(curX, curY + delta, width, height - delta), ValueAlign)
    If Not LastItem Then
      DrawLine(Canvas, Pens.Black, curX + width, curY, curX + width, curY + height)
    End If

    Return curX + width

  End Function

  Private Function FillDetail(ByVal Canvas As System.Drawing.Graphics) As Boolean

    ' Imprimeix texte pendent de la última linea de la pagina
    If TextLeft <> String.Empty Then
      If Not DrawTextLeft(Canvas) Then
        Return True
      End If
    End If

    If AmpliacioDescripcio <> String.Empty Then
      TextLeft = AmpliacioDescripcio
      AmpliacioDescripcio = String.Empty
      SettingsLeft(15, fntLine, 0)
      If Not DrawTextLeft(Canvas) Then
        Return True
      End If
    End If

    If Not LinesLeft Then
      Return False
    End If

    Do

      'If False And LastExpedientID <> String.Format("{0:D4}{1:D6}", CInt(lin("fl_any")), CInt(lin("fl_numero"))) Then
      '  TextLeft = String.Format("Referència: {0}", lin("fl_exped").ToString.Trim)
      '  LastExpedientID = String.Format("{0:D4}{1:D6}", CInt(lin("fl_any")), CInt(lin("fl_numero")))
      '  SettingsLeft(7, 0, fntLine, 1)
      '  If Not DrawTextLeft(Canvas) Then
      '    Return True
      '  End If
      '  If CurY > BottomYLines - deltaLine Then
      '    Return True
      '  End If
      'End If

      Try
        DrawLinFactura(Canvas)
      Catch ex As Exception
        MessageBox.Show("Error al imprimir linea de factura")
      End Try

      AmpliacioDescripcio = lin("fl_ampart").ToString.Trim

      LinesLeft = lin.Read

      If TextLeft <> String.Empty Then
        SettingsLeft(5, fntLine, 0)
        If Not DrawTextLeft(Canvas) Then
          Return True
        End If
      End If

      If AmpliacioDescripcio <> String.Empty Then
        TextLeft = AmpliacioDescripcio
        AmpliacioDescripcio = String.Empty
        SettingsLeft(15, fntLine, 0)
        If Not DrawTextLeft(Canvas) Then
          Return True
        End If
      End If

      If Not LinesLeft Then
        Return False
      End If

      If CurY > BottomYLines - deltaLine Then
        Return True
      End If

    Loop

  End Function

  Private Sub DrawLinFactura(ByVal Canvas As System.Drawing.Graphics)
    Dim curX As Integer
    Dim i As Integer
    Dim Import As Decimal = 0D
    Dim strImport As String = String.Empty
    Dim ImportUnitari As Decimal = 0D
    Dim strImportUnitari As String = String.Empty

    Dim fittedChars As Integer
    Dim fittedLines As Integer

    ImportUnitari = CNull(lin("fl_prart"), 0D)
    Import = CNull(lin("fl_import"), 0D)

    If ImportUnitari <> 0D Then strImportUnitari = String.Format("{0:C}", ImportUnitari) Else strImportUnitari = String.Empty
    If Import <> 0D Then strImport = String.Format("{0:C}", Import) Else strImport = String.Empty

    curX = BodyLeftOffset + 5
    i = 0

    DrawString(Canvas, lin("fl_desart").ToString, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(0) - marginBox * 2, deltaLine), sfNear)
    Canvas.MeasureString(lin("fl_desart").ToString.Trim, fntLine, New SizeF(ColumnWidth(0) - marginBox * 2, deltaLine), New StringFormat, fittedChars, fittedLines)
    TextLeft = lin("fl_desart").ToString.Trim
    If TextLeft.Length > fittedChars Then
      TextLeft = TextLeft.Substring(fittedChars)
    Else
      TextLeft = String.Empty
    End If


    curX += ColumnWidth(0)
    DrawString(Canvas, strImport, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(1) - marginBox * 2, deltaLine), sfFar)

    curX += ColumnWidth(1)
    DrawString(Canvas, strImportUnitari, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(2) - marginBox * 2, deltaLine), sfFar)

    CurY += deltaLine

  End Sub

  Private Sub FillSummary(ByVal Canvas As System.Drawing.Graphics)
    Dim marginBox As Integer = 5
    Dim delta As Integer = 20
    Dim height As Integer
    Dim TextNotes As String
    Dim i As Integer

    Dim Import As Decimal
    Dim strImport As String = String.Empty
    Dim BaseImponible As Decimal
    Dim strBaseImponible As String = String.Empty
    Dim tpcIva As Decimal
    Dim strTpcIva As String = String.Empty
    Dim ImporteIva As Decimal
    Dim strImporteIva As String = String.Empty
    Dim ImporteTotal As Decimal
    Dim strImporteTotal As String = String.Empty
    Dim Pagament As String

    Import = CNull(cap("fc__bases"), 0D)
    BaseImponible = CNull(cap("fc__base1"), 0D) + CNull(cap("fc__base2"), 0D) + CNull(cap("fc__base3"), 0D)
    tpcIva = CNull(cap("fc__tpci1"), 0D)
    ImporteIva = CNull(cap("fc__civa1"), 0D) + CNull(cap("fc__civa2"), 0D) + CNull(cap("fc__civa3"), 0D)
    ImporteTotal = CNull(cap("fc__total"), 0D)

    If Import <> 0D Then strImport = String.Format("{0:C}", Import) Else strImport = String.Empty
    If BaseImponible <> 0D Then strBaseImponible = String.Format("{0:C}", BaseImponible) Else strImport = String.Empty
    If tpcIva <> 0D Then strTpcIva = String.Format("{0:N0}%", tpcIva) Else strImport = String.Empty
    If ImporteIva <> 0D Then strImporteIva = String.Format("{0:C}", ImporteIva) Else strImport = String.Empty
    If ImporteTotal <> 0D Then strImporteTotal = String.Format("{0:C}", ImporteTotal) Else strImport = String.Empty


    CurX = BodyLeftOffset
    CurY = TopYTotals

    DrawString(Canvas, strImport, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(0) - marginBox * 2, deltaLine), sfCenter)
    CurX += TotalsWidth(0) '  DrawString(Canvas, strDescompte, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(1) - marginBox * 2, deltaLine), sfCenter)
    CurX += TotalsWidth(1) : DrawString(Canvas, strBaseImponible, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(2) - marginBox * 2, deltaLine), sfCenter)
    CurX += TotalsWidth(2) : DrawString(Canvas, strTpcIva, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(3) - marginBox * 2, deltaLine), sfCenter)
    CurX += TotalsWidth(3) : DrawString(Canvas, strImporteIva, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(4) - marginBox * 2, deltaLine), sfCenter)
    CurX += TotalsWidth(4) : DrawString(Canvas, strImporteTotal, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(5) - marginBox * 2, deltaLine), sfCenter)

    'Venciments
    Dim VencimentData As String = String.Empty
    Dim VencimentImport As String = String.Empty

    If Utils.CNull(cap("fc__total"), 0D) > 0 Then
      ' Si es una devolució no pintem res.

      Pagament = cap("fc_banc").ToString.Trim
      If Not IsNullOrEmptyValue(cap("fc_iban").ToString) Then
        Pagament += " - IBAN: " + cap("fc_iban").ToString
      End If

      CurX = 140 : CurY = TopYVenciments

      DrawString(Canvas, Pagament, fntValue, Brushes.Black, CurX, CurY)
      CurY += 20

      If reb.Rows.Count > 0 Then
        For Each r As DataRow In reb.Rows
          DrawString(Canvas, String.Format("{0:dd/MM/yyyy}", r("re_dvto")), fntValue, Brushes.Black, CurX, CurY)
          DrawString(Canvas, String.Format("{0:C}", r("re_import")), fntValue, Brushes.Black, CurX, CurY + 20)
          CurX += 110
        Next
      End If

    End If

    TextNotes = "Les seves dades personals son incorporades a un " + _
     "fitxer propietat de " + NomClientCustom + ". " + _
     "Si desitja ejercir els seus drets d'accés, rectificació, " + _
     "cancelació i/o oposició, pot " + _
     "adreçar-se per escrit a les nostres oficines."
    height = 50

    '    DrawString(Canvas, TextNotes, fntNotes, Brushes.Black, New RectangleF(CurX + marginBox, TopYComentaris + 2, widthNotes - marginBox, height), sfNear)

  End Sub

  Private Sub SettingsLeft(ByVal pTextLeft As String, ByVal pOffsetLeft As Integer, ByVal pFontLeft As Font, ByVal pColumnLeft As Integer)
    TextLeft = pTextLeft
    OffsetLeft = pOffsetLeft
    OffsetTop = 0
    fntLeft = pFontLeft
    ColumnLeft = pColumnLeft
  End Sub

  Private Sub SettingsLeft(ByVal pOffsetLeft As Integer, ByVal pFontLeft As Font, ByVal pColumnLeft As Integer)
    OffsetTop = 0
    OffsetLeft = pOffsetLeft
    fntLeft = pFontLeft
    ColumnLeft = pColumnLeft
  End Sub

  Private Sub SettingsLeft(ByVal pOffsetTop As Integer, ByVal pOffsetLeft As Integer, ByVal pFontLeft As Font, ByVal pColumnLeft As Integer)
    OffsetTop = pOffsetTop
    OffsetLeft = pOffsetLeft
    fntLeft = pFontLeft
    ColumnLeft = pColumnLeft
  End Sub

  Private Function DrawTextLeft(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim szFree As SizeF
    Dim szUsed As SizeF
    Dim sf As New StringFormat
    Dim i As Integer

    sf.Alignment = StringAlignment.Near
    sf.LineAlignment = StringAlignment.Near

    Dim linesFitted As Integer
    Dim charsFitted As Integer

    'szFree = New SizeF(ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY)
    szFree = New SizeF(ColumnWidth(ColumnLeft) - OffsetLeft, (BottomYLines - 4) - CurY)
    szUsed = Canvas.MeasureString(TextLeft, fntLeft, szFree, sf, charsFitted, linesFitted)

    CurX = BodyLeftOffset + 5
    i = 0
    Do While i < ColumnLeft
      CurX += ColumnWidth(i)
      i += 1
    Loop

    'DrawString(Canvas, TextLeft, fntLeft, Brushes.Black, New RectangleF(CurX + OffsetLeft, CurY, ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY), sf)
    DrawString(Canvas, TextLeft, fntLeft, Brushes.Black, New RectangleF(CurX + OffsetLeft, CurY + OffsetTop, 650 - OffsetLeft, (BottomYLines - 4) - CurY + OffsetTop), sf)

    If charsFitted < TextLeft.Length Then
      TextLeft = TextLeft.Remove(0, charsFitted)
    Else
      TextLeft = String.Empty
    End If

    CurY += (CInt(szUsed.Height) + OffsetTop)

    Return (TextLeft.Length = 0)

  End Function

  Protected Overrides Sub Print2Excel(ByVal FileName As String)

  End Sub

  Private Sub clsImpresFactura_GetDataSource() Handles Me.GetDataSource

    Try

      cap = dbaCap.GetFactura(OrigenDades, SerieFactura, AnyFactura, NumeroFactura)
      reb = dbaReb.GetRebuts(SerieFactura, AnyFactura, NumeroFactura)

      lin = New csTableReader(dbaLin.GetLiniesFactura(OrigenDades, SerieFactura, AnyFactura, NumeroFactura))

      LinesLeft = lin.Read

      DataLoaded = True

    Catch ex As Exception
      DataLoaded = False
      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fac " + ex.Message + vbCrLf + String.Format("Any: {0}, Número: {1}", AnyFactura, NumeroFactura))
    End Try

    LastExpedientID = Nothing
    NowPrinting = NowPrintingEnum.LinFactura

  End Sub

  Protected Overrides Sub Finalize()
    MyBase.Finalize()

    fntTitle.Dispose()
    fntValue.Dispose()
    fntItalic.Dispose()
    fntHdrValue.Dispose()
    fntLine.Dispose()
    fntAnagrama.Dispose()
    fntInscrita.Dispose()

  End Sub

#Region " Print "

  Private Sub Execute()

    EmpresaName = AppData.CurrentEmpresaName
    ReportName = "Factura"
    ReportID = "R90EXP0003C"

    CustomID = "1" 'Gestions

    LayoutOffset = csRpt.LayoutOffsetEnum.OneThird
    PageNumbering = csRpt.PageNumberEnum.PageNofM

    SetDefaultPrinter()
    ShowPrintDialog = False

    pdfShowSaveDialog = False
    pdfPathAndFileName = String.Empty
    pdfDirectori = My.Settings.OutputDirPDF
    pdfNomFitxer = FitxerPdf

    ShowMessageError = AppData.Debug

    Try

      Print()

    Catch ex As Exception

      DebugLog("R90EXP0003C: " + ex.Message)

    Finally

      MailSentOK = fmMailSentOK

    End Try

  End Sub

#End Region

End Class
