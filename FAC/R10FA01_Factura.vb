Imports csAppData
Imports csUtils
Imports csUtils.Utils
Imports System.Windows
Imports System.Drawing

Public Class R90EXP0003C_Factura
  Inherits csRpt

  Private FirstPage As Boolean
  Private mIdiomaID As Integer
  Private mCustomID As Integer
  Private Logo As System.Drawing.Image
  Private Inscrita As String
  Private Empresa_A As String
  Private Empresa_B As String

  Private dbaCap As New C00_exp_fcp
  Private dbaCli As New C00_exp_cli
  Private dbaLin As New C00_exp_fln
  Private dbaCol As New C00_exp_col
  Private dbaSis As New C00_exp_sis

  Private dtCap As DataTable
  Private cap As DataRow
  Private cli As DataRow
  Private lin As Data.DataTableReader

  Private ColumnWidth() As Integer
  Private TotalsWidth() As Integer
  Private DetailColumns As Integer
  Private TotalsColumns As Integer
  Private DataLoaded As Boolean
  Private sfCenter As New StringFormat(StringFormatFlags.NoWrap)
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

  Private deltaLine As Integer

  Private TextLeft As String = String.Empty
  Private fntLeft As Font
  Private OffsetLeft As Integer
  Private ColumnLeft As Integer
  Private LinesLeft As Boolean

  Private LastAlbaraID As String

  Private Enum NowPrintingEnum
    LinFactura
    TextClient
    Summary
  End Enum

  Private NowPrinting As NowPrintingEnum

  Public Shadows Property CustomID() As Integer
    Get
      Return mCustomID
    End Get
    Set(ByVal value As Integer)
      If mCustomID <> value Then

        mCustomID = value
        Inscrita = ""

        Try
          Logo = System.Drawing.Bitmap.FromStream(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream("ColGestors.png"))
        Catch ex As Exception
          Logo = Nothing
        End Try

      End If

    End Set
  End Property

#Region " Properties "

#Region " Propietat factura "

  Private mEmpresa As String
  Public Property Empresa() As String
    Get
      Return mEmpresa
    End Get
    Set(ByVal value As String)
      mEmpresa = value
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

  Private mDeNumeroFactura As Integer
  Public Property DeNumeroFactura() As Integer
    Get
      Return mDeNumeroFactura
    End Get
    Set(ByVal value As Integer)
      mDeNumeroFactura = value
    End Set
  End Property

  Private mANumeroFactura As Integer
  Public Property aNumeroFactura() As Integer
    Get
      Return mANumeroFactura
    End Get
    Set(ByVal value As Integer)
      mANumeroFactura = value
    End Set
  End Property

  Private mDeDataFactura As DateTime
  Public Property deDataFactura() As DateTime
    Get
      Return mDeDataFactura
    End Get
    Set(ByVal value As DateTime)
      mDeDataFactura = value
    End Set
  End Property

  Private maDataFactura As DateTime
  Public Property aDataFactura() As DateTime
    Get
      Return maDataFactura
    End Get
    Set(ByVal value As DateTime)
      maDataFactura = value
    End Set
  End Property

  Private mNifClient As String
  Public Property NifClient() As String
    Get
      Return mNifClient
    End Get
    Set(ByVal value As String)
      mNifClient = value
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

  Public Overrides Function DrawPage(ByVal Canvas As System.Drawing.Graphics) As Boolean
    Dim hasMoreData As Boolean
    Dim printingSummaryData As Boolean

    If Not DataLoaded Then
      Return False
    End If

    deltaLine = CInt(Canvas.MeasureString("Íg", fntLine).Height)


    CurrentPage += 1

    PrintHeader(Canvas)

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
          Else
            printingSummaryData = True
            If Utils.CNull(cli("T3000_TeCodiUE"), False) Then
              NowPrinting = NowPrintingEnum.TextClient
            Else
              NowPrinting = NowPrintingEnum.TextClient
            End If
          End If


        Case NowPrintingEnum.TextClient

          If TextLeft.Length = 0 Then
            TextLeft = cap("T1020_Texte").ToString.Trim
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
          FillSummary(Canvas)
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
    Dim delta As Integer
    Dim boxW As Integer
    Dim i As Integer
    Dim runningWidth As Integer
    Dim Value As String
    Dim wLogo As Integer

    Dim sFac As String = String.Empty
    Dim sEnv As String = String.Empty

    ' Imprimir logo
    If TipusSerie = 1 Then
      wLogo = 100
      curX = 10
      CurY = 0
      If Not Logo Is Nothing Then
        DrawImage(Canvas, Logo, curX, CurY, wLogo, wLogo * Logo.Height \ Logo.Width)
      End If
      curX += wLogo + 10
      DrawFittedText(Canvas, curX, CurY, 300, 50, "GESTORIA VIRGINIA MARTI", New Font("Verdana", 30, FontStyle.Regular), Brushes.Black, Brushes.White, 0, 0)
      CurY += 50 + 5
      DrawString(Canvas, "C/ Sant Magí 10, baixos", fntTitle, Brushes.Black, curX, CurY)
      curX += fntTitle.Height
      DrawString(Canvas, "43004 - TARRAGONA", fntTitle, Brushes.Black, curX, CurY)
      curX += fntTitle.Height + 5
      DrawString(Canvas, "Tel: 977 216 013   Fax: 977 213 362", fntTitle, Brushes.Black, curX, CurY)
      curX += fntTitle.Height + 5
      DrawString(Canvas, "e-mail: virginia@gestoriavirginia.com", fntTitle, Brushes.Black, curX, CurY)
      curX += fntTitle.Height + 5
      DrawString(Canvas, "Virginia Martí Llauradó", fntTitle, Brushes.Black, curX, CurY)
      curX += fntTitle.Height + 5
      DrawString(Canvas, "N.I.F.: 39 573 802 W", fntTitle, Brushes.Black, curX, CurY)
    Else
      CurY = 0
    End If


    sFac = cap("fc_nom").ToString + vbCrLf
    sFac += cap("fc_adresa").ToString + vbCrLf
    sFac += cap("fc_postal").ToString.Trim + " - "
    sFac += cap("fc_poble").ToString + vbCrLf
    sFac += dbaCli.GetProvincia(cap("fc_postal").ToString)

    curX = 420 : CurY = 210 : width = 360 : height = 110

    DrawRoundedRectangle(Canvas, curX, CurY + 20, width, height, 10, 1)
    DrawString(Canvas, sFac, fntValue, Brushes.Black, New RectangleF(curX + 10, CurY + 20, width - 10, height), sfNear)


    ' Imprimir Recuadres Info Factura -------------------------------------------------------
    ' Primer recuadre --------------------------
    CurY = 280 : runningWidth = 0

    curX = BodyLeftOffset : width = 380 : height = 40 : delta = 15
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)

    ' Numero Factura
    Value = IIf((Empresa_A + Empresa_B).Contains(cap("fc_empresa").ToString), "", cap("fc_empresa").ToString + "-").ToString
    Value += String.Format("{0:d8}", cap("fc_numero")).Substring(0, 2) + String.Format("{0:d8}", cap("fc_numero")).Substring(2)
    boxW = 90 : runningWidth = boxW
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "N.Factura", sfCenter, Value, sfCenter, False)

    ' Data factura
    boxW = 70 : runningWidth += boxW : Value = String.Format("{0:dd/MM/yyyy}", cap("fc_datfact"))
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Data", sfCenter, Value, sfCenter, False)

    ' NIF client
    boxW = 100 : runningWidth += boxW : Value = String.Format("{0}", cap("fc_nif")).Trim
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "N.I.F.", sfCenter, Value, sfCenter, False)

    ' Pàgina
    boxW = width - runningWidth : Value = String.Format("{0} / {1}", CurrentPage, TotalPages)
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Pàgina", sfCenter, Value, sfCenter, True)


    ' ----------------------------------------------------------------------------------------
    ' Imprimir recuadres linies albarà -------------------------------------------------------

    CurY += height + 10
    height = BottomYLines - CurY

    curX = BodyLeftOffset : width = WidthLines
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    delta = 20
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)

    TopYLines = CurY + delta + 3

    ' Columnes ---------------------------
    ' Codi producte del client
    i = -1
    i += 1 : boxW = ColumnWidth(i)
    DrawString(Canvas, "Codi", fntTitle, Brushes.Black, New RectangleF(curX, CurY, boxW, delta), sfCenter)
    curX += boxW : DrawLine(Canvas, Pens.Black, curX, CurY, curX, CurY + height)
    ' Codi descripció
    i += 1 : boxW = ColumnWidth(i)
    DrawString(Canvas, "Descripció", fntTitle, Brushes.Black, New RectangleF(curX, CurY, boxW, delta), sfCenter)
    curX += boxW : DrawLine(Canvas, Pens.Black, curX, CurY, curX, CurY + height)
    ' Honoraris
    i += 1 : boxW = ColumnWidth(i)
    DrawString(Canvas, "Honoraris", fntTitle, Brushes.Black, New RectangleF(curX, CurY, boxW, delta), sfCenter)
    curX += boxW : DrawLine(Canvas, Pens.Black, curX, CurY, curX, CurY + height)
    ' Suplits
    i += 1 : boxW = ColumnWidth(i)
    DrawString(Canvas, "Suplits", fntTitle, Brushes.Black, New RectangleF(curX, CurY, boxW, delta), sfCenter)
    curX += boxW : DrawLine(Canvas, Pens.Black, curX, CurY, curX, CurY + height)
    ' Bestretes
    i += 1 : boxW = ColumnWidth(i)
    DrawString(Canvas, "Bestretes", fntTitle, Brushes.Black, New RectangleF(curX, CurY, boxW, delta), sfCenter)
    curX += boxW : DrawLine(Canvas, Pens.Black, curX, CurY, curX, CurY + height)

    ' ----------------------------------------------------------------------------------------

    ' Imprimir Recuadres Peu Factura
    TotalFacturaBoxWidth = 140

    curX = BodyLeftOffset : CurY = BottomYLines + 10 : width = WidthLines - TotalFacturaBoxWidth - 10 : height = 75
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)
    TopYTotals = CurY + delta

    i = -1
    ' Honoraris
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Honoraris", sfCenter, Value, sfTotalBox, False)

    ' Suplits
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Suplits", sfCenter, Value, sfTotalBox, False)

    ' Base imposable
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Base imposable", sfCenter, Value, sfTotalBox, False)

    ' %IVA
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "% IVA", sfCenter, Value, sfTotalBox, False)

    ' Quota IVA 
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Quota IVA", sfCenter, Value, sfTotalBox, False)

    ' % Retenció
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "% Retenció", sfCenter, Value, sfTotalBox, False)

    ' Bestretes
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Bestretes", sfCenter, Value, sfTotalBox, False)

    ' Total factura
    i += 1 : boxW = TotalsWidth(i) : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Total factura", sfCenter, Value, sfTotalBox, False)

    ' ----------------------------------------------------------------------------------------

    '' Total.
    'curX = BodyLeftOffset + WidthLines - TotalFacturaBoxWidth : width = TotalFacturaBoxWidth : height = 75
    'DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    'DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)
    'TopYTotal = CurY + delta

    'boxW = width : Value = ""
    'If TipusSerie = 1 Then
    '  curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, dfi.FacPeuTotalFactura, sfCenter, Value, sfTotalBox, True)
    'Else
    '  curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, dfi.FacPeuTotalNotaCarrec, sfCenter, Value, sfTotalBox, True)
    'End If

    ' Imprimir Recuadres Venciments
    widthVto = 180

    curX = BodyLeftOffset : CurY += height + 10 : width = widthVto : height = 120
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)
    TopYVenciments = CurY + delta

    ' Venciments
    boxW = width : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Venciments", sfCenter, Value, sfTotalBox, True)

    ' Forma pago
    curX = BodyLeftOffset + widthVto + 10 : width = WidthLines - widthVto - 10 : height = 40
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)
    TopYPagament = CurY + delta

    ' Forma pago
    boxW = width : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "Forma pagament", sfNear, Value, sfTotalBox, True)

    ' Comentaris
    CurY += 10 + 40
    curX = BodyLeftOffset + widthVto + 10 : width = WidthLines - widthVto - 10 : height = 70
    DrawRoundedRectangle(Canvas, curX, CurY, width, height, 10, 1)
    DrawLine(Canvas, Pens.Black, curX, CurY + delta, curX + width, CurY + delta)
    TopYComentaris = CurY + delta

    ' Comentaris
    boxW = width : Value = ""
    curX = DrawItemBoxHeader(Canvas, curX, CurY, boxW, height, delta, "", sfNear, Value, sfTotalBox, True)

    ' ----------------------------------------------------------------------------------------

    ' Cleanup

    CurY = TopYLines

  End Sub

  Private Sub FillSummary(ByVal Canvas As System.Drawing.Graphics)
    Dim fntTotal As New Font("Arial", 10, FontStyle.Bold)
    Dim fntNotes As New Font("Arial", 7, FontStyle.Regular)

    Dim marginBox As Integer = 5
    Dim delta As Integer = 20
    Dim width As Integer
    Dim height As Integer
    Dim TextNotes As String
    Dim i As Integer

    ' Totals
    Dim TipoIva As String = String.Empty
    Dim TipoReq As String = String.Empty
    Dim SumaImports As String = String.Empty
    Dim Descompte As String = String.Empty
    Dim ProntoPago As String = String.Empty
    Dim Ports As String = String.Empty
    Dim BaseImposable As String = String.Empty
    Dim QuotaIva As String = String.Empty
    Dim QuotaReq As String = String.Empty
    Dim Pagament As String = String.Empty
    Dim Separator As String = String.Empty

    'Formategem la divisa


    If TipusSerie = 1 Then

      i = 0 : CurX = BodyLeftOffset : CurY = TopYTotals
      DrawString(Canvas, SumaImports, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, Descompte, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, ProntoPago, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, Ports, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, BaseImposable, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, TipoIva, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, QuotaIva, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, TipoReq, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)
      CurX += TotalsWidth(i) : i += 1
      DrawString(Canvas, QuotaReq, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, TotalsWidth(i) - marginBox, 55), sfTotalBox)

    Else
      CurY = TopYTotal
      CurX = BodyLeftOffset + (WidthLines - TotalFacturaBoxWidth * 2 - 10) + 10 : width = TotalFacturaBoxWidth : height = 55
      DrawString(Canvas, Ports.Replace(vbCrLf, String.Empty), fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY, TotalFacturaBoxWidth - marginBox, 55), sfFar)
    End If

    'Total factura
    CurX = BodyLeftOffset + (WidthLines - TotalFacturaBoxWidth - 10) + 10 : width = TotalFacturaBoxWidth : height = 55
    DrawString(Canvas, String.Format(mFormatValor, Utils.CNull(cap("T1020_Import"), 0D) * mCanviDivisa), fntTotal, Brushes.Black, New RectangleF(CurX, CurY, TotalFacturaBoxWidth - marginBox, 55), sfFar)

    'Venciments
    Dim VencimentData As String = String.Empty
    Dim VencimentImport As String = String.Empty

    Separator = vbCrLf

    'If Utils.CNull(cap("T1020_Import"), 0D) > 0 Then

    '  For Each r As DataRow In ds.Tables("Reb").Rows
    '    VencimentData += String.Format("{0:dd/MM/yyyy}{1}", r("T1040_DataVenciment"), Separator)
    '    VencimentImport += String.Format(mFormatValor + "{1}", Utils.CNull(r("T1040_ImportRebut"), 0D) * mCanviDivisa, Separator)
    '  Next

    '  CurX = BodyLeftOffset : CurY = TopYVenciments
    '  DrawString(Canvas, VencimentData, fntHdrValue, Brushes.Black, CurX + marginBox, CurY + marginBox)
    '  CurX = BodyLeftOffset + widthVto \ 2
    '  DrawString(Canvas, VencimentImport, fntHdrValue, Brushes.Black, New RectangleF(CurX, CurY + marginBox, widthVto \ 2 - marginBox, 120), sfTotalBox)

    'End If

    'Forma de pago
    Pagament = cap("Q3101_FormaPagament").ToString

    'If mTipusDocument = TipusDocsImpresEnum.DocumentsEnDivises And ImprimirEnDivises Then
    '  Pagament += String.Format(" ({0} 1 € = {1:N3} {2})", dfi.CanviDivisa, mCanviDivisa, cap("T2505_Simbol"))
    'End If

    'If TipusSerie = 1 Then
    '  If Not String.IsNullOrEmpty(cli("T3000_BancEntitat").ToString) Then
    '    Pagament += ". " + fpg("T3101_FormaPagament").ToString
    '  End If
    '  If Not String.IsNullOrEmpty(cli("T3000_BancCompte").ToString) Then
    '    Select Case CType(cap("T3131_TipusPagamentID"), TipusPagamentEnum)
    '      Case TipusPagamentEnum.DOMICILIACIO, TipusPagamentEnum.LCR, TipusPagamentEnum.PRELEVEMENT
    '        Dim BancCompte As String = Utils.CNull(cli("T3000_BancCompte"), "")
    '        If BancCompte.Length > 5 Then
    '          BancCompte = String.Format("{0}*****", BancCompte.Substring(0, BancCompte.Length - 5))
    '        End If

    '        Pagament += ". " + BancCompte
    '      Case Else
    '        'nothing
    '    End Select
    '  End If
    'End If
    'CurX = BodyLeftOffset + widthVto + 10 : CurY = TopYPagament
    'DrawString(Canvas, Pagament, fntHdrValue, Brushes.Black, CurX + marginBox, CurY + marginBox)

    ' Notes
    'CurX = BodyLeftOffset + widthVto + 10 : CurY = TopYComentaris : width = WidthLines - widthVto - 10 : height = 70
    'TextNotes = ""
    'If TipusSerie = 1 Then
    '  If Utils.CNull(cap("T1020_FacturaAmbBonsai"), False) Then
    '    If TextNotes.Length > 0 Then TextNotes += vbCrLf
    '    TextNotes += dfi.TexteFitosanitari
    '  End If
    '  If Utils.CNull(fpg("T3101_AsseguradaCyC"), 0) > 0 And Utils.CNull(cap("T3000_ImprimeixCreditoyCaucion"), False) Then
    '    If TextNotes.Length > 0 Then TextNotes += vbCrLf
    '    TextNotes += dfi.FacPeuCreditoCaucion
    '  End If
    'End If

    TextNotes = ""
    DrawString(Canvas, TextNotes, fntNotes, Brushes.Black, New RectangleF(CurX + marginBox, CurY + 2, width - marginBox * 2, delta), sfNear)

    fntTotal.Dispose()
    fntNotes.Dispose()

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

    If Not LinesLeft Then
      Return False
    End If

    Do

      If Not DrawTextLeft(Canvas) Then
        Return True
      End If

      If CurY > BottomYLines - deltaLine Then
        Return True
      End If

      DrawLinFactura(Canvas)

      LinesLeft = lin.Read

      If TextLeft <> String.Empty Then
        SettingsLeft(5, fntItalic, 3)
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

  Private Sub SettingsLeft(ByVal pTextLeft As String, ByVal pOffsetLeft As Integer, ByVal pFontLeft As Font, ByVal pColumnLeft As Integer)
    TextLeft = pTextLeft
    OffsetLeft = pOffsetLeft
    fntLeft = pFontLeft
    ColumnLeft = pColumnLeft
  End Sub

  Private Sub SettingsLeft(ByVal pOffsetLeft As Integer, ByVal pFontLeft As Font, ByVal pColumnLeft As Integer)
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

    szFree = New SizeF(ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY)
    szUsed = Canvas.MeasureString(TextLeft, fntLeft, szFree, sf, charsFitted, linesFitted)

    CurX = BodyLeftOffset + 5
    Do While i < ColumnLeft
      CurX += ColumnWidth(i)
      i += 1
    Loop

    DrawString(Canvas, TextLeft, fntLeft, Brushes.Black, New RectangleF(CurX + OffsetLeft, CurY, ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY), sf)

    If charsFitted < TextLeft.Length Then
      TextLeft = TextLeft.Remove(0, charsFitted)
    Else
      TextLeft = String.Empty
    End If

    CurY += CInt(szUsed.Height)

    Return (TextLeft.Length = 0)

  End Function

  Private Sub DrawLinFactura(ByVal Canvas As System.Drawing.Graphics)
    Dim curX As Integer
    Dim i As Integer
    Dim Import As Decimal = 0D
    Dim strImport As String = ""
    Dim Suplits As Decimal
    Dim Bestretes As Decimal

    Suplits = CNull(lin("fl_tasas"), 0D) + CNull(lin("fl_tasacom"), 0D)
    Bestretes = CNull(lin("fl_suplits"), 0D) + CNull(lin("fl_polgest"), 0D)


    curX = BodyLeftOffset + 5
    i = 0

    DrawString(Canvas, lin("fl_gestio").ToString, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(i) - marginBox * 2, deltaLine), sfNear)
    curX += ColumnWidth(i) : i += 1
    DrawString(Canvas, lin("fl_desc").ToString, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(i) - marginBox * 2, deltaLine), sfNear)
    curX += ColumnWidth(i) : i += 1
    DrawString(Canvas, CNull(lin("fl_honora"), 0D).ToString("c"), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(i) - marginBox * 2, deltaLine), sfFar)
    curX += ColumnWidth(i) : i += 1
    DrawString(Canvas, Suplits.ToString("c"), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(i) - marginBox * 2, deltaLine), sfFar)
    curX += ColumnWidth(i) : i += 1
    DrawString(Canvas, Bestretes.ToString("c"), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(i) - marginBox * 2, deltaLine), sfFar)
    curX += ColumnWidth(i) : i += 1


    TextLeft = lin("fl_desc").ToString

    CurY += deltaLine

  End Sub

  Public Overrides Sub BeginPrint()
    FirstPassReport = True
    FirstPage = True
    LoadDataSource = True
    CurrentPage = 0
    DataNeeded = True
    DrawingTotalsAndExit = False
  End Sub

  Protected Overrides Sub Print2Excel(ByVal FileName As String)

  End Sub

  Public Sub New()
    MyBase.New()

    PrtSettings.DefaultPageSettings.Landscape = False

    DetailColumns = 8
    TotalsColumns = 9
    WidthLines = 760
    BodyLeftOffset = 20

    ReDim ColumnWidth(DetailColumns - 1)
    ReDim TotalsWidth(TotalsColumns - 1)

    BottomYLines = 900

    ColumnWidth(0) = 90
    ColumnWidth(1) = 30
    ColumnWidth(2) = 60
    ColumnWidth(3) = 290 ' si canvia repasar DrawTextLeft()
    ColumnWidth(4) = 60
    ColumnWidth(5) = 80
    ColumnWidth(6) = 60
    ColumnWidth(7) = WidthLines
    For cw As Integer = 0 To DetailColumns - 2
      ColumnWidth(DetailColumns - 1) -= ColumnWidth(cw)
    Next

    TotalsWidth(0) = 80
    TotalsWidth(1) = 80
    TotalsWidth(2) = 65
    TotalsWidth(3) = 65
    TotalsWidth(4) = 80
    TotalsWidth(5) = 55
    TotalsWidth(6) = 60
    TotalsWidth(7) = 45
    TotalsWidth(8) = 80

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

    fntLine = New Font("Arial", 8, FontStyle.Regular)

    marginBox = 5

  End Sub

  Private Sub clsImpresFactura_GetDataSource() Handles Me.GetDataSource

    LastAlbaraID = Nothing
    NowPrinting = NowPrintingEnum.LinFactura

  End Sub

  Protected Overrides Sub Finalize()
    MyBase.Finalize()

    fntTitle.Dispose()
    fntValue.Dispose()
    fntItalic.Dispose()
    fntHdrValue.Dispose()
    fntLine.Dispose()

  End Sub

#Region " Print "

  Private Sub Execute()

    EmpresaName = AppData.CurrentEmpresaName
    ReportName = "Factura"
    ReportID = "R90EXP0003C"

    LayoutOffset = csRpt.LayoutOffsetEnum.OneThird
    PageNumbering = csRpt.PageNumberEnum.PageNofM

    Destination = Destinacio
    SetDefaultPrinter()
    ShowPrintDialog = False

    pdfShowSaveDialog = False
    pdfPathAndFileName = String.Empty
    pdfDirectori = My.Settings.OutputDirPDF
    pdfNomFitxer = FitxerPdf

    ShowMessageError = AppData.Debug

    'DebugLog("pepete")

    Try

      Print()

    Catch ex As Exception

    Finally

      MailSentOK = fmMailSentOK

    End Try

  End Sub

#End Region

End Class
