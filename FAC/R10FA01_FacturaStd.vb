Imports csAppData
Imports csUtils
Imports csUtils.Utils
Imports System.Windows
Imports System.Drawing

Public Class R90EXP0003C_FacturaStd
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

  Private char_x As Decimal
  Private char_y As Decimal

  Private cap_dv, lin_dv, sum_dv As dataview

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


    For Each drv As datarowview In cap_dv
      curx = 0

    Next

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

    char_x = 10.0
    char_Y = 16.66667

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
