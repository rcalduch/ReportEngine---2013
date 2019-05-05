Imports csAppData
Imports csUtils
Imports csUtils.Utils
Imports System.Windows
Imports System.Drawing
Imports System.Linq
Imports System.Reflection

Public Class R90FAC0003C_Factura
    Inherits csRpt

    Private _mCustomId As String
    Private _backgroundImage As Image

    Private _traceLocation As String

    Private ReadOnly _dbaCap As New C00_gst_fac
    Private ReadOnly _dbaLin As New C00_gst_fal
    Private ReadOnly _dbaReb As New C00_gst_reb
    Private ReadOnly _dbaFpg As New C00_gst_fpg
    ' Private dbaSys As New C00_gst_sys

    Private _cap As DataRow
    Private _lin As csTableReader
    Private _reb As DataTable

    Private ReadOnly _columnWidth() As Integer
    Private ReadOnly _totalsWidth() As Integer
    Private ReadOnly _detailColumns As Integer
    Private ReadOnly _totalsColumns As Integer
    Private _dataLoaded As Boolean
    Private ReadOnly _sfCenter As StringFormat
    Private ReadOnly _sfNear As New StringFormat
    Private ReadOnly _sfFar As New StringFormat
    Private ReadOnly _sfTotalBox As New StringFormat
    Private _bottomYLines As Integer
    Private _marginBox As Integer
    Private ReadOnly _bodyLeftOffset As Integer

    Private _topYVenciments As Integer
    Private _topYTotals As Integer

    Private ReadOnly _fntTitle As Font
    Private ReadOnly _fntValue As Font
    Private ReadOnly _fntHdrValue As Font
    Private ReadOnly _fntLine As Font
    Private ReadOnly _fntItalic As Font
    Private ReadOnly _fntAnagrama As Font
    Private ReadOnly _fntInscrita As Font
    Private ReadOnly _fntGrossa As Font

    Private _deltaLine As Integer

    Private _textLeft As String = String.Empty
    Private _ampliacioDescripcio As String = String.Empty
    Private _fntLeft As Font
    Private _offsetLeft As Integer
    Private _offsetTop As Integer
    Private _columnLeft As Integer
    Private _linesLeft As Boolean

    Private Enum NowPrintingEnum
        LinFactura
        TextClient
        Summary
    End Enum

    Private _nowPrinting As NowPrintingEnum

    Public Shadows Property CustomId As String
        Get
            Return _mCustomId
        End Get
        Set
            If _mCustomId <> Value Then

                _mCustomId = Value

                Try
                    _backgroundImage = Bitmap.FromStream(Assembly.GetExecutingAssembly.GetManifestResourceStream("csDosReportEngine.ipmfactura.png"))
                Catch ex As Exception
                    If AppData.Debug Then
                        MessageBox.Show($"Logo no trobat al Resources")
                    End If
                    _backgroundImage = Nothing
                End Try

            End If

        End Set
    End Property

#Region " Properties "

#Region " Propietat factura "

    Public Property OrigenDades As String

    Public Property SerieFactura As String

    Public Property NumeroFactura As Integer

    Public Property AnyFactura As String

#End Region

#Region " Propietats llistat "

    Private mDestinacio As ReportDestinationEnum
    Public Property Destinacio As ReportDestinationEnum
        Get
            Return mDestinacio
        End Get
        Set
            Destination = Value
            mDestinacio = Value
        End Set
    End Property

    Private mCopies As Integer
    Public Property Copies As Integer
        Get
            Return mCopies
        End Get
        Set
            mCopies = Value
        End Set
    End Property

    Private mFitxerPdf As String
    Public Property FitxerPdf As String
        Get
            Return mFitxerPdf
        End Get
        Set
            mFitxerPdf = Value
        End Set
    End Property

#End Region

#Region " Propietats correo "

    Private mMailFeedback As String
    Public Property MailFeedback As String
        Get
            Return mMailFeedback
        End Get
        Set
            mMailFeedback = Value
            MyBase.fmMailFeedBack = Value
        End Set
    End Property

    Private mMailFrom As String
    Public Property MailFrom As String
        Get
            Return mMailFrom
        End Get
        Set
            mMailFrom = Value
            MyBase.fmMailFrom = Value
        End Set
    End Property

    Private mMailReplyTo As String
    Public Property MailReplyTo As String
        Get
            Return mMailReplyTo
        End Get
        Set
            mMailReplyTo = Value
            MyBase.fmMailReplyTo = Value
        End Set
    End Property

    Private mMailTo As String
    Public Property MailTo As String
        Get
            Return mMailTo
        End Get
        Set
            mMailTo = Value.Trim
            MyBase.fmMailTo = Value.Trim
        End Set
    End Property

    Private mMailSentOK As Boolean
    Public Property MailSentOK As Boolean
        Get
            Return mMailSentOK
        End Get
        Set
            mMailSentOK = Value
        End Set
    End Property

    Private mSmtpLogin As String
    Public Property SmtpLogin As String
        Get
            Return mSmtpLogin
        End Get
        Set
            mSmtpLogin = Value
            MyBase.fmSmtpLogin = Value
        End Set
    End Property

    Private mSmtpPassword As String
    Public Property SmtpPassword As String
        Get
            Return mSmtpPassword
        End Get
        Set
            mSmtpPassword = Value
            MyBase.fmSmtpPassword = Value
        End Set
    End Property

    Private mSmtpServer As String
    Public Property SmtpServer As String
        Get
            Return mSmtpServer
        End Get
        Set
            mSmtpServer = Value
            MyBase.fmSmtpServer = Value
        End Set
    End Property

    Private mShowForm As Boolean
    Public Property ShowForm As Boolean
        Get
            Return mShowForm
        End Get
        Set
            mShowForm = Value
        End Set
    End Property

    Private mSubject As String
    Public Property Subject As String
        Get
            Return mSubject
        End Get
        Set
            mSubject = Value
            MyBase.fmSubject = Value
        End Set
    End Property

    Private mBody As String
    Public Property Body As String
        Get
            Return mBody
        End Get
        Set
            mBody = Value
            MyBase.fmBody = Value
        End Set
    End Property

    Private mNomFitxer As String
    Public Property NomFitxer As String
        Get
            Return mNomFitxer
        End Get
        Set
            mNomFitxer = Value
        End Set
    End Property

#End Region

#End Region

    Public Sub New()
        MyBase.New()

        PrtSettings.DefaultPageSettings.Landscape = False

        _detailColumns = 5
        _totalsColumns = 7
        _bodyLeftOffset = 25

        ReDim _columnWidth(_detailColumns - 1)
        ReDim _totalsWidth(_totalsColumns - 1)

        _bottomYLines = 940

        _columnWidth(0) = 60 ' si canvia repasar DrawTextLeft()
        _columnWidth(1) = 60
        _columnWidth(2) = 400
        _columnWidth(3) = 100
        _columnWidth(4) = 100

        _totalsWidth(0) = 145
        _totalsWidth(1) = 160
        _totalsWidth(2) = 150
        _totalsWidth(3) = 60
        _totalsWidth(4) = 140
        _totalsWidth(5) = 150

        'brshBackground = New SolidBrush(clrDarkAzure)

        _sfCenter = New StringFormat(StringFormatFlags.NoWrap)
        _sfCenter.Alignment = StringAlignment.Center
        _sfCenter.LineAlignment = StringAlignment.Center
        _sfCenter.Trimming = StringTrimming.Character

        _sfNear.Alignment = StringAlignment.Near
        _sfNear.LineAlignment = StringAlignment.Center

        _sfTotalBox.Alignment = StringAlignment.Far
        _sfTotalBox.LineAlignment = StringAlignment.Near

        _sfFar.Alignment = StringAlignment.Far
        _sfFar.LineAlignment = StringAlignment.Center

        _fntTitle = New Font("Arial Narrow", 8, FontStyle.Regular)
        _fntValue = New Font("Arial", 10, FontStyle.Regular)
        _fntItalic = New Font("Arial", 10, FontStyle.Italic)
        _fntHdrValue = New Font("Arial Narrow", 9, FontStyle.Regular)
        _fntAnagrama = New Font("HPDXCB", 10, FontStyle.Regular)
        _fntGrossa = New Font("HPDXCB", 22, FontStyle.Regular)
        _fntInscrita = New Font("Arial Narrow", 6, FontStyle.Regular)
        _fntLine = New Font("Arial", 9, FontStyle.Regular)

        _marginBox = 5

    End Sub

    Public Overrides Sub BeginPrint()
        FirstPassReport = True
        LoadDataSource = True
        CurrentPage = 0
        DataNeeded = True
        DrawingTotalsAndExit = False
    End Sub

    Public Overrides Function DrawPage(canvas As Graphics) As Boolean
        Dim hasMoreData As Boolean
        Dim printingSummaryData As Boolean

        If Not _dataLoaded Then
            '   Return False
        End If

        _deltaLine = CInt(canvas.MeasureString("Íg", _fntLine).Height)

        CurrentPage += 1

        Try
            PrintHeader(canvas)
        Catch ex As Exception
            If AppData.Debug Then
                MessageBox.Show("Error al Imprimir la capçalera:" + _traceLocation)
            End If
        End Try


        If FirstPassReport Then
            ' Inicilització primera pasada del llistat
            FirstPassReport = False
        End If

        Do While True
            Select Case _nowPrinting

                Case NowPrintingEnum.LinFactura

                    hasMoreData = FillDetail(canvas)

                    If hasMoreData Then
                        Exit Do
                    End If

                    _nowPrinting = NowPrintingEnum.Summary

                Case NowPrintingEnum.TextClient

                    If _textLeft.Length = 0 Then
                        _textLeft = ""
                    End If

                    If _textLeft.Length = 0 Then
                        _nowPrinting = NowPrintingEnum.Summary
                    Else
                        If DrawTextLeft(canvas) Then
                            _nowPrinting = NowPrintingEnum.Summary
                        Else
                            hasMoreData = True
                            Exit Do
                        End If
                    End If


                Case NowPrintingEnum.Summary
                    Try
                        FillSummary(canvas)
                    Catch ex As Exception
                        MessageBox.Show($"Error al imprimir el sumari")
                    End Try

                    printingSummaryData = False

                    Exit Do

            End Select

        Loop

        Return hasMoreData Or printingSummaryData

    End Function

    Private Sub PrintHeader(canvas As Graphics)
        ' Imprimeix les parts fixes de l'albara.
        Dim localCurX As Integer
        Dim width As Integer
        Dim height As Integer
        Dim wBackgroundImage As Integer
        Dim sFac As String
        Dim albs As New ArrayList
        Dim numAlbs As String = string.Empty

        ''Regla de paper per a mesurar
        'For x As Integer = 0 To 850 Step 10
        '    DrawLine(Canvas, Pens.Black, x, 0, x, 10)
        '    If x Mod 100 = 0 And x > 0 Then
        '        DrawLine(Canvas, Pens.Black, x, 12, x, 22)
        '        DrawString(Canvas, x.ToString, fntLine, Brushes.Black, x - 10, 25)
        '    End If
        'Next

        'For y As Integer = 0 To 1200 Step 10
        '    DrawLine(Canvas, Pens.Black, 0, y, 10, y)
        '    If y Mod 100 = 0 And y > 0 Then
        '        DrawLine(Canvas, Pens.Black, 12, y, 22, y)
        '        DrawString(Canvas, y.ToString, fntLine, Brushes.Black, 25, y - 2)
        '    End If
        'Next

        'Return

        ' Imprimir logo
        _traceLocation = "Logo"
        wBackgroundImage = 820

        If Not _backgroundImage Is Nothing Then
            DrawImage(canvas, _backgroundImage, 0, 0, wBackgroundImage, wBackgroundImage * _backgroundImage.Height \ _backgroundImage.Width)
        End If

        ' Cos

        _bottomYLines = 780

        ' Sumary

        _topYTotals = 825

        ' Pagament

        _topYVenciments = 880


        ' TraceLocation = "Agafant camps de la adreça postal"
        sFac = _cap("fc_nomcli").ToString + vbCrLf
        If Not IsNullOrEmptyValue(_cap("fc_anagram").ToString) Then sFac += _cap("fc_anagram").ToString + vbCrLf
        sFac += _cap("fc_adrcli").ToString + vbCrLf
        sFac += _cap("fc_cpcli").ToString.Trim + " - "
        sFac += _cap("fc_pobcli").ToString + vbCrLf
        sFac += _cap("fc_procli").ToString

        localCurX = 390 : CurY = 160 : width = 360 : height = 120
        DrawString(canvas, sFac, _fntValue, Brushes.Black, New RectangleF(localCurX, CurY, width - 10, height), _sfNear)

        DrawString(canvas, $"{_cap("fc_numero").ToString.Trim}", _fntValue, Brushes.Black, 80, 170)
        DrawString(canvas, $"{_cap("fc_data"):dd/MM/yyyy}", _fntValue, Brushes.Black, 210, 170)
        DrawString(canvas, $"{_cap("fc_codcli")}", _fntValue, Brushes.Black, 210, 210)
        DrawString(canvas, $"N.I.F.: {_cap("fc_nifcli")}", _fntValue, Brushes.Black, 390, 290)

        DrawString(canvas, $"{_cap("fc_suref")}".Trim, _fntValue, Brushes.Black, 80, 285)

        ' Albarans facturats

        For Each r As DataRow In _lin.InternalDataTable.Rows
            If Not albs.Contains(r("fl_anumero")) Then
                albs.Add(r("fl_anumero"))
            End If
        Next

        If albs.Count > 0 Then
            numAlbs = ""
            For Each s As String In albs
                numAlbs += "-" + s
            Next
            numAlbs = numAlbs.Substring(1)
            DrawString(canvas, numAlbs, _fntValue, Brushes.Black, 60, 250)
        End If

        ' Cleanup

        CurY = 340

    End Sub

    Private Function DrawItemBoxHeader(canvas As Graphics, pCurX As Integer, pCurY As Integer, width As Integer, height As Integer, delta As Integer, title As String, titleAlign As StringFormat, value As String, valueAlign As StringFormat, lastItem As Boolean) As Integer

        DrawString(canvas, title, _fntTitle, Brushes.Black, New RectangleF(pCurX, pCurY, width, delta), titleAlign)
        DrawString(canvas, value, _fntHdrValue, Brushes.Black, New RectangleF(pCurX, pCurY + delta, width, height - delta), valueAlign)
        If Not lastItem Then
            DrawLine(canvas, Pens.Black, pCurX + width, pCurY, pCurX + width, pCurY + height)
        End If

        Return pCurX + width

    End Function

    Private Function FillDetail(canvas As Graphics) As Boolean

        ' Imprimeix texte pendent de la última linea de la pagina
        If _textLeft <> String.Empty Then
            If Not DrawTextLeft(canvas) Then
                Return True
            End If
        End If

        If _ampliacioDescripcio <> String.Empty Then
            _textLeft = _ampliacioDescripcio
            _ampliacioDescripcio = String.Empty
            SettingsLeft(15, _fntLine, 1)
            If Not DrawTextLeft(canvas) Then
                Return True
            End If
        End If

        If Not _linesLeft Then
            Return False
        End If

        Do

            Try
                DrawLinFactura(canvas)
            Catch ex As Exception
                MessageBox.Show($"Error al imprimir linea de factura")
            End Try

            ' AmpliacioDescripcio = lin("fl_ampart").ToString.Trim

            _linesLeft = _lin.Read

            If _textLeft <> String.Empty Then
                SettingsLeft(5, _fntLine, 1)
                If Not DrawTextLeft(canvas) Then
                    Return True
                End If
            End If

            If _ampliacioDescripcio <> String.Empty Then
                _textLeft = _ampliacioDescripcio
                _ampliacioDescripcio = String.Empty
                SettingsLeft(15, _fntLine, 1)
                If Not DrawTextLeft(canvas) Then
                    Return True
                End If
            End If

            If Not _linesLeft Then
                Return False
            End If

            If CurY > _bottomYLines - _deltaLine Then
                Return True
            End If

        Loop

    End Function

    Private Sub DrawLinFactura(canvas As Graphics)
        Dim lCurX As Integer
        Dim i As Integer
        Dim import As Decimal = 0D
        Dim strImport As String = String.Empty
        Dim importUnitari As Decimal = 0D
        Dim unitats As Decimal = 0D
        Dim mides As String
        Dim unitatPreu As String
        Dim strUnitats As String
        Dim strImportUnitari As String = String.Empty
        Dim descripcio As String

        Dim fittedChars As Integer
        Dim fittedLines As Integer

        unitats = CNull(_lin("fl_quant"), 0D)
        importUnitari = CNull(_lin("fl_prart"), 0D)
        import = CNull(_lin("fl_import"), 0D)

        strUnitats = $"{unitats:N0}"

        If String.IsNullOrWhiteSpace($"{_lin("fl_desart")}") Then
            mides = String.Empty
            canvas.MeasureString(_lin("fl_ampart").ToString.Trim, _fntLine, New SizeF(350, _deltaLine), New StringFormat, fittedChars, fittedLines)
            If fittedLines = 1 Then
                descripcio = $"{_lin("fl_ampart")}".Trim
                _textLeft = ""
            Else
                descripcio = $"{_lin("fl_ampart")}".Substring(0, fittedChars)
                _textLeft = $"{_lin("fl_ampart")}".Substring(fittedChars)
            End If
        Else
            mides = $"{_lin("fl_desart")}".Substring(0, 15)
            descripcio = $"{_lin("fl_desart")}".Substring(15).TrimEnd
            _textLeft = $"{_lin("fl_ampart")}".Trim
        End If

        If CNull(_lin("fl_uenvas"), 0) = 1 Then
            unitatPreu = "u."
        Else
            unitatPreu = "m."
        End If

        If importUnitari <> 0D Then strImportUnitari = $"{importUnitari:N3}" Else strImportUnitari = String.Empty
        If import <> 0D Then strImport = $"{import:C}" Else strImport = String.Empty

        DrawString(canvas, strUnitats, _fntLine, Brushes.Black, New RectangleF(30, CurY, 70, _deltaLine), _sfFar)
        DrawString(canvas, mides, _fntLine, Brushes.Black, New RectangleF(130, CurY, 60, _deltaLine), _sfNear)
        DrawString(canvas, descripcio, _fntLine, Brushes.Black, New RectangleF(210, CurY, 350, _deltaLine), _sfNear)
        DrawString(canvas, strImportUnitari, _fntLine, Brushes.Black, New RectangleF(570, CurY, 70, _deltaLine), _sfFar)
        DrawString(canvas, unitatPreu, _fntLine, Brushes.Black, New RectangleF(637, CurY, 20, _deltaLine), _sfNear)
        DrawString(canvas, strImport, _fntLine, Brushes.Black, New RectangleF(700, CurY, 80, _deltaLine), _sfFar)

        CurY += _deltaLine

    End Sub

    Private Sub FillSummary(canvas As Graphics)
        Const lMarginBox As Integer = 5

        Dim import As Decimal
        Dim strImport As String = String.Empty
        Dim baseImponible As Decimal
        Dim strBaseImponible As String = String.Empty
        Dim tpcIva As Decimal
        Dim strTpcIva As String = String.Empty
        Dim importeIva As Decimal
        Dim strImporteIva As String = String.Empty
        Dim importTotal As Decimal
        Dim strImportTotal As String = String.Empty
        Dim pagament As String
        Dim importVenciment As Decimal

        import = CNull(_cap("fc__bases"), 0D)
        baseImponible = CNull(_cap("fc__base1"), 0D) + CNull(_cap("fc__base2"), 0D) + CNull(_cap("fc__base3"), 0D)
        tpcIva = CNull(_cap("fc__tpci1"), 0D)
        importeIva = CNull(_cap("fc__civa1"), 0D) + CNull(_cap("fc__civa2"), 0D) + CNull(_cap("fc__civa3"), 0D)
        importTotal = CNull(_cap("fc__total"), 0D)

        If import <> 0D Then strImport = $"{import:C}" Else strImport = String.Empty
        If baseImponible <> 0D Then strBaseImponible = $"{baseImponible:C}" Else strImport = String.Empty
        If tpcIva <> 0D Then strTpcIva = $"{tpcIva:N0}%" Else strImport = String.Empty
        If importeIva <> 0D Then strImporteIva = $"{importeIva:C}" Else strImport = String.Empty
        If importTotal <> 0D Then strImportTotal = $"{importTotal:C}" Else strImport = String.Empty


        CurX = _bodyLeftOffset
        CurY = _topYTotals

        DrawString(canvas, strImport, _fntLine, Brushes.Black, New RectangleF(CurX, CurY, _totalsWidth(0) - lMarginBox * 2, _deltaLine), _sfCenter)
        CurX += _totalsWidth(0) '  DrawString(Canvas, strDescompte, fntLine, Brushes.Black, New RectangleF(CurX, CurY, TotalsWidth(1) - marginBox * 2, deltaLine), sfCenter)
        CurX += _totalsWidth(1) : DrawString(canvas, strBaseImponible, _fntLine, Brushes.Black, New RectangleF(CurX, CurY, _totalsWidth(2) - lMarginBox * 2, _deltaLine), _sfCenter)
        CurX += _totalsWidth(2) : DrawString(canvas, strTpcIva, _fntLine, Brushes.Black, New RectangleF(CurX, CurY, _totalsWidth(3) - lMarginBox * 2, _deltaLine), _sfCenter)
        CurX += _totalsWidth(3) : DrawString(canvas, strImporteIva, _fntLine, Brushes.Black, New RectangleF(CurX, CurY, _totalsWidth(4) - lMarginBox * 2, _deltaLine), _sfCenter)
        CurX += _totalsWidth(4) : DrawString(canvas, strImportTotal, _fntLine, Brushes.Black, New RectangleF(CurX, CurY, _totalsWidth(5) - lMarginBox * 2, _deltaLine), _sfCenter)

        If True Then
            'Venciments

            If CNull(_cap("fc__total"), 0D) > 0 Then
                ' Si es una devolució no pintem res.
                ' Forma de pagament
                CurX = 70
                CurY = _topYVenciments
                pagament = _dbaFpg.GetFormaPagament(CNull(_cap("fc_fpago")))
                DrawString(canvas, pagament, _fntLine, Brushes.Black, CurX, CurY)

                CurY += 20
                pagament = _cap("fc_banc").ToString.Trim

                DrawString(canvas, pagament, _fntLine, Brushes.Black, CurX, CurY)
                If Not IsNullOrEmptyValue(_cap("fc_iban").ToString) Then
                    pagament = "IBAN: " + Utils.Transform(_cap("fc_iban").ToString, "XX99-9999-9999-9999-9999-9999")
                    DrawString(canvas, pagament, _fntLine, Brushes.Black, CurX, CurY + _deltaLine)
                End If

                ' Venciments

                CurY = _topYVenciments
                If _reb.Rows.Count > 0 Then
                    For Each r As DataRow In _reb.Rows
                        importVenciment = CNull(r("re_import"), 0D) + CNull(r("re_devol"), 0D) + CNull(r("re_cobrat"), 0D)
                        DrawString(canvas, $"{r("re_dvto"):dd/MM/yyyy}", _fntLine, Brushes.Black, New RectangleF(440, CurY, 80, _deltaLine), _sfNear)
                        DrawString(canvas, $"{importVenciment:C}", _fntLine, Brushes.Black, New RectangleF(550, CurY, 90, _deltaLine), _sfFar)
                        CurY += 20
                    Next
                End If

            End If
        End If

        DrawString(canvas, $"{_cap("fc_nota")}", _fntLine, Brushes.Black, New RectangleF(CurX, 970, 720, _deltaLine), _sfNear)

    End Sub

    Private Sub SettingsLeft(pTextLeft As String, pOffsetLeft As Integer, pFontLeft As Font, pColumnLeft As Integer)
        _textLeft = pTextLeft
        _offsetLeft = pOffsetLeft
        _offsetTop = 0
        _fntLeft = pFontLeft
        _columnLeft = pColumnLeft
    End Sub

    Private Sub SettingsLeft(pOffsetLeft As Integer, pFontLeft As Font, pColumnLeft As Integer)
        _offsetTop = 0
        _offsetLeft = pOffsetLeft
        _fntLeft = pFontLeft
        _columnLeft = pColumnLeft
    End Sub

    Private Sub SettingsLeft(pOffsetTop As Integer, pOffsetLeft As Integer, pFontLeft As Font, pColumnLeft As Integer)
        _offsetTop = pOffsetTop
        _offsetLeft = pOffsetLeft
        _fntLeft = pFontLeft
        _columnLeft = pColumnLeft
    End Sub

    Private Function DrawTextLeft(canvas As Graphics) As Boolean
        Dim szFree As SizeF
        Dim szUsed As SizeF
        Dim sf As New StringFormat
        Dim i As Integer

        sf.Alignment = StringAlignment.Near
        sf.LineAlignment = StringAlignment.Near

        Dim linesFitted As Integer
        Dim charsFitted As Integer

        'szFree = New SizeF(ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY)
        szFree = New SizeF(_columnWidth(_columnLeft) - _offsetLeft, (_bottomYLines - 4) - CurY)
        szUsed = canvas.MeasureString(_textLeft, _fntLeft, szFree, sf, charsFitted, linesFitted)

        CurX = _bodyLeftOffset + 5
        i = 0
        Do While i < _columnLeft
            CurX += _columnWidth(i)
            i += 1
        Loop

        'DrawString(Canvas, TextLeft, fntLeft, Brushes.Black, New RectangleF(CurX + OffsetLeft, CurY, ColumnWidth(ColumnLeft) - OffsetLeft - marginBox * 2, (BottomYLines - 4) - CurY), sf)
        DrawString(canvas, _textLeft, _fntLeft, Brushes.Black, New RectangleF(CurX + _offsetLeft, CurY + _offsetTop, 650 - _offsetLeft, (_bottomYLines - 4) - CurY + _offsetTop), sf)

        If charsFitted < _textLeft.Length Then
            _textLeft = _textLeft.Remove(0, charsFitted)
        Else
            _textLeft = String.Empty
        End If

        CurY += (CInt(szUsed.Height) + _offsetTop)

        Return (_textLeft.Length = 0)

    End Function

    Protected Overrides Sub Print2Excel(FileName As String)

    End Sub

    Private Sub clsImpresFactura_GetDataSource() Handles Me.GetDataSource

        Try

            _cap = _dbaCap.GetFactura(OrigenDades, SerieFactura, AnyFactura, NumeroFactura)
            _reb = _dbaReb.GetRebuts(SerieFactura, AnyFactura, NumeroFactura)

            _lin = New csTableReader(_dbaLin.GetLiniesFactura(OrigenDades, SerieFactura, AnyFactura, NumeroFactura))

            _linesLeft = _lin.Read

            _dataLoaded = True

        Catch ex As Exception
            _dataLoaded = False
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_gst_fac " + ex.Message + vbCrLf +
                                    $"Any: {AnyFactura}, Número: {NumeroFactura}")
        End Try

        _nowPrinting = NowPrintingEnum.LinFactura

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        _fntTitle.Dispose()
        _fntValue.Dispose()
        _fntItalic.Dispose()
        _fntHdrValue.Dispose()
        _fntLine.Dispose()
        _fntAnagrama.Dispose()
        _fntInscrita.Dispose()
        _fntGrossa.Dispose()

    End Sub

#Region " Print "

    Private Sub Execute()

        EmpresaName = AppData.CurrentEmpresaName
        ReportName = "Factura"
        ReportID = "R90EXP0003C"

        CustomId = "1" 'Gestions

        LayoutOffset = LayoutOffsetEnum.OneThird
        PageNumbering = PageNumberEnum.PageNofM

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
