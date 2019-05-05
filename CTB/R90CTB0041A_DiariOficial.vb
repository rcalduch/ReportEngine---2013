Imports csAppData
Imports csUtils
Imports csUtils.Utils

Public Class R90CTB0041A_DiariOficial
    Inherits ReportBaseClass
    Private dba As New C00_ctb_reports

    Public Overrides Sub Execute(workInfo As String)
        Dim rpt As New R90CTB0041C_DiariOficial

        Dim PrinterID As Integer
        Dim numCopies As Integer
        Dim Output As String

        Dim ReportOutput As TipusEnviamentDocumentEnum

        Dim params As FetchXmlParameter

        Try

            rpt.CustomID = My.Settings.ClientCustom
            rpt.PageNumbering = csRpt.PageNumberEnum.PageNofM

            If String.IsNullOrEmpty(workInfo) Then
                rpt.Print()
                Return
            End If

            params = New FetchXmlParameter(workInfo)

            PrinterID = CNull(params.GetValue("PrinterID"), 0)

            rpt.CodiLlistat = CNull(params.GetValue("CodiLlistat"))
            rpt.NomClient = CNull(params.GetValue("NomClient"))
            rpt.CodiEmpresa = CNull(params.GetValue("CodiEmpresa"))
            rpt.NomEmpresa = CNull(params.GetValue("NomEmpresa"))
            rpt.TitolLlistat = CNull(params.GetValue("TitolLlistat")).Replace("DIARIO CONTABLE", "DIARI COMPTABLE")
            rpt.DataLlistat = CNull(params.GetValue("DataLlistat"))

            numCopies = CInt(params.GetValue("Copies"))
            Output = params.GetValue("Output")

            Select Case Output.ToLower
                Case "mail", "email", "e-mail"
                    ReportOutput = TipusEnviamentDocumentEnum.Email
                Case "pdf"
                    ReportOutput = TipusEnviamentDocumentEnum.Pdf
                Case "paper", "impressora"
                    ReportOutput = TipusEnviamentDocumentEnum.Impressora
                Case Else
                    ReportOutput = TipusEnviamentDocumentEnum.Impressora
            End Select

            ' email info

            Select Case ReportOutput

                Case TipusEnviamentDocumentEnum.Pdf, TipusEnviamentDocumentEnum.Email
                    rpt.Destinacio = csRpt.ReportDestinationEnum.PDF
                    rpt.pdfShowSaveDialog = False
                    rpt.pdfPathAndFileName = String.Empty
                    rpt.pdfDirectori = My.Settings.OutputDirPDF
                    rpt.pdfNomFitxer = String.Format("{0} - {1}.PDF", rpt.NomEmpresa, rpt.TitolLlistat.Replace("/", "_"))

                Case TipusEnviamentDocumentEnum.Impressora
                    rpt.Destinacio = csRpt.ReportDestinationEnum.Printer
                    rpt.ShowPrintDialog = False
                    rpt.SetDefaultPrinter()

            End Select

            rpt.Print()

        Catch ex As Exception

            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90CTB0041A_DiariOficial " + ex.Message)

        End Try

    End Sub

End Class

Public Class R90CTB0041C_DiariOficial
    Inherits csRpt

    Private FirstPage As Boolean
    Private mIdiomaID As Integer
    Private mCustomID As String

    Private TraceLocation As String

    Private dbaRpt As New C00_ctb_reports

    Private lin As csTableReader

    Private ColumnWidth() As Integer
    Private DetailColumns As Integer

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
    Private LinesLeft As Boolean
    Private ColumnGap As Integer

    Private fntTitle As Font
    Private fntLine As Font
    Private fntLineB As Font

    Private deltaLine As Integer

    Private Enum NowPrintingEnum
        HeaderRows
        SaldoAnterior
        Extracte
        SaldoFinal
    End Enum

    Private NowPrinting As NowPrintingEnum

    Public Shadows Property CustomID As String
        Get
            Return mCustomID
        End Get
        Set
            mCustomID = value
        End Set
    End Property

#Region " Properties "

#Region " Propietat Llistat "

    Public CodiLlistat As String
    Public NomClient As String
    Public CodiEmpresa As String
    Public NomEmpresa As String
    Public TitolLlistat As String
    Public DataLlistat As String

#End Region

#Region " Propietats llistat "

    Private mDestinacio As ReportDestinationEnum
    Public Property Destinacio As csRpt.ReportDestinationEnum
        Get
            Return mDestinacio
        End Get
        Set
            Destination = value
            mDestinacio = value
        End Set
    End Property

    Private mCopies As Integer
    Public Property Copies As Integer
        Get
            Return mCopies
        End Get
        Set
            mCopies = value
        End Set
    End Property

    Private mFitxerPdf As String
    Public Property FitxerPdf As String
        Get
            Return mFitxerPdf
        End Get
        Set
            mFitxerPdf = value
        End Set
    End Property

#End Region

#End Region

    Public Sub New()
        MyBase.New()

        PrtSettings.DefaultPageSettings.Landscape = False

        DetailColumns = 9

        WidthLines = 750
        BodyLeftOffset = 25
        ColumnGap = 5

        ReDim ColumnWidth(DetailColumns - 1)

        BottomYLines = 1100

        ColumnWidth(0) = 55
        ColumnWidth(1) = 35
        ColumnWidth(2) = 50
        ColumnWidth(3) = 90
        ColumnWidth(4) = 160
        ColumnWidth(5) = 70
        ColumnWidth(6) = 60
        ColumnWidth(7) = 140
        ColumnWidth(8) = WidthLines

        For i As Integer = 0 To DetailColumns - 2
            ColumnWidth(DetailColumns - 1) -= ColumnWidth(i)
            ColumnWidth(i) -= ColumnGap
        Next

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

        fntTitle = New Font("Arial Narrow", 8, FontStyle.Bold)
        fntLine = New Font("Arial Narrow", 7, FontStyle.Regular)
        fntLineB = New Font("Arial Narrow", 7, FontStyle.Bold)

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

    Public Overrides Function DrawPage(Canvas As System.Drawing.Graphics) As Boolean
        Dim hasMoreData As Boolean

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

        hasMoreData = FillDetail(Canvas)

        Return hasMoreData

    End Function

    Private Sub PrintHeader(Canvas As System.Drawing.Graphics)
        Dim curX As Integer
        Dim col As Integer

        CurY = 20
        DrawLine(Canvas, BodyLeftOffset, CurY, BodyLeftOffset + WidthLines, CurY, 1, Color.Black)
        CurY += 5
        DrawString(Canvas, TitolLlistat, fntTitle, Brushes.Black, New RectangleF(BodyLeftOffset, CurY, WidthLines, 12), sfCenter)
        DrawString(Canvas, NomClient, fntTitle, Brushes.Black, New RectangleF(BodyLeftOffset, CurY, 150, 12), sfNear)
        DrawString(Canvas, String.Format("Pàgina {0}", CurrentPage), fntTitle, Brushes.Black, New RectangleF(600, CurY, 175, 12), sfFar)
        CurY += 17
        DrawLine(Canvas, BodyLeftOffset, CurY, BodyLeftOffset + WidthLines, CurY, 3, Color.Black)

        CurY += 20

        curX = BodyLeftOffset
        col = 0

        curX += ColumnWidth(col) + ColumnGap : col += 1
        curX += ColumnWidth(col) + ColumnGap : col += 1
        curX += ColumnWidth(col) + ColumnGap : col += 1
        ' DrawString(Canvas, "--- Compte ---", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfCenter)
        DrawOverTitle(Canvas, curX, CurY, ColumnWidth(col), "Compte")

        CurY += 14
        curX = BodyLeftOffset
        col = 0

        DrawString(Canvas, "Asment", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Línia", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Data", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Deure", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        DrawString(Canvas, "Haver", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Titol", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Deure", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Haver", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Descripció asentament", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Sèrie-Document", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)

        CurY += 14 + 2

        col = 0
        curX = BodyLeftOffset

        For Each w As Integer In ColumnWidth
            DrawLine(Canvas, curX, CurY, curX + ColumnWidth(col), CurY, 1, Color.Black)
            curX += ColumnWidth(col) + ColumnGap : col += 1
        Next

        CurY += 7

    End Sub

    Private Sub DrawOverTitle(Canvas As Graphics, x As Integer, y As Integer, w As Integer, t As String)
        Dim string_size As SizeF
        Dim width_line As Integer
        Dim width_text As Integer

        string_size = Canvas.MeasureString(t, fntLineB)
        width_text = CInt(string_size.Width)

        width_line = (w - width_text - ColumnGap * 2) \ 2

        DrawLine(Canvas, Pens.Black, x, y + deltaLine \ 2, x + width_line, y + deltaLine \ 2)
        DrawString(Canvas, t, fntLineB, Brushes.Black, x + width_line + ColumnGap, y)
        DrawLine(Canvas, Pens.Black, x + width_line + width_text + ColumnGap * 2, y + deltaLine \ 2, x + w, y + deltaLine \ 2)

    End Sub

    Private Function FillDetail(Canvas As System.Drawing.Graphics) As Boolean

        Do

            Try
                DrawLinDiariOficial(Canvas)
            Catch ex As Exception
                MessageBox.Show("Error al imprimir linea d'extracte comptes")
            End Try

            LinesLeft = lin.Read

            If Not LinesLeft Then
                Return False
            End If

            If CurY > BottomYLines - deltaLine Then
                Return True
            End If

        Loop

    End Function

    Private Sub DrawLinDiariOficial(Canvas As System.Drawing.Graphics)
        Dim curX As Integer
        Dim col As Integer
        Dim Titol As String
        Dim SerDoc As String

        curX = BodyLeftOffset
        col = 0

        Titol = lin("Titulo").ToString.Trim

        If Titol = String.Empty Then
            CurY += deltaLine
            lin.Read()
            Titol = lin("Titulo").ToString.Trim
        End If

        If Titol.StartsWith("TOTAL DIARIO") And CNull(lin("Asiento"), 0) = 0 Then
            Titol = "TOTAL DIARI"

            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, Titol, fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 16), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_debe"), 2, True), fntTitle, Brushes.Black, New RectangleF(curX - 20, CurY, ColumnWidth(col) + 20, 16), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            CurY += deltaLine
            DrawString(Canvas, fmtValueToStr(lin("Imp_Haber"), 2, True), fntTitle, Brushes.Black, New RectangleF(curX - 20, CurY, ColumnWidth(col) + 20, 16), sfFar)

        ElseIf Titol.StartsWith("DIFERENCIA") Then

            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, "DIFERÈNCIA", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_debe"), 2, True), fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 16), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Haber"), 2, True), fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 16), sfFar)

        Else

            DrawString(Canvas, fmtValueToStr(lin("Asiento"), 0, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Linea"), 0, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Fecha"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1

            If String.IsNullOrWhiteSpace(CNull(lin("cta_debe"))) Then
                DrawString(Canvas, fmtValueToStr(lin("cta_haber"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            Else
                DrawString(Canvas, fmtValueToStr(lin("cta_debe"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            End If
            curX += ColumnWidth(col) + ColumnGap : col += 1

            DrawString(Canvas, fmtValueToStr(lin("Titulo"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Debe"), 2, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Haber"), 2, True), fntLine, Brushes.Black, New RectangleF(curX - 10, CurY, ColumnWidth(col) + 10, 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Descripcio"), 2, False), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1

            If lin("serie").ToString.Trim = String.Empty Then
                SerDoc = fmtValueToStr(lin("Documento"), 0, True)
            Else
                SerDoc = lin("Serie").ToString.Trim + "-" + fmtValueToStr(lin("Documento"), 0, True)
            End If
            DrawString(Canvas, SerDoc, fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)

        End If

        CurY += deltaLine

    End Sub

    Private Function fmtValueToStr(value As Object, decimals As Integer, ZeroAsEmpty As Boolean) As String
        Dim retValue As String

        If value.GetType = GetType(Date) Then
            retValue = String.Format("{0:dd/MM/yy}", value)
        ElseIf value.GetType = GetType(Integer) Then
            If CNull(value, 0) = 0 Then
                retValue = String.Empty
            Else
                retValue = String.Format("{0}", value)
            End If
        ElseIf value.GetType = GetType(Decimal) Then
            If CNull(value, 0D) = 0D And ZeroAsEmpty Then
                retValue = String.Empty
            Else
                retValue = String.Format("{0:N" + decimals.ToString + "}", value)
            End If
        Else
            retValue = String.Format("{0}", value)
        End If

        Return retValue

    End Function

    Protected Overrides Sub Print2Excel(FileName As String)

    End Sub

    Private Sub clsImpresFactura_GetDataSource() Handles Me.GetDataSource

        Try

            lin = New csTableReader(dbaRpt.get_diari_oficial)

            LinesLeft = lin.Read

            DataLoaded = True

        Catch ex As Exception
            DataLoaded = False
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_ctb_ofi " + ex.Message)
        End Try

        NowPrinting = NowPrintingEnum.Extracte

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()

        fntTitle.Dispose()
        fntLine.Dispose()

    End Sub

#Region " Print "

    Private Sub Execute()

        EmpresaName = AppData.CurrentEmpresaName
        ReportName = "Diari Oficial"
        ReportID = "R90CTB0041C"

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

            DebugLog("R90CTB0041C: " + ex.Message)

        Finally


        End Try

    End Sub

#End Region

End Class
