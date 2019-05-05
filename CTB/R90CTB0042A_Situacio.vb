Imports csAppData
Imports csUtils
Imports csUtils.Utils

Public Class R90CTB0042A_Situacio
    Inherits ReportBaseClass
    Private dba As New C00_ctb_reports

    Public Overrides Sub Execute(workInfo As String)
        Dim rpt As New R90CTB0042C_Situacio

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
            rpt.TitolLlistat = CNull(params.GetValue("TitolLlistat")).Replace("BALANCE DE SITUACION", "BALANÇ DE SITUACIÓ")
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

            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90CTB0042A_SumesISaldos " + ex.Message)

        End Try

    End Sub

End Class

Public Class R90CTB0042C_Situacio

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
    Private BodyExtraOffset As Integer
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

        DetailColumns = 4

        WidthLines = 750
        BodyLeftOffset = 25
        BodyExtraOffset = 150
        ColumnGap = 5

        ReDim ColumnWidth(DetailColumns - 1)

        BottomYLines = 1120

        ColumnWidth(0) = 80
        ColumnWidth(1) = 190
        ColumnWidth(2) = 80
        ColumnWidth(3) = 80

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

        curX = BodyLeftOffset + BodyExtraOffset
        col = 0


        ' DrawString(Canvas, "COMPTE", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        ' DrawString(Canvas, "TÍTOL", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "SALDO INICIAL", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "SALDO FINAL", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)

        CurY += 14 + 2

        col = 0
        curX = BodyLeftOffset + BodyExtraOffset

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

    Private Function FillDetail(Canvas As Graphics) As Boolean
        Dim ForceNewPage As Boolean

        Do

            Try
                ForceNewPage = DrawLinBalansSit(Canvas)
            Catch ex As Exception
                MessageBox.Show("Error al imprimir linea balanç SiS")
            End Try

            LinesLeft = lin.Read

            If Not LinesLeft Then
                Return False
            End If

            If ForceNewPage Then
                Return True
            End If

            If CurY > BottomYLines - deltaLine Then
                Return True
            End If

        Loop

    End Function

    Private Function DrawLinBalansSit(Canvas As System.Drawing.Graphics) As Boolean
        Dim curX As Integer
        Dim col As Integer
        Dim Titol As String
        Dim Compte As String
        Dim ForceNewPage As Boolean
        Dim fnt2Use As Font

        ForceNewPage = False

        curX = BodyLeftOffset + BodyExtraOffset
        col = 0

        Titol = lin("Titol").ToString.Trim
        Compte = lin("Compte").ToString.Trim

        If Titol.StartsWith("COMPTES") And String.IsNullOrWhiteSpace(Compte) Then

            DrawString(Canvas, Titol, fntTitle, Brushes.Black, New RectangleF(curX, CurY, 150, 16), sfNear)
            CurY += deltaLine \ 2

        ElseIf (Titol.StartsWith("TOTAL") Or Titol.StartsWith("PREVISIO RESULTATS:")) And String.IsNullOrWhiteSpace(Compte) Then

            CurY += 5

            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Titol"), 0, True), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("sdo_inici"), 2, True), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("sdo_final"), 2, True), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)

            If Titol.StartsWith("TOTAL ACTIU") Then
                ForceNewPage = True
            End If

        Else

            If Compte.Length = 1 Then
                fnt2Use = fntLineB
                CurY += 5
            Else
                fnt2Use = fntLine
            End If

            DrawString(Canvas, fmtValueToStr(lin("Compte"), 0, True), fnt2Use, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Titol"), 0, True), fnt2Use, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("sdo_inici"), 2, False), fnt2Use, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("sdo_final"), 2, False), fnt2Use, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)

            If Compte.Length = 1 Then
                CurY += 5
            End If

        End If

        CurY += deltaLine

        Return ForceNewPage

    End Function

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

            lin = New csTableReader(dbaRpt.get_balans_situacio)

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
        ReportName = "Balanç Situació"
        ReportID = "R90CTB0042C_SIT"

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

            DebugLog("R90CTB0042C: " + ex.Message)

        Finally


        End Try

    End Sub

#End Region

End Class

