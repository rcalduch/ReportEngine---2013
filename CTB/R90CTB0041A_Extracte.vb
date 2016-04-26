Imports csAppData
Imports csUtils
Imports csUtils.Utils

Public Class R90CTB0041A_Extracte
    Inherits ReportBaseClass
    Private dba As New C00_ctb_reports

    Public Overrides Sub Execute(workInfo As String)
        Dim rpt As New R90CTB0041C_Extracte

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
            rpt.TitolLlistat = CNull(params.GetValue("TitolLlistat"))
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

            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90CTB0041A_Extracte " + ex.Message)

        End Try

    End Sub

End Class


Public Class R90CTB0041C_Extracte
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

    Public Shadows Property CustomID() As String
        Get
            Return mCustomID
        End Get
        Set(ByVal value As String)
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

        ColumnWidth(0) = 70
        ColumnWidth(1) = 60
        ColumnWidth(2) = 45
        ColumnWidth(3) = 210
        ColumnWidth(4) = 80
        ColumnWidth(5) = 80
        ColumnWidth(6) = 80
        ColumnWidth(7) = 50
        ColumnWidth(8) = WidthLines

        For i As Integer = 0 To DetailColumns - 2
            ColumnWidth(8) -= ColumnWidth(i)
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

        fntTitle = New Font("Arial Narrow", 10, FontStyle.Bold)
        fntLine = New Font("Arial", 8, FontStyle.Regular)
        fntLineB = New Font("Arial", 8, FontStyle.Bold)

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

    Private Sub PrintHeader(ByVal Canvas As System.Drawing.Graphics)
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

        DrawString(Canvas, "Data", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Asment", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Línia", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Descripcio", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Deure", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Haver", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Saldo", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Sèrie", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
        curX += ColumnWidth(col) + ColumnGap : col += 1
        DrawString(Canvas, "Document", fntTitle, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)

        CurY += 14 + 2

        col = 0
        curX = BodyLeftOffset

        For Each w As Integer In ColumnWidth
            DrawLine(Canvas, curX, CurY, curX + ColumnWidth(col), CurY, 1, Color.Black)
            curX += ColumnWidth(col) + ColumnGap : col += 1
        Next

        CurY += 7

    End Sub

    Private Function FillDetail(ByVal Canvas As System.Drawing.Graphics) As Boolean

        Do

            Try
                DrawLinExtracte(Canvas)
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

    Private Sub DrawLinExtracte(ByVal Canvas As System.Drawing.Graphics)
        Dim curX As Integer
        Dim col As Integer
        Dim Descripcio As String
        Dim Compte As String
        Dim Titol As String

        curX = BodyLeftOffset
        col = 0

        Descripcio = lin("Descripcio").ToString.Trim

        If Descripcio = String.Empty Then
            CurY += deltaLine
            lin.Read()
            Descripcio = lin("Descripcio").ToString.Trim
        End If

        If Descripcio.StartsWith("CUENTA: ") And CNull(lin("Asiento"), 0) = 0 Then
            Compte = Descripcio.Replace("CUENTA", "COMPTE")
            lin.Read()
            Titol = lin("Descripcio").ToString.Trim

            DrawString(Canvas, Compte, fntTitle, Brushes.Black, New RectangleF(curX, CurY, 170, 16), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, Titol, fntTitle, Brushes.Black, New RectangleF(curX, CurY, 270, 16), sfNear)

            CurY += deltaLine

        ElseIf Descripcio.StartsWith("SALDO INICIAL") Then

            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, "SALDO INICIAL", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Saldo"), 2, False), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)

        ElseIf Descripcio.StartsWith("SALDO MOVNTOS") Then

            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, "SALDO MOVIMENTS ANTERIORS", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Debe"), 2, True), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Haber"), 2, True), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Saldo"), 2, False), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)

        ElseIf Descripcio.StartsWith("SALDO FINAL") Then

            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, "SALDO FINAL", fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Saldo"), 2, False), fntLineB, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 14), sfFar)

        Else

            DrawString(Canvas, fmtValueToStr(lin("Fecha"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Asiento"), 0, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Linea"), 0, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Descripcio"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Debe"), 2, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Haber"), 2, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Imp_Saldo"), 2, False), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfFar)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Serie"), Nothing, Nothing), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)
            curX += ColumnWidth(col) + ColumnGap : col += 1
            DrawString(Canvas, fmtValueToStr(lin("Documento"), 0, True), fntLine, Brushes.Black, New RectangleF(curX, CurY, ColumnWidth(col), 12), sfNear)

        End If

        CurY += deltaLine

    End Sub

    Private Function fmtValueToStr(value As Object, decimals As Integer, ZeroAsEmpty As Boolean) As String
        Dim retValue As String

        If value.GetType = GetType(Date) Then
            retValue = String.Format("{0:dd/MM/yyyy}", value)
        ElseIf value.GetType = GetType(Integer) Then
            If CNull(value, 0) = 0 Then
                retValue = String.Empty
            Else
                retValue = String.Format("{0}", value)
            End If
        ElseIf value.GetType = GetType(Decimal) Then
            If CNull(value, 0) = 0D And ZeroAsEmpty Then
                retValue = String.Empty
            Else
                retValue = String.Format("{0:N" + decimals.ToString + "}", value)
            End If
        Else
            retValue = String.Format("{0}", value)
        End If

        Return retValue

    End Function

    Protected Overrides Sub Print2Excel(ByVal FileName As String)

    End Sub

    Private Sub clsImpresFactura_GetDataSource() Handles Me.GetDataSource

        Try

            lin = New csTableReader(dbaRpt.get_extracte_comptes)

            LinesLeft = lin.Read

            DataLoaded = True

        Catch ex As Exception
            DataLoaded = False
            DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " C00_ctb_ext " + ex.Message)
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
        ReportName = "Extracte de comtpes"
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
