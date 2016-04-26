Imports System
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.ComponentModel

Namespace csBarcode

  Public Enum BarcodeSymbologies
    EAN13
    DUN14
    I2of5
    Code39
    Code128Auto
    Code128A
    Code128B
    Code128C
    UCCEAN128
    SSCC
    Pdf417
  End Enum

  Public Class csBarCode

#Region " Properties "

    Private m_Data As String
    Private m_Symbol As BarcodeSymbologies
    Private m_Size As Size
    Private m_Image As Bitmap
    Private m_DrawReadableData As Boolean
    Private m_ReadableData As String
    Private m_FontSize As Integer
    Private m_IsDerty As Boolean
    Private m_Width As Integer
    Private m_Height As Integer

    Property Data() As String
      Get
        Return m_Data
      End Get
      Set(ByVal value As String)
        m_Data = value
        m_IsDerty = True
      End Set
    End Property

    <Obsolete()> Property Symbology() As BarcodeSymbologies
      Get
        Return m_Symbol
      End Get
      Set(ByVal value As BarcodeSymbologies)
        m_Symbol = value
        m_IsDerty = True
      End Set
    End Property

    Property Simbology() As BarcodeSymbologies
      Get
        Return m_Symbol
      End Get
      Set(ByVal value As BarcodeSymbologies)
        m_Symbol = value
        m_IsDerty = True
      End Set
    End Property

    Property DrawReadableData() As Boolean
      Get
        Return m_DrawReadableData
      End Get
      Set(ByVal value As Boolean)
        m_DrawReadableData = value
        m_IsDerty = True
      End Set
    End Property

    Property ReadableData() As String
      Get
        If String.IsNullOrEmpty(m_ReadableData) Then
          If DrawReadableData Then
            Return (m_Data)
          Else
            Return String.Empty
          End If
        Else
          Return (m_ReadableData)
        End If
      End Get
      Set(ByVal value As String)
        m_ReadableData = value
        m_IsDerty = True
      End Set
    End Property

    Property FontSize() As Integer
      Get
        Return m_FontSize
      End Get
      Set(ByVal value As Integer)
        m_FontSize = value
        m_IsDerty = True
      End Set
    End Property

    Property BarcodeSize() As Size
      Get
        Return m_Size
      End Get
      Set(ByVal value As Size)
        m_Size = value
        m_Width = value.Width
        m_Height = value.Height
      End Set
    End Property

    Property Width() As Integer
      Get
        Return m_Width
      End Get
      Set(ByVal value As Integer)
        m_Width = value
      End Set
    End Property

    Property Height() As Integer
      Get
        Return m_Height
      End Get
      Set(ByVal value As Integer)
        m_Height = value
      End Set
    End Property
#End Region

    Public Function BarcodeImage() As Bitmap
      If m_IsDerty Then
        GetBarcodeImage()
      End If
      Return m_Image
    End Function

    Public Function BarcodeImage(ByVal width As Integer, ByVal height As Integer) As Bitmap
      Me.Width = width
      Me.Height = height
      m_IsDerty = True
      Return BarcodeImage()
    End Function

    Private Sub GetBarcodeImage()
      Dim bc As csBarcodeBase
      Select Case Simbology
        Case BarcodeSymbologies.EAN13
          bc = New csEAN13(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.Code128Auto, _
             BarcodeSymbologies.Code128A, _
             BarcodeSymbologies.Code128B, _
             BarcodeSymbologies.Code128C, _
             BarcodeSymbologies.UCCEAN128
          bc = New csCode128(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.Code39
          bc = New csCode39(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.DUN14
          bc = New csDUN14(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.I2of5
          bc = New csI2of5(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.Pdf417
          bc = New csPdf417(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
        Case BarcodeSymbologies.SSCC
          bc = New csSSCC(Simbology, Data, ReadableData, Width, Height, DrawReadableData, FontSize)
      End Select

      m_Image = bc.Image

    End Sub

  End Class

  Friend MustInherit Class csBarcodeBase
    Private m_Data As String
    Private m_Text As String
    Private m_Width As Integer
    Private m_Height As Integer
    Private m_ShowText As Boolean
    Private m_FontSize As Integer
    Private m_Simbology As BarcodeSymbologies

    Property Data() As String
      Get
        Return m_Data
      End Get
      Set(ByVal value As String)
        m_Data = value
      End Set
    End Property

    Property Text() As String
      Get
        Return m_Text
      End Get
      Set(ByVal value As String)
        m_Text = value
      End Set
    End Property

    Property Width() As Integer
      Get
        Return m_Width
      End Get
      Set(ByVal value As Integer)
        m_Width = value
      End Set
    End Property

    Property Height() As Integer
      Get
        Return m_Height
      End Get
      Set(ByVal value As Integer)
        m_Height = value
      End Set
    End Property

    Property ShowText() As Boolean
      Get
        Return m_ShowText
      End Get
      Set(ByVal value As Boolean)
        m_ShowText = value
      End Set
    End Property

    Property FontSize() As Integer
      Get
        Return m_FontSize
      End Get
      Set(ByVal value As Integer)
        m_FontSize = value
      End Set
    End Property

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      m_Data = Data
      m_Text = Text
      m_Width = Width
      m_Height = Height
      m_ShowText = ShowText
      m_FontSize = FontSize
      m_Simbology = Simbology
    End Sub

    Property Simbology() As BarcodeSymbologies
      Get
        Return m_Simbology
      End Get
      Set(ByVal value As BarcodeSymbologies)
        m_Simbology = value
      End Set
    End Property

    'Public Sub New()
    '  'Nothing 
    'End Sub

    Protected Function OneD_DrawPattern(ByVal Pattern As String, ByVal Width As Integer, ByVal Height As Integer) As Bitmap
      ' ens ho pasa en 1/100" i les impressores arriben fins a 1200 ppp

      Dim barWidth As Single
      Dim CurX As Single

      barWidth = Width \ Pattern.Length
      CurX = 0

      Dim bmp As New Bitmap(Width, Height)
      ' ens ho pasa en 1/100" i les impressores arriben fins a 1200 ppp

      Dim gr As Graphics
      gr = Graphics.FromImage(bmp)

      'gr.ScaleTransform(CSng(Width / barWidth * Pattern.Length), 1)

      For Each c As Char In Pattern.ToCharArray
        If c = "1"c Then
          gr.FillRectangle(Brushes.Black, CurX, 0, barWidth, bmp.Height)
        End If
        CurX += barWidth
      Next
      'gr.ResetTransform()

      If m_ShowText Then

        If m_FontSize = 0 Then m_FontSize = 9

        Using fnt As New Font("Arial", m_FontSize, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString(m_Text, fnt, New Point(0, 0), Nothing)
          Dim rds As SizeF = New Size(bmp.Width, CInt(ds.Height))

          gr.FillRectangle(Brushes.White, 0, bmp.Height - rds.Height, rds.Width, rds.Height)
          'gr.DrawRectangle(Pens.Black, 1, bmp.Height - rds.Height + 2, rds.Width - 2, rds.Height - 2)

          If ds.Width > rds.Width Then
            gr.TranslateTransform(0, bmp.Height - rds.Height)
            gr.ScaleTransform(rds.Width / ds.Width, 1)
            gr.DrawString(m_Text, fnt, Brushes.Black, 0, 0)
            gr.ResetTransform()
          Else
            gr.DrawString(m_Text, fnt, Brushes.Black, (rds.Width - ds.Width) / 2, bmp.Height - rds.Height)
          End If

        End Using

      End If
      gr.Dispose()

      Return bmp

    End Function

    Public MustOverride Function Image() As Bitmap

  End Class

  Friend Class csEAN13
    Inherits csBarcodeBase

    Public Overloads Function Image(ByVal pWidth As Integer, ByVal pHeight As Integer) As Bitmap
      Me.Width = pWidth
      Me.Height = pHeight
      Return Me.Image
    End Function

    Public Overrides Function Image() As Bitmap

      Dim bmp As Bitmap
      Dim barWidth As Single
      Dim imageWidth As Single
      Dim Pattern As String

      Dim scaleX As Single
      Dim scaleY As Single

      Dim bmpWidth As Integer
      Dim bmpHeight As Integer

      bmpWidth = Me.Width * 12
      bmpHeight = Me.Height * 12

      bmp = New Bitmap(bmpWidth, bmpHeight)

      Dim gr As Graphics
      gr = Graphics.FromImage(bmp)
      gr.Clear(Color.White)

      Data = EAN13_GetValidCode(Data)
      If Data.Length = 0 Then
        Using pencil As New Pen(Color.Black, 5)
          gr.DrawRectangle(pencil, 5, 5, bmp.Width - 10, bmp.Height - 10)
          gr.DrawLine(pencil, 5, 5, bmp.Width - 10, bmp.Height - 10)
          gr.DrawLine(pencil, 5, bmp.Height - 10, bmp.Width - 10, 5)
        End Using
        Return bmp
      End If

      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic

      ' Suposarem un ratio estandard de que la amplada es 4 vegades l'alçada.

      Pattern = EAN13_GetMask(Data) '=113
      imageWidth = barWidth * Pattern.Length

      ' si no pinta readdable data, el codi ocupara tota la superficie
      ' sino, per l'esquerra deixara lloc per a pintar el primer digit
      ' i per baix, la mitat de l'alsada del numero

      scaleX = CSng(bmpWidth / 113)
      scaleY = CSng(bmpHeight / 100)

      'gr.Transform.Scale(scaleX, scaleY)
      gr.ScaleTransform(scaleX, scaleY)

      Dim CurX As Single

      barWidth = bmp.Width \ Pattern.Length
      CurX = 0

      For Each c As Char In Pattern.ToCharArray
        If c = "1"c Then
          If ShowText Then
            If (CurX > 13 And CurX < 56) Or (CurX > 59 And CurX < 102) Then
              gr.FillRectangle(Brushes.Black, CurX, 0, 1, 70)
            Else
              gr.FillRectangle(Brushes.Black, CurX, 0, 1, 90)
            End If
          Else
            gr.FillRectangle(Brushes.Black, CurX, 0, 1, 100)
          End If
        End If
        CurX += 1
      Next

      If ShowText Then
        Dim LateralEsquerre As Integer = 14
        Dim GrupDades As Integer = 42
        Dim SeparadorCentral As Integer = 5
        Dim LateralDret As Integer = 10

        Using fnt As New Font("Arial", 12, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString("000000", fnt, New Point(0, 0), Nothing)

          scaleX = CSng(bmpWidth * GrupDades / 113 / ds.Width)
          scaleY = CSng(bmpHeight * 0.3 / ds.Height)

          gr.ResetTransform()
          gr.ScaleTransform(scaleX, scaleY)

          gr.DrawString(Data.Substring(0, 1), fnt, Brushes.Black, 0, CSng(bmpHeight * 0.7 / scaleY))
          gr.DrawString(Data.Substring(1, 6), fnt, Brushes.Black, CSng(bmpWidth * 14 / 113 / scaleX), CSng(bmpHeight * 0.7 / scaleY))
          gr.DrawString(Data.Substring(7, 6), fnt, Brushes.Black, CSng(bmpWidth * 61 / 113 / scaleX), CSng(bmpHeight * 0.7 / scaleY))

        End Using

      End If
      gr.Dispose()

      Return bmp
    End Function

    Private Function ImageOld() As Bitmap

      Dim bmp As New Bitmap(Width * 12, Height * 12)

      ' ens ho pasa en 1/100" i les impressores arriben fins a 1200 ppp

      Dim gr As Graphics
      gr = Graphics.FromImage(bmp)
      gr.Clear(Color.White)

      Data = EAN13_GetValidCode(Data)
      If Data.Length = 0 Then
        Using pencil As New Pen(Color.Black, 5)
          gr.DrawRectangle(pencil, 5, 5, bmp.Width - 10, bmp.Height - 10)
          gr.DrawLine(pencil, 5, 5, bmp.Width - 10, bmp.Height - 10)
          gr.DrawLine(pencil, 5, bmp.Height - 10, bmp.Width - 10, 5)
        End Using
        Return bmp
      End If

      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic
      Dim barWidth As Integer
      Dim Pattern As String = EAN13_GetMask(Data)
      Dim CurX As Integer

      barWidth = bmp.Width \ Pattern.Length
      CurX = 0

      For Each c As Char In Pattern.ToCharArray
        If c = "1"c Then
          gr.FillRectangle(Brushes.Black, CurX, 0, barWidth, bmp.Height)
        End If
        CurX += barWidth
      Next

      If ShowText Then
        Dim LateralEsquerre As Integer = 14
        Dim GrupDades As Integer = 42
        Dim SeparadorCentral As Integer = 5
        Dim LateralDret As Integer = 10

        Using fnt As New Font("Arial", 36, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString("000000", fnt, New Point(0, 0), Nothing)
          Dim dw As Integer = GrupDades * barWidth
          Dim dh As Integer = CInt(ds.Height * dw / ds.Width)

          If dh > CInt(bmp.Height) \ 3 Then
            dh = CInt(bmp.Height) \ 3
          End If

          gr.FillRectangle(Brushes.White, 0, bmp.Height - dh \ 2, bmp.Width, dh)

          Dim dd As SizeF = gr.MeasureString("0", fnt, New Point(0, 0), Nothing)

          Using bmpData As New Bitmap(CInt(dd.Width), CInt(dd.Height))
            Dim grData As Graphics
            grData = Graphics.FromImage(bmpData)
            grData.DrawString(Data.Substring(0, 1), fnt, Brushes.Black, 0, 0)
            gr.DrawImage(bmpData, 0, bmp.Height - dh, barWidth * 10, dh)
          End Using

          Using bmpData As New Bitmap(CInt(ds.Width), CInt(ds.Height))
            Dim grData As Graphics
            grData = Graphics.FromImage(bmpData)
            grData.DrawString(Data.Substring(1, 6), fnt, Brushes.Black, 0, 0)
            gr.FillRectangle(Brushes.White, LateralEsquerre * barWidth, bmp.Height - dh, dw, dh)
            gr.DrawImage(bmpData, LateralEsquerre * barWidth, bmp.Height - dh, dw, dh)
          End Using

          Using bmpData As New Bitmap(CInt(ds.Width), CInt(ds.Height))
            Dim grData As Graphics
            grData = Graphics.FromImage(bmpData)
            grData.DrawString(Data.Substring(7, 6), fnt, Brushes.Black, 0, 0)
            gr.FillRectangle(Brushes.White, (LateralEsquerre + GrupDades + SeparadorCentral) * barWidth, bmp.Height - dh, dw, dh)
            gr.DrawImage(bmpData, (LateralEsquerre + GrupDades + SeparadorCentral) * barWidth, bmp.Height - dh, dw, dh)
          End Using

        End Using

      End If
      gr.Dispose()

      Return bmp
    End Function

    Private Function EAN13_GetValidCode(ByVal Data As String) As String
      Dim ReturnValue As String
      If Data.Length >= 12 Then
        ReturnValue = Data.Substring(0, 12)
        For Each c As Char In ReturnValue.ToCharArray
          If Not Char.IsDigit(c) Then
            ReturnValue = ""
            Return ReturnValue
          End If
        Next
        ReturnValue = ReturnValue + EAN13_GetChkDigit(ReturnValue)
      Else
        ReturnValue = ""
      End If

      Return ReturnValue

    End Function

    Private Function EAN13_GetChkDigit(ByVal Ean13 As String) As String
      Dim Suma As Integer = 0
      Dim DigitControl As Integer
      Dim Multiplica As Integer = 1

      For i As Integer = 0 To 11
        Suma += CInt(Ean13.Substring(i, 1)) * Multiplica
        Multiplica = CInt(IIf(Multiplica = 1, 3, 1))
      Next

      DigitControl = 10 - (Suma Mod 10)

      If DigitControl = 10 Then
        DigitControl = 0
      End If

      Return DigitControl.ToString

    End Function

    Private Function EAN13_GetMask(ByVal Data As String) As String
      Dim Ean13 As New System.Text.StringBuilder()
      Data = (EAN13_GetValidCode(Data))

      If String.IsNullOrEmpty(Data) Then
        Return ""
      End If

      Dim Joc(,) As String = { _
         {"0001101", "0011001", "0010011", "0111101", "0100011", "0110001", "0101111", "0111011", "0110111", "0001011"}, _
         {"0100111", "0110011", "0011011", "0100001", "0011101", "0111001", "0000101", "0010001", "0001001", "0010111"}, _
         {"1110010", "1100110", "1101100", "1000010", "1011100", "1001110", "1010000", "1000100", "1001000", "1110100"}}

      Dim Mascara() As String = _
         {"111111", "112122", "112212", "112221", "121122", "122112", "122211", "121212", "121221", "122121"}

      Dim LateralEsquerre As String = "00000000000101"
      Dim SeparadorCentral As String = "01010"
      Dim LateralDret As String = "1010000000"

      Dim WorkingMask As String = Mascara(CInt(Data.Substring(0, 1)))

      Ean13.Append(LateralEsquerre)

      For i As Integer = 1 To 6
        Ean13.Append(Joc(CInt(WorkingMask.Substring(i - 1, 1)) - 1, CInt(Data.Substring(i, 1))))
      Next

      Ean13.Append(SeparadorCentral)

      For i As Integer = 7 To 12
        Ean13.Append(Joc(2, CInt(Data.Substring(i, 1))))
      Next

      Ean13.Append(LateralDret)

      Return Ean13.ToString

    End Function

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

  End Class

  Friend Class csI2of5
    Inherits csBarcodeBase

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As Bitmap
      Dim bmp As New Bitmap(Width * 12, Height * 12)
      ' ens ho pasa en 1/100" i les impressores arriben fins a 1200 ppp

      Dim gr As Graphics
      gr = Graphics.FromImage(bmp)
      gr.Clear(Color.White)

      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic
      Dim barWidth As Integer
      Dim Pattern As String = I2of5_GetMask(Data)
      Dim CurX As Integer

      barWidth = bmp.Width \ Pattern.Length
      CurX = 0

      For Each c As Char In Pattern.ToCharArray
        If c = "1"c Then
          gr.FillRectangle(Brushes.Black, CurX, 0, barWidth, bmp.Height)
        End If
        CurX += barWidth
      Next

      If ShowText Then

        If FontSize = 0 Then
          FontSize = 8
        End If

        Using fnt As New Font("Arial", FontSize, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString(Text, fnt, New Point(0, 0), Nothing)

          gr.FillRectangle(Brushes.White, 0, bmp.Height - ds.Height, bmp.Width, ds.Height)

          If ds.Width > bmp.Width Then
            Using bmpData As New Bitmap(CInt(bmp.Width), CInt(ds.Height))
              Dim grData As Graphics
              grData = Graphics.FromImage(bmpData)
              grData.DrawString(Text, fnt, Brushes.Black, 0, 0)
              gr.DrawImage(bmpData, 0, bmp.Height - ds.Height, barWidth * 10, ds.Height)
            End Using
          Else
            Dim fs As New StringFormat
            fs.Alignment = StringAlignment.Center
            gr.DrawString(Text, fnt, Brushes.Black, bmp.Width \ 2, bmp.Height - ds.Height, fs)
          End If

        End Using

      End If
      gr.Dispose()

      Return bmp
    End Function

    Private Function I2of5_GetValidCode(ByVal Data As String) As String
      Dim ReturnValue As String
      ReturnValue = Data
      For Each c As Char In ReturnValue.ToCharArray
        If Not Char.IsDigit(c) Then
          ReturnValue = ""
          Return ReturnValue
        End If
      Next

      If (ReturnValue.Length Mod 2) = 1 Then
        ReturnValue = "0" + ReturnValue
      End If

      Return ReturnValue

    End Function

    Friend Function I2of5_GetMask(ByVal Data As String) As String
      Dim nDig1 As Integer
      Dim nDig2 As Integer

      Dim cI2of5 As New System.Text.StringBuilder()
      Data = (I2of5_GetValidCode(Data))

      If String.IsNullOrEmpty(Data) Then
        Return ""
      End If

      Dim aDigit() As String = {"NNWWN", "WNNNW", "NWNNW", "WWNNN", "NNWNW", "WNWNN", "NWWNN", "NNNWW", "WNNWN", "NWNWN"}

      Dim ns As String = "0"
      Dim ws As String = "000"
      Dim nb As String = "1"
      Dim wb As String = "111"

      cI2of5.Append(nb)
      cI2of5.Append(ns)
      cI2of5.Append(nb)
      cI2of5.Append(ns)

      For i As Integer = 0 To Data.Length - 1 Step 2

        nDig1 = CInt(Data.Substring(i, 1))
        nDig2 = CInt(Data.Substring(i + 1, 1))
        For j As Integer = 0 To 4
          ' Barres
          If aDigit(nDig1).Substring(j, 1) = "N" Then
            cI2of5.Append(nb)
          Else
            cI2of5.Append(wb)
          End If
          ' Espais
          If aDigit(nDig2).Substring(j, 1) = "N" Then
            cI2of5.Append(ns)
          Else
            cI2of5.Append(ws)
          End If
        Next
      Next

      cI2of5.Append(wb)
      cI2of5.Append(ns)
      cI2of5.Append(nb)

      Return cI2of5.ToString

    End Function

  End Class

  Friend Class csCode39
    Inherits csBarcodeBase

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As Bitmap
      Dim bmp As New Bitmap(Width * 12, Height * 12)
      ' ens ho pasa en 1/100" i les impressores arriben fins a 1200 ppp

      Dim gr As Graphics
      gr = Graphics.FromImage(bmp)
      gr.Clear(Color.White)

      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic
      Dim barWidth As Integer
      Dim Pattern As String = Code39_GetMask(Data)
      Dim CurX As Integer

      barWidth = bmp.Width \ Pattern.Length
      CurX = 0

      For Each c As Char In Pattern.ToCharArray
        If c = "1"c Then
          gr.FillRectangle(Brushes.Black, CurX, 0, barWidth, bmp.Height)
        End If
        CurX += barWidth
      Next

      If ShowText Then

        If FontSize = 0 Then
          FontSize = 8
        End If

        Using fnt As New Font("Arial", FontSize, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString(Text, fnt, New Point(0, 0), Nothing)

          gr.FillRectangle(Brushes.White, 0, bmp.Height - ds.Height, bmp.Width, ds.Height)

          If ds.Width > bmp.Width Then
            Using bmpData As New Bitmap(CInt(bmp.Width), CInt(ds.Height))
              Dim grData As Graphics
              grData = Graphics.FromImage(bmpData)
              grData.DrawString(Text, fnt, Brushes.Black, 0, 0)
              gr.DrawImage(bmpData, 0, bmp.Height - ds.Height, barWidth * 10, ds.Height)
            End Using
          Else
            Dim fs As New StringFormat
            fs.Alignment = StringAlignment.Center
            gr.DrawString(Text, fnt, Brushes.Black, bmp.Width \ 2, bmp.Height - ds.Height, fs)
          End If

        End Using

      End If
      gr.Dispose()

      Return bmp


    End Function

    Private Function Code39_GetValidCode(ByVal Data As String) As String
      Dim ReturnValue As String
      Dim ValidChars As String = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ-. *$/+%"
      ReturnValue = Data
      For Each c As Char In ReturnValue.ToCharArray
        If ValidChars.IndexOf(c) < 0 Then
          ReturnValue = ""
          Return ReturnValue
        End If
      Next
      Return ReturnValue

    End Function

    Private Function Code39_GetMask(ByVal Data As String) As String
      Dim Code39 As New System.Text.StringBuilder()
      Dim ValidChars As String = "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ-. *$/+%"

      Data = Data.ToUpper
      Data = (Code39_GetValidCode(Data))

      If String.IsNullOrEmpty(Data) Then
        Return ""
      End If

      Data = String.Format("*{0}*", Data)

      Dim ns As String = "0"
      Dim ws As String = "00"
      Dim nb As String = "1"
      Dim wb As String = "11"

      Dim aBit(43) As String

      aBit(0) = wb + ns + nb + ws + nb + ns + nb + ns + wb
      aBit(1) = nb + ns + wb + ws + nb + ns + nb + ns + wb
      aBit(2) = wb + ns + wb + ws + nb + ns + nb + ns + nb
      aBit(3) = nb + ns + nb + ws + wb + ns + nb + ns + wb
      aBit(4) = wb + ns + nb + ws + wb + ns + nb + ns + nb
      aBit(5) = nb + ns + wb + ws + wb + ns + nb + ns + nb
      aBit(6) = nb + ns + nb + ws + nb + ns + wb + ns + wb
      aBit(7) = wb + ns + nb + ws + nb + ns + wb + ns + nb
      aBit(8) = nb + ns + wb + ws + nb + ns + wb + ns + nb
      aBit(9) = nb + ns + nb + ws + wb + ns + wb + ns + nb
      aBit(10) = wb + ns + nb + ns + nb + ws + nb + ns + wb
      aBit(11) = nb + ns + wb + ns + nb + ws + nb + ns + wb
      aBit(12) = wb + ns + wb + ns + nb + ws + nb + ns + nb
      aBit(13) = nb + ns + nb + ns + wb + ws + nb + ns + wb
      aBit(14) = wb + ns + nb + ns + wb + ws + nb + ns + nb
      aBit(15) = nb + ns + wb + ns + wb + ws + nb + ns + nb
      aBit(16) = nb + ns + nb + ns + nb + ws + wb + ns + wb
      aBit(17) = wb + ns + nb + ns + nb + ws + wb + ns + nb
      aBit(18) = nb + ns + wb + ns + nb + ws + wb + ns + nb
      aBit(19) = nb + ns + nb + ns + wb + ws + wb + ns + nb
      aBit(20) = wb + ns + nb + ns + nb + ns + nb + ws + wb
      aBit(21) = nb + ns + wb + ns + nb + ns + nb + ws + wb
      aBit(22) = wb + ns + wb + ns + nb + ns + nb + ws + nb
      aBit(23) = nb + ns + nb + ns + wb + ns + nb + ws + wb
      aBit(24) = wb + ns + nb + ns + wb + ns + nb + ws + nb
      aBit(25) = nb + ns + wb + ns + wb + ns + nb + ws + nb
      aBit(26) = nb + ns + nb + ns + nb + ns + wb + ws + wb
      aBit(27) = wb + ns + nb + ns + nb + ns + wb + ws + nb
      aBit(28) = nb + ns + wb + ns + nb + ns + wb + ws + nb
      aBit(29) = nb + ns + nb + ns + wb + ns + wb + ws + nb
      aBit(30) = wb + ws + nb + ns + nb + ns + nb + ns + wb
      aBit(31) = nb + ws + wb + ns + nb + ns + nb + ns + wb
      aBit(32) = wb + ws + wb + ns + nb + ns + nb + ns + nb
      aBit(33) = nb + ws + nb + ns + wb + ns + nb + ns + wb
      aBit(34) = wb + ws + nb + ns + wb + ns + nb + ns + nb
      aBit(35) = nb + ws + wb + ns + wb + ns + nb + ns + nb
      aBit(36) = nb + ws + nb + ns + nb + ns + wb + ns + wb
      aBit(37) = wb + ws + nb + ns + nb + ns + wb + ns + nb
      aBit(38) = nb + ws + wb + ns + nb + ns + wb + ns + nb
      aBit(39) = nb + ws + nb + ns + wb + ns + wb + ns + nb
      aBit(40) = nb + ws + nb + ws + nb + ws + nb + ns + nb
      aBit(41) = nb + ws + nb + ws + nb + ns + nb + ws + nb
      aBit(42) = nb + ws + nb + ns + nb + ws + nb + ws + nb
      aBit(43) = nb + ns + nb + ws + nb + ws + nb + ws + nb

      For Each c As Char In Data
        Code39.Append(aBit(ValidChars.IndexOf(c)))
        Code39.Append(ns)
      Next

      Return Code39.ToString

    End Function

  End Class

  Friend Class csDUN14
    Inherits csBarcodeBase

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As Bitmap
      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic
      Dim Pattern As String
      Dim i2of5 As New csI2of5(BarcodeSymbologies.I2of5, Data, "", 0, 0, False, 0)
      i2of5.Data = Data
      Pattern = i2of5.I2of5_GetMask(Data)

      Return OneD_DrawPattern(Pattern, Width, Height)

    End Function

    Private Function DUN14_GetChkDigit(ByVal DUN14 As String) As String
      Dim nSuma As Integer
      Dim cDigit As String

      DUN14 = DUN14.Trim
      If DUN14.Length < 13 Then
        Return ""
      End If

      ' Suma Posicions senars * 3
      nSuma = 0
      For i As Integer = 0 To 6
        nSuma += Val(DUN14(i * 2))
      Next
      nSuma *= 3
      ' Suma posicions parelles
      For i As Integer = 0 To 5
        nSuma += Val(DUN14(i * 2 + 1))
      Next
      nSuma = 10 - nSuma Mod 10
      If nSuma = 10 Then
        nSuma = 0
      End If
      cDigit = nSuma.ToString

      Return cDigit

    End Function

    Private Function DUN14_GetValidCode(ByVal Data As String) As String
      Dim ReturnValue As String
      ReturnValue = Data.Trim

      For Each c As Char In ReturnValue.ToCharArray
        If Not Char.IsDigit(c) Then
          ReturnValue = ""
          Return ReturnValue
        End If
      Next

      If ReturnValue.Length = 13 Then
        ReturnValue += DUN14_GetChkDigit(ReturnValue)
      End If

      Return ReturnValue

    End Function

  End Class

  Friend Class csCode128
    Inherits csBarcodeBase
 
    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As System.Drawing.Bitmap
      'Suposarem de partida que els cero i el 1 te la mateixa amplada
      'Repartirem el codi amb la amplada del grafic
      Dim gr As Graphics
      Dim wrkBitmap As New Bitmap(Width * 12, Height * 12)

      gr = Graphics.FromImage(wrkBitmap)
      gr.Clear(Color.White)
      Code128_DrawBarcode(gr, 0, 0, Width * 12, Height * 12)
      gr.Dispose()

      Return wrkBitmap

    End Function

    Private Sub Code128_DrawBarcode(ByVal gr As Graphics, ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer)

      Dim pattern As String
      Dim barWidth As Integer
      Dim barHeight As Integer
      Dim CurX As Single

      pattern = Code128_GetMask(Data)
      barWidth = Width \ pattern.Length
      CurX = 0

      If ShowText Then
        barHeight = CInt(Height * 0.8)
      Else
        barHeight = Height
      End If

      'Apliquem transformacions d'escala per al eix x

      gr.TranslateTransform(CSng(x), CSng(y))
      gr.ScaleTransform(CSng(Width / (barWidth * pattern.Length)), 1)

      For Each c As Char In pattern.ToCharArray
        If c = "1"c Then
          gr.FillRectangle(Brushes.Black, CurX, 0, barWidth, Height)
        End If
        CurX += barWidth
      Next

      gr.ResetTransform()

      If ShowText Then

        If FontSize = 0 Then FontSize = 60

        Using fnt As New Font("Arial", FontSize, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString(Text, fnt, New Point(0, 0), Nothing)

          gr.FillRectangle(Brushes.White, x, y + barHeight, Width, Height - barHeight)

          gr.TranslateTransform(x, y + barHeight)
          gr.ScaleTransform(CSng(Width / ds.Width), CSng((Height - barHeight) / ds.Height))

          gr.DrawString(Text, fnt, Brushes.Black, 0, 0)
          gr.ResetTransform()

        End Using

      End If

      gr.Dispose()

    End Sub

    Private Function Code128_GetMask(ByVal Data As String) As String
      Dim c128 As New mw6Code128
      Select Case Simbology
        Case BarcodeSymbologies.Code128Auto
          c128.Code128Auto(Data)
        Case BarcodeSymbologies.Code128A
          c128.Code128A(Data)
        Case BarcodeSymbologies.Code128B
          c128.Code128B(Data)
        Case BarcodeSymbologies.Code128C
          c128.Code128C(Data)
        Case BarcodeSymbologies.UCCEAN128
          c128.UCCEAN128(Data)
        Case BarcodeSymbologies.SSCC
          c128.UCCEAN128(Data)
      End Select
      Return c128.Mask
    End Function

  End Class

  Friend Class csSSCC
    Inherits csBarcodeBase
    Private bc As csCode128

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
      bc = New csCode128(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As System.Drawing.Bitmap
      Return bc.Image
    End Function


  End Class

  Friend Class csPdf417
    Inherits csBarcodeBase

    Public Sub New(ByVal Simbology As BarcodeSymbologies, ByVal Data As String, ByVal Text As String, ByVal Width As Integer, ByVal Height As Integer, ByVal ShowText As Boolean, ByVal FontSize As Integer)
      MyBase.New(Simbology, Data, Text, Width, Height, ShowText, FontSize)
    End Sub

    Public Overrides Function Image() As System.Drawing.Bitmap
      Dim gr As Graphics
      Dim barHeight As Integer
      Dim boxWidth As Integer
      Dim boxHeight As Integer
      Dim CurX As Single
      Dim CurY As Single

      If ShowText Then
        barHeight = CInt(Height * 0.5)
      Else
        barHeight = Height
      End If

      Dim pdf As New Pdf417lib
      pdf.setText(Data)
      pdf.AspectRatio = barHeight / Width

      pdf.paintCode()

      boxWidth = Width * 12 \ pdf.BitColumns
      boxHeight = boxWidth * 3

      Dim mask As ArrayList

      Dim cols As Integer = (pdf.BitColumns - 1) \ 8 + 1

      CurX = 0
      CurY = 0

      mask = pdf.getMaskString()

      Dim bmHeight As Integer
      Dim bmWidth As Integer

      bmHeight = mask.Count * boxHeight
      bmWidth = pdf.BitColumns * boxWidth

      If ShowText Then
        bmHeight = CInt(bmHeight * 1.5)
      End If

      Dim wrkBitmap As New Bitmap(bmWidth, bmHeight)
      gr = Graphics.FromImage(wrkBitmap)
      gr.Clear(Color.White)

      For Each l As String In mask
        For Each c As Char In l
          If c = "1"c Then
            gr.FillRectangle(Brushes.Black, CurX, CurY, boxWidth, boxHeight)
          End If
          CurX += boxWidth
        Next
        CurY += boxHeight
        CurX = 0
      Next
      'gr.DrawRectangle(Pens.Black, 0, 0, bmWidth - 1, bmHeight - 1)

      If ShowText Then

        If FontSize = 0 Then FontSize = 60

        Using fnt As New Font("Arial", FontSize, FontStyle.Regular)

          Dim ds As SizeF = gr.MeasureString(Text, fnt, New Point(0, 0), Nothing)

          gr.TranslateTransform(0, CurY)
          ' disminuim l'amplada amb -2 per try-error
          gr.ScaleTransform(CSng((bmWidth) / ds.Width), CSng((bmHeight - CurY) / ds.Height))

          gr.DrawString(Text, fnt, Brushes.Black, 0, 0)
          gr.ResetTransform()

        End Using

      End If

      gr.Dispose()

      Return wrkBitmap

    End Function

  End Class

#Region " Auxiliary Classes "
  Friend Class mw6Code128
    ' VB / VBA Functions for Code128(A, B,C), UCC/EAN 128
    ' Copyright 2004 by MW6 Technologies Inc. All rights reserved.
    '
    ' This code may not be modified or distributed unless you purchase
    ' the license from MW6.

    Private I As Integer
    Private StrLen As Integer
    Private Sum As Integer
    Private CurrSet As Integer
    Private CurrChar As Integer
    Private NextChar As Integer
    Private Mascara As New System.Text.StringBuilder
    Private Weight As Integer
    Private Pattern() As String = { _
          "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213", _
          "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132", _
          "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211", _
          "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313", _
          "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331", _
          "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111", _
          "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214", _
          "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111", _
          "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141", _
          "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141", _
          "114131", "311141", "411131", "211412", "211214", "211232", "2331112"}

    Public ReadOnly Property Mask() As String
      Get
        Return Mascara.ToString
      End Get
    End Property


    Public Sub Code128Auto(ByVal Src As String)
      StrLen = Len(Src)
      Sum = 104

      ' 2 indicates Set B
      CurrSet = 2

      ' start character with value 104 for Set B
      Mascara.Length = 0
      Mascara.Append(GetSymbolMask(104))

      CurrChar = Asc(Mid(Src, 1, 1))
      If (CurrChar <= 31 And CurrChar >= 0) Then
        ' switch to Set A
        ' 1 indicates Set A
        CurrSet = 1

        ' start character with value 103 for Set A
        Mascara.Length = 0
        Mascara.Append(GetSymbolMask(103))

        Sum = 103
      End If

      If Src.Length > 2 Then
        If Char.IsDigit(Src(0)) AndAlso Char.IsDigit(Src(1)) AndAlso Char.IsDigit(Src(2)) Then
          ' switch to Set C
          ' 1 indicates Set C
          CurrSet = 3

          ' start character with value 105 for Set C
          Mascara.Length = 0
          Mascara.Append(GetSymbolMask(105))

          Sum = 105
        End If
      End If

      Weight = 1
      GeneralEncode(Src)

    End Sub

    Public Sub UCCEAN128(ByVal Src As String)
      StrLen = Len(Src)
      Sum = 105

      ' 3 indicates Set C
      CurrSet = 3

      ' start character (203) + FNC1 (200)
      Mascara.Append(GetSymbolMask(105))
      Mascara.Append(GetSymbolMask(102))

      Sum += 102
      Weight = 2

      GeneralEncode(Src)

    End Sub

    Public Sub GeneralEncode(ByVal Src As String)
      Dim tmp As Integer
      Dim CurrDone As Boolean

      I = 1
      While (I <= StrLen)
        CurrChar = Asc(Mid(Src, I, 1))
        CurrDone = False
        If ((I + 1) <= StrLen) Then
          NextChar = Asc(Mid(Src, I + 1, 1))

          If (CurrChar >= Asc("0") And CurrChar <= Asc("9") And _
              NextChar >= Asc("0") And NextChar <= Asc("9")) Then
            tmp = (CurrChar - Asc("0")) * 10 + (NextChar - Asc("0"))

            ' 2 digits
            If (CurrSet <> 3) Then
              ' the previous set is not Set C
              Mascara.Append(GetSymbolMask(99))
              Sum = Sum + Weight * 99
              Weight = Weight + 1
              CurrSet = 3
            End If

            Mascara.Append(GetSymbolMask(tmp))
            Sum = Sum + Weight * tmp
            I = I + 2

            CurrDone = True
          End If
        End If

        If (Not CurrDone) Then
          If (CurrChar >= 0 And CurrChar <= 31) Then
            ' choose Set A
            If (CurrSet <> 1) Then
              ' the previous set is not Set A
              Mascara.Append(GetSymbolMask(101))
              Sum = Sum + Weight * 101
              Weight = Weight + 1
              CurrSet = 1
            End If

            If (CurrChar = 31) Then
              Mascara.Append(GetSymbolMask(95))
              Sum = Sum + Weight * 95
            Else
              Mascara.Append(GetSymbolMask(CurrChar + 64))
              Sum = Sum + Weight * (CurrChar + 64)
            End If
          Else
            ' choose Set B
            If (CurrSet <> 2) Then
              ' the previous set is not Set B
              Mascara.Append(GetSymbolMask(100))
              Sum = Sum + Weight * 100
              Weight = Weight + 1
              CurrSet = 2
            End If

            If (CurrChar = 32) Then
              Mascara.Append(GetSymbolMask(0))
            ElseIf (CurrChar = 127) Then
              Mascara.Append(GetSymbolMask(95))
              Sum = Sum + Weight * 95
            ElseIf (CurrChar < 127 And CurrChar > 32) Then
              Mascara.Append(GetSymbolMask(CurrChar - 32))
              Sum = Sum + Weight * (CurrChar - 32)
            End If
          End If

          I = I + 1
        End If

        Weight = Weight + 1
      End While

      ' add CheckDigit
      Sum = Sum Mod 103
      Mascara.Append(GetSymbolMask(Sum))

      ' add stop character (204)
      Mascara.Append(GetSymbolMask(106))

    End Sub

    Public Sub Code128A(ByVal Src As String)
      StrLen = Len(Src)
      Sum = 103

      ' start character (201) for Set A
      Mascara.Length = 0
      Mascara.Append(GetSymbolMask(103))

      Weight = 1
      For I = 1 To StrLen
        CurrChar = Asc(Mid(Src, I, 1))
        If (CurrChar = 32) Then
          Mascara.Append(GetSymbolMask(0))
        ElseIf (CurrChar = 31) Then
          Mascara.Append(GetSymbolMask(95))
          Sum = Sum + Weight * 95
        ElseIf (CurrChar <= 95 And CurrChar > 32) Then
          Mascara.Append(GetSymbolMask(CurrChar - 32))
          Sum = Sum + Weight * (CurrChar - 32)
        ElseIf (CurrChar >= 0 And CurrChar <= 31) Then
          Mascara.Append(GetSymbolMask(CurrChar + 64))
          Sum = Sum + Weight * (CurrChar + 64)
        Else
          Code128Auto(Src)
          Return
        End If
        Weight = Weight + 1
      Next I

      ' add CheckDigit
      Sum = Sum Mod 103
      Mascara.Append(GetSymbolMask(Sum))

      ' add stop character (204)
      Mascara.Append(GetSymbolMask(106))

    End Sub

    Public Sub Code128B(ByVal Src As String)
      StrLen = Len(Src)
      Sum = 104

      ' start character (202) for Set B
      Mascara.Length = 0
      Mascara.Append(GetSymbolMask(104))

      Weight = 1
      For I = 1 To StrLen
        CurrChar = Asc(Mid(Src, I, 1))
        If (CurrChar = 32) Then
          Mascara.Append(GetSymbolMask(0))
        ElseIf (CurrChar = 127) Then
          Mascara.Append(GetSymbolMask(95))
          Sum = Sum + Weight * 95
        ElseIf (CurrChar < 127 And CurrChar > 32) Then
          Mascara.Append(GetSymbolMask(CurrChar - 32))
          Sum = Sum + Weight * (CurrChar - 32)
        Else
          Code128Auto(Src)
          Return
        End If

        Weight = Weight + 1
      Next I

      ' add CheckDigit
      Sum = Sum Mod 103
      Mascara.Append(GetSymbolMask(Sum))

      ' add stop character (204)
      Mascara.Append(GetSymbolMask(106))

    End Sub

    Public Sub Code128C(ByVal Src As String)
      Dim tmp As Integer

      StrLen = Len(Src)
      Sum = 105

      ' start character (203) for Set C
      Mascara.Length = 0
      Mascara.Append(GetSymbolMask(105))

      Weight = 1
      I = 1
      While (I <= StrLen)
        CurrChar = Asc(Mid(Src, I, 1))
        If ((I + 1) <= StrLen) Then
          NextChar = Asc(Mid(Src, I + 1, 1))

          If (CurrChar >= Asc("0") And CurrChar <= Asc("9") And _
              NextChar >= Asc("0") And NextChar <= Asc("9")) Then
            '2 digits
            tmp = (CurrChar - Asc("0")) * 10 + (NextChar - Asc("0"))

            Mascara.Append(GetSymbolMask(tmp))

            Sum = Sum + Weight * tmp
            I = I + 2
          Else
            Code128Auto(Src)

            Return
          End If
        Else
          Mascara.Append(GetSymbolMask(100))
          Sum = Sum + Weight * 100
          Weight = Weight + 1

          If (CurrChar = 32) Then
            Mascara.Append(GetSymbolMask(0))
          ElseIf (CurrChar = 127) Then
            Mascara.Append(GetSymbolMask(95))
            Sum = Sum + Weight * 95
          ElseIf (CurrChar < 127 And CurrChar > 32) Then
            Mascara.Append(GetSymbolMask(CurrChar - 32))
            Sum = Sum + Weight * (CurrChar - 32)

          Else
            Code128Auto(Src)
            Return
          End If
          I = I + 1
        End If

        Weight = Weight + 1
      End While

      ' add CheckDigit
      Sum = Sum Mod 103
      Mascara.Append(GetSymbolMask(Sum))

      ' add stop character (204)
      Mascara.Append(GetSymbolMask(106))

    End Sub

    Private Function GetSymbolMask(ByVal SymbolValue As Integer) As String
      Dim SymbolMask As New System.Text.StringBuilder
      ' Pattern = BsBsBsBsB
      For i As Integer = 0 To Pattern(SymbolValue).Length - 1
        SymbolMask.Append(CChar(((i + 1) Mod 2).ToString), CInt(Pattern(SymbolValue).Substring(i, 1)))
      Next
      Return SymbolMask.ToString
    End Function
  End Class

  '
  ' Copyright 2003-2005 by Paulo Soares.
  '
  ' The contents of this file are subject to the Mozilla Public License Version 1.1
  ' (the "License") you may not use this file except in compliance with the License.
  ' You may obtain a copy of the License at http:'www.mozilla.org/MPL/
  '
  ' Software distributed under the License is distributed on an "AS IS" basis,
  ' WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
  ' for the specific language governing rights and limitations under the License.
  '
  ' The Original Code is 'pdf417lib, a library to generate the bidimensional barcode PDF417'.
  '
  ' The Initial Developer of the Original Code is Paulo Soares. Portions created by
  ' the Initial Developer are Copyright (C) 2003 by Paulo Soares.
  ' All Rights Reserved.
  '
  ' Contributor(s): all the names of the contributors are added in the source code
  ' where applicable.
  '
  ' Alternatively, the contents of this file may be used under the terms of the
  ' LGPL license (the "GNU LIBRARY GENERAL PUBLIC LICENSE"), in which case the
  ' provisions of LGPL are applicable instead of those above.  If you wish to
  ' allow use of your version of this file only under the terms of the LGPL
  ' License and not to allow others to use your version of this file under
  ' the MPL, indicate your decision by deleting the provisions above and
  ' replace them with the notice and other provisions required by the LGPL.
  ' If you do not delete the provisions above, a recipient may use your version
  ' of this file under either the MPL or the GNU LIBRARY GENERAL PUBLIC LICENSE.
  '
  ' This library is free software you can redistribute it and/or modify it
  ' under the terms of the MPL as stated above or under the terms of the GNU
  ' Library General Public License as published by the Free Software Foundation
  ' either version 2 of the License, or any later version.
  '
  ' This library is distributed in the hope that it will be useful, but WITHOUT
  ' ANY WARRANTY without even the implied warranty of MERCHANTABILITY or FITNESS
  ' FOR A PARTICULAR PURPOSE. See the GNU Library general Public License for more
  ' details.
  '
  ' If you didn't download this code from the following link, you should check if
  ' you aren't using an obsolete version:
  ' http:'sourceforge.net/projects/pdf417lib
  ' This code is also used in iText (http:'www.lowagie.com/iText)
  '

#End Region

  Friend Class Pdf417lib

#Region "Variables Públiques"
    ''' <summary>
    ''' Auto-size is made based on <CODE>aspectRatio</CODE> and <CODE>yHeight</CODE>.
    ''' </summary>
    Public Const PDF417_USE_ASPECT_RATIO As Integer = 0
    ''' <summary>
    ''' The size of the barcode will be at least <CODE>codeColumns*codeRows</CODE>.
    ''' </summary>
    Public Const PDF417_FIXED_RECTANGLE As Integer = 1
    ''' <summary>
    ''' The size will be at least <CODE>codeColumns</CODE> with a variable number of <CODE>codeRows</CODE>.
    ''' </summary>
    Public Const PDF417_FIXED_COLUMNS As Integer = 2
    ''' <summary>
    ''' The size will be at least <CODE>codeRows</CODE> with a variable number of <CODE>codeColumns</CODE>.
    ''' </summary>
    Public Const PDF417_FIXED_ROWS As Integer = 4
    ''' <summary>
    ''' The error level correction is set automatically according to ISO 15438 recomendations.
    ''' </summary>
    Public Const PDF417_AUTO_ERROR_LEVEL As Integer = 0
    ''' <summary>
    ''' The error level correction is set by the user. It can be 0 to 8.
    ''' </summary>
    Public Const PDF417_USE_ERROR_LEVEL As Integer = 16
    ''' <summary>
    ''' interpretation is done and the content of <CODE>codewords</CODE> is used directly.
    ''' </summary>
    Public Const PDF417_USE_RAW_CODEWORDS As Integer = 64
    ''' <summary>
    ''' Inverts the output bits of the raw bitmap that is normally bit one for black. 
    ''' It has only effect for the raw bitmap.
    ''' </summary>
    Public Const PDF417_INVERT_BITMAP As Integer = 128

    Protected Friend bitPtr As Integer
    Protected Friend cwPtr As Integer
    Protected Friend segmentList As cSegmentList
#End Region

    Private Sub InitBlock()
      ReDim mCodewords(MAX_DATA_CODEWORDS + 2)
    End Sub

    Protected Friend ReadOnly Property MaxSquare() As Integer
      Get
        If (CodeColumns > 21) Then
          CodeColumns = 29
          CodeRows = 32
        Else
          CodeColumns = 16
          CodeRows = 58
        End If

        Return MAX_DATA_CODEWORDS + 2
      End Get

    End Property

    ''' <summary>Gets the raw image bits of the barcode. The image will have to
    ''' be scaled in the Y direction by <CODE>yHeight</CODE>.
    ''' </summary>
    ''' <returns> The raw barcode image
    ''' </returns>
    Public Overridable Property OutBits() As SByte()
      Get
        Return Me.mOutBits
      End Get
      Set(ByVal value As SByte())
        Me.mOutBits = value
      End Set
    End Property

    ''' <summary>Gets the number of X pixels of <CODE>outBits</CODE>.</summary>
    ''' <returns> the number of X pixels of <CODE>outBits</CODE>
    ''' </returns>
    Public Overridable Property BitColumns() As Integer
      Get
        Return mBitColumns
      End Get
      Set(ByVal value As Integer)
        mBitColumns = value
      End Set
    End Property

    ''' <summary>Gets the number of Y pixels of <CODE>outBits</CODE>.
    ''' It is also the number of rows in the barcode.
    ''' </summary>
    Public Overridable Property CodeRows() As Integer
      Get
        Return mCodeRows
      End Get
      Set(ByVal value As Integer)
        mCodeRows = value
      End Set
    End Property

    ''' <summary>Gets the number of barcode data columns.</summary>
    ''' <returns>The number of barcode data columns
    ''' </returns>
    Public Overridable Property CodeColumns() As Integer
      Get
        Return mCodeColumns
      End Get
      Set(ByVal value As Integer)
        mCodeColumns = value
      End Set
    End Property

    ''' <summary>Gets the codeword array. This array is always 928 elements long.
    ''' It can be writen to if the option <CODE>PDF417_USE_RAW_CODEWORDS</CODE>
    ''' is set.
    ''' </summary>
    ''' <returns> the codeword array
    ''' </returns>
    Public Overridable ReadOnly Property Codewords() As Integer()
      Get
        Return mCodewords
      End Get
    End Property

    ''' <summary>
    ''' Gets the length of the codewords.
    ''' Sets the length of the codewords.
    ''' </summary>
    Public Overridable Property LenCodewords() As Integer
      Get
        Return mLenCodewords
      End Get
      Set(ByVal value As Integer)
        mLenCodewords = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the error level correction used for the barcode. It may different
    ''' from the previously set value.
    ''' Sets the error level correction for the barcode.
    ''' </summary>
    Public Overridable Property ErrorLevel() As Integer
      Get
        Return mErrorLevel
      End Get
      Set(ByVal value As Integer)
        mErrorLevel = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the options to generate the barcode.
    ''' Sets the options to generate the barcode. This can be all the <CODE>PDF417_*</CODE> constants.
    ''' </summary>
    Public Overridable Property Options() As Integer
      Get
        Return mOptions
      End Get

      Set(ByVal value As Integer)
        mOptions = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the barcode aspect ratio.
    ''' Sets the barcode aspect ratio. A ratio or 0.5 will make the
    ''' barcode width twice as large as the height.
    ''' </summary>
    Public Overridable Property AspectRatio() As Double
      Get
        Return mAspectRatio
      End Get
      Set(ByVal value As Double)
        mAspectRatio = value
      End Set
    End Property

    ''' <summary>
    ''' Gets the Y pixel height relative to X.
    ''' Sets the Y pixel height relative to X. It is usually 3.
    ''' </summary>
    Public Overridable Property YHeight() As Double
      Get
        Return mYHeight
      End Get
      Set(ByVal value As Double)
        mYHeight = value
      End Set
    End Property



    '''<summary>Creates a new <CODE>BarcodePDF417</CODE> with the default settings.</summary>
    Public Sub New()
      InitBlock()
      setDefaultParameters()
    End Sub

    Protected Friend Overridable Function checkSegmentType(ByVal segment As Segment, ByVal type As Char) As Boolean
      If segment Is Nothing Then Return False

      Return segment.type = type
    End Function

    Protected Friend Overridable Function getSegmentLength(ByVal segment As Segment) As Integer
      If segment Is Nothing Then Return 0

      Return segment.iEnd - segment.start
    End Function

    ''' <summary>Set the default settings that correspond to <CODE>PDF417_USE_ASPECT_RATIO</CODE>
    ''' and <CODE>PDF417_AUTO_ERROR_LEVEL</CODE>.
    ''' </summary>
    Public Overridable Sub setDefaultParameters()
      Options = 0
      mOutBits = Nothing
      mText = Nothing
      YHeight = 3
      AspectRatio = 0.5F
    End Sub

    Protected Friend Overridable Sub outCodeword17(ByVal codeword As Integer)
      Dim bytePtr As Integer = CInt(Math.Truncate(bitPtr / 8))
      Dim bit As Integer = bitPtr - bytePtr * 8
      Dim ibyte As Integer = (mOutBits(bytePtr) Or (codeword >> (9 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256)) 'agrupem els bits en bytes
      bytePtr += 1
      ibyte = (mOutBits(bytePtr) Or (codeword >> (1 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256))
      bytePtr += 1
      codeword <<= 8
      ibyte = (mOutBits(bytePtr) Or (codeword >> (1 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256))
      bitPtr += 17
    End Sub

    Protected Friend Overridable Sub outCodeword18(ByVal codeword As Integer)
      Dim bytePtr As Integer = CInt(Math.Truncate(bitPtr / 8))
      Dim bit As Integer = bitPtr - bytePtr * 8
      Dim ibyte As Integer = (mOutBits(bytePtr) Or (codeword >> (10 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256))
      bytePtr += 1
      ibyte = (mOutBits(bytePtr) Or (codeword >> (2 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256))
      bytePtr += 1
      codeword <<= 8
      ibyte = (mOutBits(bytePtr) Or (codeword >> (2 + bit))) Mod 256
      If (ibyte < 128) Then
        mOutBits(bytePtr) = CSByte(ibyte Mod 128)
      Else
        mOutBits(bytePtr) = CSByte((ibyte Mod 256) - 256)
      End If
      'mOutBits(bytePtr) = IIf(ibyte < 127, CSByte(ibyte Mod 128), CSByte((ibyte Mod 256) - 256))
      If (bit = 7) Then
        bytePtr += 1
        mOutBits(bytePtr) = CSByte(OutBits(bytePtr) - 128) 'unchecked((sbyte)0x80)
      End If
      bitPtr += 18
    End Sub

    Protected Friend Overridable Sub outCodeword(ByVal codeword As Integer)
      outCodeword17(codeword)
    End Sub

    Protected Friend Overridable Sub outStopPattern()
      outCodeword18(STOP_PATTERN)
    End Sub

    Protected Friend Overridable Sub outStartPattern()
      outCodeword17(START_PATTERN)
    End Sub

    Protected Friend Overridable Sub outPaintCode()
      Dim codePtr As Integer = 0
      BitColumns = START_CODE_SIZE * (CodeColumns + 3) + STOP_SIZE
      Dim lenBits As Integer = (((BitColumns - 1) \ 8) + 1) * CodeRows
      ReDim mOutBits(lenBits - 1)
      For row As Integer = 0 To CodeRows - 1
        bitPtr = (((BitColumns - 1) \ 8) + 1) * 8 * row
        Dim rowMod As Integer = row Mod 3
        Dim cluster() As Integer = CLUSTERS(rowMod)
        outStartPattern()
        Dim edge As Integer = 0
        Select Case rowMod
          Case 0
            edge = 30 * (row \ 3) + ((CodeRows - 1) \ 3)
          Case 1
            edge = 30 * (row \ 3) + ErrorLevel * 3 + ((CodeRows - 1) Mod 3)
          Case Else
            edge = 30 * (row \ 3) + CodeColumns - 1
        End Select

        outCodeword(cluster(edge))

        For column As Integer = 0 To CodeColumns - 1
          outCodeword(cluster(mCodewords(codePtr)))
          codePtr += 1
        Next

        Select Case rowMod
          Case 0
            edge = 30 * (row \ 3) + CodeColumns - 1
          Case 1
            edge = 30 * (row \ 3) + ((CodeRows - 1) \ 3)
          Case Else
            edge = 30 * (row \ 3) + ErrorLevel * 3 + ((CodeRows - 1) Mod 3)
        End Select

        outCodeword(cluster(edge))
        outStopPattern()
      Next

      If ((Options And PDF417_INVERT_BITMAP) <> 0) Then
        For k As Integer = 0 To OutBits.Length - 1
          mOutBits(k) = CSByte((((255 - mOutBits(k)) Mod 128) + CInt(IIf(mOutBits(k) < 0, 128, -128))) Mod 128)
        Next
      End If
    End Sub

    Protected Friend Overridable Sub calculateErrorCorrection(ByVal dest As Integer)
      If (ErrorLevel < 0 Or ErrorLevel > 8) Then ErrorLevel = 0

      Dim A As Integer() = ERROR_LEVEL(ErrorLevel)
      Dim Alength As Integer = 2 << ErrorLevel
      For k As Integer = 0 To Alength - 1
        mCodewords(dest + k) = 0
      Next

      Dim lastE As Integer = Alength - 1
      For k As Integer = 0 To LenCodewords - 1
        Dim t1 As Integer = mCodewords(k) + mCodewords(dest)
        For e As Integer = 0 To lastE
          Dim t2 As Integer = (t1 * A(lastE - e)) Mod MODUL
          Dim t3 As Integer = MODUL - t2
          If e = lastE Then
            mCodewords(dest + e) = (0 + t3) Mod MODUL
          Else
            mCodewords(dest + e) = (mCodewords(dest + e + 1) + t3) Mod MODUL
          End If
        Next
      Next

      For k As Integer = 0 To Alength - 1
        mCodewords(dest + k) = (MODUL - mCodewords(dest + k)) Mod MODUL
      Next
    End Sub

    Protected Friend Overridable Function getTextTypeAndValue(ByVal maxLength As Integer, ByVal idx As Integer) As Integer
      If (idx >= maxLength) Then Return 0
      Dim c As Char = Chr(mText(idx))
      If (c >= "A" And c <= "Z") Then Return (ALPHA + mText(idx) - &H41) ' &H41 = "A"
      If (c >= "a" And c <= "z") Then Return (LOWER + mText(idx) - &H61) ' &H61 = "a"
      If (c = " ") Then Return (ALPHA + LOWER + MIXED + SPACE)

      Dim ms As Integer = MIXED_SET.IndexOf(c)
      Dim ps As Integer = PUNCTUATION_SET.IndexOf(c)
      If (ms < 0 And ps < 0) Then Return (ISBYTE + mText(idx))
      If (ms = ps) Then Return (MIXED + PUNCTUATION + ms)
      If (ms >= 0) Then Return (MIXED + ms)

      Return (PUNCTUATION + ps)
    End Function

    Protected Friend Overridable Sub textCompaction(ByVal start As Integer, ByVal length As Integer)

      Dim dest(ABSOLUTE_MAX_TEXT_SIZE * 2) As Integer
      Dim mode As Integer = ALPHA
      Dim ptr As Integer = 0
      Dim fullBytes As Integer = 0
      Dim v As Integer = 0
      Dim size As Integer = 0

      length += start
      For k As Integer = start To length - 1

        v = getTextTypeAndValue(length, k)
        If ((v And mode) <> 0) Then
          dest(ptr) = System.Math.Abs(v - mode) Mod 128
          ptr += 1
          Continue For
        End If

        If ((v And ISBYTE) <> 0) Then
          If ((ptr And 1) <> 0) Then
            If (mode And PUNCTUATION) <> 0 Then
              dest(ptr) = PAL
              ptr += 1
            Else
              dest(ptr) = PS
              ptr += 1
            End If

            If (mode And PUNCTUATION) <> 0 Then mode = ALPHA
          End If
          dest(ptr) = BYTESHIFT
          ptr += 1
          dest(ptr) = v Mod 128
          ptr += 1
          fullBytes += 2
          Continue For
        End If

        Select Case mode
          Case ALPHA
            If ((v And LOWER) <> 0) Then
              dest(ptr) = LL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = LOWER
            ElseIf ((v And MIXED) <> 0) Then
              dest(ptr) = ML
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = MIXED
            ElseIf ((getTextTypeAndValue(length, k + 1) And getTextTypeAndValue(length, k + 2) And PUNCTUATION) <> 0) Then
              dest(ptr) = ML
              ptr += 1
              dest(ptr) = PL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = PUNCTUATION
            Else
              dest(ptr) = PS
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
            End If

          Case LOWER
            If ((v And ALPHA) <> 0) Then
              If ((getTextTypeAndValue(length, k + 1) And getTextTypeAndValue(length, k + 2) And ALPHA) <> 0) Then
                dest(ptr) = ML
                ptr += 1
                dest(ptr) = AL
                ptr += 1
                mode = ALPHA
              Else
                dest(ptr) = AS_
                ptr += 1
              End If
              dest(ptr) = v Mod ALPHA
              ptr += 1
            ElseIf ((v And MIXED) <> 0) Then
              dest(ptr) = ML
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = MIXED
            ElseIf ((getTextTypeAndValue(length, k + 1) And getTextTypeAndValue(length, k + 2) And PUNCTUATION) <> 0) Then
              dest(ptr) = ML
              ptr += 1
              dest(ptr) = PL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = PUNCTUATION
            Else
              dest(ptr) = PS
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
            End If

          Case MIXED
            If ((v And LOWER) <> 0) Then
              dest(ptr) = LL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = LOWER
            ElseIf ((v And ALPHA) <> 0) Then
              dest(ptr) = AL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = ALPHA
            ElseIf ((getTextTypeAndValue(length, k + 1) And getTextTypeAndValue(length, k + 2) And PUNCTUATION) <> 0) Then
              dest(ptr) = PL
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
              mode = PUNCTUATION
            Else
              dest(ptr) = PS
              ptr += 1
              dest(ptr) = v Mod 128
              ptr += 1
            End If

          Case PUNCTUATION
            dest(ptr) = PAL
            ptr += 1
            mode = ALPHA

            k -= 1
        End Select
      Next

      If ((ptr And 1) <> 0) Then
        dest(ptr) = PS
        ptr += 1
      End If

      size = (ptr + fullBytes) \ 2
      If (size + cwPtr > MAX_DATA_CODEWORDS) Then
        Throw New System.IndexOutOfRangeException("The text is too big.")
      End If

      length = ptr
      ptr = 0
      While (ptr < length)
        v = dest(ptr)
        ptr += 1
        If (v >= 30) Then
          mCodewords(cwPtr) = v
          cwPtr += 1
          mCodewords(cwPtr) = dest(ptr)
          ptr += 1
          cwPtr += 1
        Else
          mCodewords(cwPtr) = v * 30 + dest(ptr)
          ptr += 1
          cwPtr += 1
        End If
      End While
    End Sub

    Protected Friend Overridable Sub basicNumberCompaction(ByVal start As Integer, ByVal length As Integer)
      Dim ret As Integer = CInt(cwPtr)
      Dim retLast As Integer = (length \ 3)
      cwPtr += retLast + 1

      For k As Integer = 0 To retLast - 1
        mCodewords(ret + k) = 0
      Next
      mCodewords(ret + retLast) = 1
      length += start
      For ni As Integer = start To length - 1
        ' multiply by 10
        For k As Integer = retLast To 0 Step -1
          mCodewords(ret + k) *= 10
        Next
        ' add the digit
        mCodewords(ret + retLast) += mText(ni) - Asc("0")
        ' propagate carry
        For k As Integer = retLast To 0 Step -1
          mCodewords(ret + k - 1) += mCodewords(ret + k) \ 900
          mCodewords(ret + k) = mCodewords(ret + k) Mod 900
        Next
      Next
    End Sub

    Protected Friend Overridable Sub numberCompaction(ByVal start As Integer, ByVal length As Integer)
      Dim full As Integer = (length \ 44) * 15
      Dim size As Integer = length Mod 44

      If (size = 0) Then
        size = full
      Else
        size = full + (size \ 3) + 1
      End If

      If (size + cwPtr > MAX_DATA_CODEWORDS) Then
        Throw New System.IndexOutOfRangeException("The text is too big.")
      End If

      length += start
      For k As Integer = start To length - 1 Step 44
        size = CInt(IIf(length - k < 44, length - k, 44))
        basicNumberCompaction(k, size)
      Next
    End Sub

    Protected Friend Overridable Sub byteCompaction6(ByVal start As Integer)
      Dim length As Integer = 6
      Dim ret As Integer = CInt(cwPtr)
      Dim retLast As Integer = 4

      cwPtr += retLast + 1
      For k As Integer = 0 To retLast - 1
        mCodewords(ret + k) = 0
      Next

      length += start
      For ni As Integer = start To length
        ' multiply by 256
        For k As Integer = retLast To 0 Step -1
          mCodewords(ret + k) *= 256
        Next
        ' add the digit
        mCodewords(ret + retLast) += mText(ni)
        ' propagate carry
        For k As Integer = retLast To 0 Step -1
          mCodewords(ret + k - 1) += mCodewords(ret + k) \ 900
          mCodewords(ret + k) = mCodewords(ret + k) Mod 900
        Next
      Next
    End Sub

    Private Sub byteCompaction(ByVal start As Integer, ByVal ilength As Integer)
      Dim length As Integer = ilength
      Dim size As Integer = (length \ 6) * 5 + (length Mod 6)
      If (size + cwPtr > MAX_DATA_CODEWORDS) Then
        Throw New System.IndexOutOfRangeException("The text is too big.")
      End If

      length += start
      For k As Integer = start To length Step 6
        If length - k < 44 Then
          size = length - k
        Else : size = 6
        End If

        If (size < 6) Then
          For j As Integer = 0 To size - 1
            mCodewords(cwPtr) = mText(k + j)
            cwPtr += 1
          Next
        Else
          byteCompaction6(k)
        End If
      Next
    End Sub

    Private Sub breakString()
      Dim textLength As Integer = mText.Length
      Dim lastP As Integer = 0
      Dim startN As Integer = 0
      Dim nd As Integer = 0
      Dim c As Char = "0"c
      Dim ptrS As Integer = 0
      Dim lastTxt, txt As Boolean
      Dim k As Integer = 0

      Dim v As Segment
      Dim vp As Segment
      Dim vn As Segment

      For k = 0 To textLength - 1
        c = Chr(mText(k))
        If (c >= "0" And c <= "9") Then
          If nd = 0 Then startN = k
          nd = nd + 1
          Continue For
        End If

        If (nd >= 13) Then
          If (lastP <> startN) Then
            c = Chr(mText(lastP))
            ptrS = lastP
            lastTxt = (Chr(mText(lastP)) >= " " And mText(lastP) < 127) Or c = Chr(13) Or c = Chr(10) Or c = Chr(9)
            For j As Integer = lastP To startN - 1
              c = Chr(mText(j))
              txt = (Chr(mText(j)) >= " " And mText(j) < 127) Or c = Chr(13) Or c = Chr(10) Or c = Chr(9)
              If (txt <> lastTxt) Then
                segmentList.add(CChar(IIf(lastTxt, "T"c, "B"c)), lastP, j)
                lastP = j
                lastTxt = txt
              End If
            Next
            segmentList.add(CChar(IIf(lastTxt, "T"c, "B"c)), lastP, startN)
          End If
          segmentList.add("N"c, startN, k)
          lastP = k
        End If
        nd = 0
      Next

      If (nd < 13) Then startN = textLength

      If (lastP <> startN) Then
        c = Chr(mText(lastP))
        ptrS = lastP
        lastTxt = (Chr(mText(lastP)) >= " " And mText(lastP) < 127) Or c = Chr(13) Or c = Chr(10) Or c = Chr(9)
        For j As Integer = lastP To startN - 1
          c = Chr(mText(j))
          txt = (Chr(mText(j)) >= " " And mText(j) < 127) Or c = Chr(13) Or c = Chr(10) Or c = Chr(9)
          If (txt <> lastTxt) Then
            segmentList.add(CChar(IIf(lastTxt, "T"c, "B"c)), lastP, j)
            lastP = j
            lastTxt = txt
          End If
        Next
        segmentList.add(CChar(IIf(lastTxt, "T"c, "B"c)), lastP, startN)
      End If

      If (nd >= 13) Then segmentList.add("N"c, startN, textLength)

      'optimize. merge short binary
      For k = 0 To segmentList.size() - 1
        v = segmentList.get_Renamed(k)
        vp = segmentList.get_Renamed(k - 1)
        vn = segmentList.get_Renamed(k + 1)
        If (checkSegmentType(v, "B"c) And getSegmentLength(v) = 1) Then
          If (checkSegmentType(vp, "T"c) And checkSegmentType(vn, "T"c) And getSegmentLength(vp) + getSegmentLength(vn) >= 3) Then
            vp.iEnd = vn.iEnd
            segmentList.remove(k)
            segmentList.remove(k)
            k = -1
            Continue For
          End If
        End If
      Next

      'merge text sections
      For k = 0 To segmentList.size() - 1
        v = segmentList.get_Renamed(k)
        vp = segmentList.get_Renamed(k - 1)
        vn = segmentList.get_Renamed(k + 1)
        If (checkSegmentType(v, "T"c) And getSegmentLength(v) >= 5) Then
          Dim redo As Boolean = False
          If ((checkSegmentType(vp, "B"c) And getSegmentLength(vp) = 1) Or checkSegmentType(vp, "T"c)) Then
            redo = True
            v.start = vp.start
            segmentList.remove(k - 1)
            k -= 1
          End If

          If ((checkSegmentType(vn, "B"c) And getSegmentLength(vn) = 1) Or checkSegmentType(vn, "T"c)) Then
            redo = True
            v.iEnd = vn.iEnd
            segmentList.remove(k + 1)
          End If

          If (redo) Then
            k = -1
            Continue For
          End If
        End If
      Next

      'merge binary sections
      For k = 0 To segmentList.size() - 1
        v = segmentList.get_Renamed(k)
        vp = segmentList.get_Renamed(k - 1)
        vn = segmentList.get_Renamed(k + 1)
        If (checkSegmentType(v, "B"c)) Then
          Dim redo As Boolean = False
          If ((checkSegmentType(vp, "T"c) And getSegmentLength(vp) < 5) Or checkSegmentType(vp, "B"c)) Then
            redo = True
            v.start = vp.start
            segmentList.remove(k - 1)
            k -= 1
          End If
          If ((checkSegmentType(vn, "T"c) And getSegmentLength(vn) < 5) Or checkSegmentType(vn, "B"c)) Then
            redo = True
            v.iEnd = vn.iEnd
            segmentList.remove(k + 1)
          End If
          If (redo) Then
            k = -1
            Continue For
          End If
        End If
      Next

      ' check if all numbers
      v = segmentList.get_Renamed(0)
      If (segmentList.size() = 1 And v.type = "T"c And getSegmentLength(v) >= 8) Then
        For k = 0 To v.iEnd - 1
          c = Chr(mText(k))
          If (c < "0"c Or c > "9"c) Then Return
        Next
        If (k = v.iEnd) Then v.type = "N"c
      End If
    End Sub

    Protected Friend Overridable Sub assemble()
      If (segmentList.size() = 0) Then Return
      cwPtr = 1
      For k As Integer = 0 To segmentList.size() - 1
        Dim v As Segment = segmentList.get_Renamed(k)
        Select Case v.type
          Case "T"c
            If (k <> 0) Then
              mCodewords(cwPtr) = TEXT_MODE
              cwPtr += 1
            End If
            textCompaction(v.start, getSegmentLength(v))
          Case "N"c
            mCodewords(cwPtr) = NUMERIC_MODE
            cwPtr += 1
            numberCompaction(v.start, getSegmentLength(v))
          Case "B"c
            mCodewords(cwPtr) = CSByte(CInt(IIf((getSegmentLength(v) Mod 6) <> 0, BYTE_MODE, BYTE_MODE_6)) Mod 128)
            cwPtr += 1
            byteCompaction(v.start, getSegmentLength(v))
        End Select
      Next
    End Sub

    Protected Friend Function maxPossibleErrorLevel(ByVal remain As Integer) As Integer
      Dim level As Integer = 8
      Dim size As Integer = 512
      While (level > 0)
        If (remain >= size) Then Return level
        level -= 1
        size >>= 1
      End While

      Return 0
    End Function

    Protected Friend Overridable Sub dumpList()
      If (segmentList.size() = 0) Then Return

      For k As Integer = 0 To segmentList.size() - 1
        Dim v As Segment = segmentList.get_Renamed(k)
        Dim Len As Integer = getSegmentLength(v)
        Dim c(Len) As Char
        For j As Integer = 0 To Len - 1
          c(j) = CChar(mText(v.start + j) & &HFF)
          If (c(j) = Chr(13)) Then c(j) = Chr(10)
        Next
        System.Console.Out.WriteLine("" + v.type + New System.String(c))
      Next
    End Sub

    Protected Friend Function Filter(ByVal sValue As String) As SByte()
      Dim sArray() As SByte
      ReDim sArray(sValue.Length - 1)

      Try
        Dim cArray() As Char = sValue.ToCharArray()
        For k As Integer = 0 To cArray.Length - 1
          sArray(k) = Convert.ToSByte(cArray(k))
        Next

        Return sArray

      Catch
        Return Nothing
      End Try
    End Function

    ''' <summary>Paints the barcode. If no exception was thrown a valid barcode is available. </summary>
    Public Overridable Sub paintCode()
      Dim maxErr, lenErr, tot, pad As Integer

      If Options = PDF417_USE_RAW_CODEWORDS Then
        If (LenCodewords > MAX_DATA_CODEWORDS Or LenCodewords < 1 Or LenCodewords <> mCodewords(0)) Then
          Throw New System.ArgumentException("Invalid codeword size.")
        End If
      Else
        If (mText Is Nothing) Then
          Throw New System.NullReferenceException("Text cannot be nothing.")
        End If
        If (mText.Length > ABSOLUTE_MAX_TEXT_SIZE) Then
          Throw New System.IndexOutOfRangeException("The text is too big.")
        End If

        segmentList = New cSegmentList(Me)
        breakString()
        assemble()
        segmentList = Nothing
        mCodewords(0) = cwPtr
        LenCodewords = cwPtr
      End If

      maxErr = maxPossibleErrorLevel(MAX_DATA_CODEWORDS + 2 - LenCodewords)
      If ((Options And PDF417_USE_ERROR_LEVEL) = 0) Then
        If (LenCodewords < 41) Then
          ErrorLevel = 2
        ElseIf (LenCodewords < 161) Then
          ErrorLevel = 3
        ElseIf (LenCodewords < 321) Then
          ErrorLevel = 4
        Else : ErrorLevel = 5
        End If
      End If

      If (ErrorLevel < 0) Then
        ErrorLevel = 0
      ElseIf (ErrorLevel > maxErr) Then
        ErrorLevel = maxErr
      End If

      If (CodeColumns < 1) Then
        CodeColumns = 1
      ElseIf (CodeColumns > 30) Then
        CodeColumns = 30
      End If

      If (CodeRows < 3) Then
        CodeRows = 3
      ElseIf (CodeRows > 90) Then
        CodeRows = 90
      End If

      lenErr = 2 << ErrorLevel
      Dim fixedColumn As Boolean = ((Options And PDF417_FIXED_ROWS) = 0)
      Dim skipRowColAdjust As Boolean = False
      tot = LenCodewords + lenErr

      If ((Options And PDF417_FIXED_RECTANGLE) <> 0) Then
        tot = CodeColumns * CodeRows
        If (tot > MAX_DATA_CODEWORDS + 2) Then tot = MaxSquare
        If (tot < LenCodewords + lenErr) Then
          tot = LenCodewords + lenErr
        Else
          skipRowColAdjust = True
        End If
      ElseIf (Options <> PDF417_FIXED_COLUMNS) And (Options <> PDF417_FIXED_ROWS) Then
        Dim c As Double
        Dim b As Double

        fixedColumn = True
        If (AspectRatio < 0.001) Then
          AspectRatio = 0.001F
        ElseIf (AspectRatio > 1000) Then
          AspectRatio = 1000
        End If

        b = 73 * AspectRatio - 4
        c = (-b + System.Math.Sqrt(b * b + 4 * 17 * AspectRatio * (LenCodewords + lenErr) * YHeight)) / (2 * 17 * AspectRatio)
        'UPGRADE_WARNING: Narrowing conversions may produce unexpected results in C#. 'ms-help:'MS.VSCC.2003/commoner/redir/redirect.htm?keyword="jlca1042"'
        CodeColumns = CInt(c)
        If (CodeColumns < 1) Then
          CodeColumns = 1
        ElseIf (CodeColumns > 30) Then
          CodeColumns = 30
        End If
      End If

      If Not skipRowColAdjust Then
        If (fixedColumn) Then
          CodeRows = ((tot - 1) \ CodeColumns) + 1

          If (CodeRows < 3) Then
            CodeRows = 3
          ElseIf (CodeRows > 90) Then
            CodeRows = 90
            CodeColumns = (tot - 1) \ 90 + 1
          End If
        Else
          CodeColumns = ((tot - 1) \ CodeRows) + 1
          If (CodeColumns > 30) Then
            CodeColumns = 30
            CodeRows = ((tot - 1) \ 30) + 1
          End If
        End If
        tot = CodeRows * CodeColumns
      End If

      If (tot > MAX_DATA_CODEWORDS + 2) Then tot = MaxSquare

      ErrorLevel = maxPossibleErrorLevel(tot - LenCodewords)
      lenErr = 2 << ErrorLevel
      pad = tot - lenErr - LenCodewords
      cwPtr = LenCodewords

      While (pad <> 0)
        mCodewords(cwPtr) = TEXT_MODE
        cwPtr += 1
        pad -= 1
      End While
      mCodewords(0) = cwPtr
      LenCodewords = cwPtr
      calculateErrorCorrection(LenCodewords)
      LenCodewords = tot
      outPaintCode()
    End Sub

    ''' <summary>Gets the bytes that form the barcode. This bytes should
    ''' be interpreted in the codepage Cp437.
    ''' </summary>
    ''' <returns> the bytes that form the barcode
    ''' </returns>
    Public Overridable Function getText() As SByte()
      Return mText
    End Function

    ''' <summary>Sets the bytes that form the barcode. This bytes should
    ''' be interpreted in the codepage Cp437.
    ''' </summary>
    ''' <param name="text">the bytes that form the barcode
    ''' </param>
    Public Overridable Sub setText(ByVal text As SByte())
      mText = text
    End Sub

    ''' <summary>Sets the text that will form the barcode. This text is converted
    ''' to bytes using the encoding Cp437.
    ''' </summary>
    ''' <param name="s">the text that will form the barcode
    ''' @throws UnsupportedEncodingException if the encoding Cp437 is not supported
    ''' </param>
    Public Overridable Sub setText(ByVal s As String)
      mText = Filter(s)
    End Sub

    Public Function getMaskString() As ArrayList
      Dim Ar As New ArrayList

      Me.paintCode()

      Dim cols As Integer = (Me.BitColumns - 1) \ 8 + 1
      Dim sLine As String = ""
      Dim sHex As String
      Dim out_Renamed() As SByte = Me.OutBits
      For k As Integer = 0 To out_Renamed.Length - 1
        If ((k Mod cols) = 0) And k > 0 Then
          Ar.Add(sLine)
          sLine = ""
        End If
        'sHex = System.Convert.ToString((out_Renamed(k) And &HFF) Or &H100, 16).Substring(1).ToUpper()
        sHex = System.Convert.ToString((out_Renamed(k) And &HFF) Or &H100, 2).Substring(1).ToUpper()
        sLine += sHex
        If k >= out_Renamed.Length - 1 Then
          Ar.Add(sLine)
        End If
      Next

      Return Ar
    End Function


    Protected Friend Const START_PATTERN As Integer = &H1FEA8
    Protected Friend Const STOP_PATTERN As Integer = &H3FA29
    Protected Friend Const START_CODE_SIZE As Integer = 17
    Protected Friend Const STOP_SIZE As Integer = 18
    Protected Friend Const MODUL As Integer = 929
    Protected Friend Const ALPHA As Integer = &H10000
    Protected Friend Const LOWER As Integer = &H20000
    Protected Friend Const MIXED As Integer = &H40000
    Protected Friend Const PUNCTUATION As Integer = &H80000
    Protected Friend Const ISBYTE As Integer = &H100000
    Protected Friend Const BYTESHIFT As Integer = 913
    Protected Friend Const PL As Integer = 25
    Protected Friend Const LL As Integer = 27
    Protected Friend Const AS_ As Integer = 27
    Protected Friend Const ML As Integer = 28
    Protected Friend Const AL As Integer = 28
    Protected Friend Const PS As Integer = 29
    Protected Friend Const PAL As Integer = 29
    Protected Friend Const SPACE As Integer = 26
    Protected Friend Const TEXT_MODE As Integer = 900
    Protected Friend Const BYTE_MODE_6 As Integer = 924
    Protected Friend Const BYTE_MODE As Integer = 901
    Protected Friend Const NUMERIC_MODE As Integer = 902
    Protected Friend Const ABSOLUTE_MAX_TEXT_SIZE As Integer = 5420
    Protected Friend Const MAX_DATA_CODEWORDS As Integer = 926

    Private MIXED_SET As String = "0123456789&" & Chr(13) & Chr(9) & ",:#-.$/+%*=^"
    Private PUNCTUATION_SET As String = ";<>@[\]_`~!" & Chr(13) & Chr(9) & ",:" & Chr(10) & "-.$/""|*()?{}'"

    Private CLUSTERS()() As Integer = {New Integer() {&H1D5C0, &H1EAF0, &H1F57C, &H1D4E0, &H1EA78, &H1F53E, &H1A8C0, &H1D470, &H1A860, &H15040, &H1A830, &H15020, &H1ADC0, &H1D6F0, &H1EB7C, &H1ACE0, &H1D678, &H1EB3E, &H158C0, &H1AC70, &H15860, &H15DC0, &H1AEF0, &H1D77C, &H15CE0, &H1AE78, &H1D73E, &H15C70, &H1AE3C, &H15EF0, &H1AF7C, &H15E78, &H1AF3E, &H15F7C, &H1F5FA, &H1D2E0, &H1E978, &H1F4BE, &H1A4C0, &H1D270, &H1E93C, &H1A460, &H1D238, &H14840, &H1A430, &H1D21C, &H14820, &H1A418, &H14810, &H1A6E0, &H1D378, &H1E9BE, &H14CC0, &H1A670, &H1D33C, &H14C60, &H1A638, &H1D31E, &H14C30, &H1A61C, &H14EE0, &H1A778, &H1D3BE, &H14E70, &H1A73C, &H14E38, &H1A71E, &H14F78, &H1A7BE, &H14F3C, &H14F1E, &H1A2C0, &H1D170, &H1E8BC, &H1A260, &H1D138, &H1E89E, &H14440, &H1A230, &H1D11C, &H14420, &H1A218, &H14410, &H14408, &H146C0, &H1A370, &H1D1BC, &H14660, &H1A338, &H1D19E, &H14630, &H1A31C, &H14618, &H1460C, &H14770, &H1A3BC, &H14738, &H1A39E, &H1471C, &H147BC, &H1A160, &H1D0B8, &H1E85E, &H14240, &H1A130, &H1D09C, &H14220, &H1A118, &H1D08E, &H14210, &H1A10C, &H14208, &H1A106, &H14360, &H1A1B8, &H1D0DE, &H14330, &H1A19C, &H14318, &H1A18E, &H1430C, &H14306, &H1A1DE, &H1438E, &H14140, &H1A0B0, &H1D05C, &H14120, &H1A098, &H1D04E, &H14110, &H1A08C, &H14108, &H1A086, &H14104, &H141B0, &H14198, &H1418C, &H140A0, &H1D02E, &H1A04C, &H1A046, &H14082, &H1CAE0, &H1E578, &H1F2BE, &H194C0, &H1CA70, &H1E53C, &H19460, &H1CA38, &H1E51E, &H12840, &H19430, &H12820, &H196E0, &H1CB78, &H1E5BE, &H12CC0, &H19670, &H1CB3C, &H12C60, &H19638, &H12C30, &H12C18, &H12EE0, &H19778, &H1CBBE, &H12E70, &H1973C, &H12E38, &H12E1C, &H12F78, &H197BE, &H12F3C, &H12FBE, &H1DAC0, &H1ED70, &H1F6BC, &H1DA60, &H1ED38, &H1F69E, &H1B440, &H1DA30, &H1ED1C, &H1B420, &H1DA18, &H1ED0E, &H1B410, &H1DA0C, &H192C0, &H1C970, &H1E4BC, &H1B6C0, &H19260, &H1C938, &H1E49E, &H1B660, &H1DB38, &H1ED9E, &H16C40, &H12420, &H19218, &H1C90E, &H16C20, &H1B618, &H16C10, &H126C0, &H19370, &H1C9BC, &H16EC0, &H12660, &H19338, &H1C99E, &H16E60, &H1B738, &H1DB9E, &H16E30, &H12618, &H16E18, &H12770, _
    &H193BC, &H16F70, &H12738, &H1939E, &H16F38, &H1B79E, &H16F1C, &H127BC, &H16FBC, &H1279E, &H16F9E, &H1D960, &H1ECB8, &H1F65E, &H1B240, &H1D930, &H1EC9C, &H1B220, &H1D918, &H1EC8E, &H1B210, &H1D90C, &H1B208, &H1B204, &H19160, &H1C8B8, &H1E45E, &H1B360, &H19130, &H1C89C, &H16640, &H12220, &H1D99C, &H1C88E, &H16620, &H12210, &H1910C, &H16610, &H1B30C, &H19106, &H12204, &H12360, &H191B8, &H1C8DE, &H16760, &H12330, &H1919C, &H16730, &H1B39C, &H1918E, &H16718, &H1230C, &H12306, &H123B8, &H191DE, &H167B8, &H1239C, &H1679C, &H1238E, &H1678E, &H167DE, &H1B140, &H1D8B0, &H1EC5C, &H1B120, &H1D898, &H1EC4E, &H1B110, &H1D88C, &H1B108, &H1D886, &H1B104, &H1B102, &H12140, &H190B0, &H1C85C, &H16340, &H12120, &H19098, &H1C84E, &H16320, &H1B198, &H1D8CE, &H16310, &H12108, &H19086, &H16308, &H1B186, &H16304, &H121B0, &H190DC, &H163B0, &H12198, &H190CE, &H16398, &H1B1CE, &H1638C, &H12186, &H16386, &H163DC, &H163CE, &H1B0A0, &H1D858, &H1EC2E, &H1B090, &H1D84C, &H1B088, &H1D846, &H1B084, &H1B082, &H120A0, &H19058, &H1C82E, &H161A0, &H12090, &H1904C, &H16190, &H1B0CC, &H19046, &H16188, &H12084, &H16184, &H12082, &H120D8, &H161D8, &H161CC, &H161C6, &H1D82C, &H1D826, &H1B042, &H1902C, &H12048, &H160C8, &H160C4, &H160C2, &H18AC0, &H1C570, &H1E2BC, &H18A60, &H1C538, &H11440, &H18A30, &H1C51C, &H11420, &H18A18, &H11410, &H11408, &H116C0, &H18B70, &H1C5BC, &H11660, &H18B38, &H1C59E, &H11630, &H18B1C, &H11618, &H1160C, &H11770, &H18BBC, &H11738, &H18B9E, &H1171C, &H117BC, &H1179E, &H1CD60, &H1E6B8, &H1F35E, &H19A40, &H1CD30, &H1E69C, &H19A20, &H1CD18, &H1E68E, &H19A10, &H1CD0C, &H19A08, &H1CD06, &H18960, &H1C4B8, &H1E25E, &H19B60, &H18930, &H1C49C, &H13640, &H11220, &H1CD9C, &H1C48E, &H13620, &H19B18, &H1890C, &H13610, &H11208, &H13608, &H11360, &H189B8, &H1C4DE, &H13760, &H11330, &H1CDDE, &H13730, &H19B9C, &H1898E, &H13718, &H1130C, &H1370C, &H113B8, &H189DE, &H137B8, &H1139C, &H1379C, &H1138E, &H113DE, &H137DE, &H1DD40, &H1EEB0, &H1F75C, &H1DD20, &H1EE98, &H1F74E, &H1DD10, &H1EE8C, &H1DD08, &H1EE86, &H1DD04, &H19940, &H1CCB0, _
    &H1E65C, &H1BB40, &H19920, &H1EEDC, &H1E64E, &H1BB20, &H1DD98, &H1EECE, &H1BB10, &H19908, &H1CC86, &H1BB08, &H1DD86, &H19902, &H11140, &H188B0, &H1C45C, &H13340, &H11120, &H18898, &H1C44E, &H17740, &H13320, &H19998, &H1CCCE, &H17720, &H1BB98, &H1DDCE, &H18886, &H17710, &H13308, &H19986, &H17708, &H11102, &H111B0, &H188DC, &H133B0, &H11198, &H188CE, &H177B0, &H13398, &H199CE, &H17798, &H1BBCE, &H11186, &H13386, &H111DC, &H133DC, &H111CE, &H177DC, &H133CE, &H1DCA0, &H1EE58, &H1F72E, &H1DC90, &H1EE4C, &H1DC88, &H1EE46, &H1DC84, &H1DC82, &H198A0, &H1CC58, &H1E62E, &H1B9A0, &H19890, &H1EE6E, &H1B990, &H1DCCC, &H1CC46, &H1B988, &H19884, &H1B984, &H19882, &H1B982, &H110A0, &H18858, &H1C42E, &H131A0, &H11090, &H1884C, &H173A0, &H13190, &H198CC, &H18846, &H17390, &H1B9CC, &H11084, &H17388, &H13184, &H11082, &H13182, &H110D8, &H1886E, &H131D8, &H110CC, &H173D8, &H131CC, &H110C6, &H173CC, &H131C6, &H110EE, &H173EE, &H1DC50, &H1EE2C, &H1DC48, &H1EE26, &H1DC44, &H1DC42, &H19850, &H1CC2C, &H1B8D0, &H19848, &H1CC26, &H1B8C8, &H1DC66, &H1B8C4, &H19842, &H1B8C2, &H11050, &H1882C, &H130D0, &H11048, &H18826, &H171D0, &H130C8, &H19866, &H171C8, &H1B8E6, &H11042, &H171C4, &H130C2, &H171C2, &H130EC, &H171EC, &H171E6, &H1EE16, &H1DC22, &H1CC16, &H19824, &H19822, &H11028, &H13068, &H170E8, &H11022, &H13062, &H18560, &H10A40, &H18530, &H10A20, &H18518, &H1C28E, &H10A10, &H1850C, &H10A08, &H18506, &H10B60, &H185B8, &H1C2DE, &H10B30, &H1859C, &H10B18, &H1858E, &H10B0C, &H10B06, &H10BB8, &H185DE, &H10B9C, &H10B8E, &H10BDE, &H18D40, &H1C6B0, &H1E35C, &H18D20, &H1C698, &H18D10, &H1C68C, &H18D08, &H1C686, &H18D04, &H10940, &H184B0, &H1C25C, &H11B40, &H10920, &H1C6DC, &H1C24E, &H11B20, &H18D98, &H1C6CE, &H11B10, &H10908, &H18486, &H11B08, &H18D86, &H10902, &H109B0, &H184DC, &H11BB0, &H10998, &H184CE, &H11B98, &H18DCE, &H11B8C, &H10986, &H109DC, &H11BDC, &H109CE, &H11BCE, &H1CEA0, &H1E758, &H1F3AE, &H1CE90, &H1E74C, &H1CE88, &H1E746, &H1CE84, &H1CE82, &H18CA0, &H1C658, &H19DA0, &H18C90, &H1C64C, &H19D90, &H1CECC, &H1C646, &H19D88, _
    &H18C84, &H19D84, &H18C82, &H19D82, &H108A0, &H18458, &H119A0, &H10890, &H1C66E, &H13BA0, &H11990, &H18CCC, &H18446, &H13B90, &H19DCC, &H10884, &H13B88, &H11984, &H10882, &H11982, &H108D8, &H1846E, &H119D8, &H108CC, &H13BD8, &H119CC, &H108C6, &H13BCC, &H119C6, &H108EE, &H119EE, &H13BEE, &H1EF50, &H1F7AC, &H1EF48, &H1F7A6, &H1EF44, &H1EF42, &H1CE50, &H1E72C, &H1DED0, &H1EF6C, &H1E726, &H1DEC8, &H1EF66, &H1DEC4, &H1CE42, &H1DEC2, &H18C50, &H1C62C, &H19CD0, &H18C48, &H1C626, &H1BDD0, &H19CC8, &H1CE66, &H1BDC8, &H1DEE6, &H18C42, &H1BDC4, &H19CC2, &H1BDC2, &H10850, &H1842C, &H118D0, &H10848, &H18426, &H139D0, &H118C8, &H18C66, &H17BD0, &H139C8, &H19CE6, &H10842, &H17BC8, &H1BDE6, &H118C2, &H17BC4, &H1086C, &H118EC, &H10866, &H139EC, &H118E6, &H17BEC, &H139E6, &H17BE6, &H1EF28, &H1F796, &H1EF24, &H1EF22, &H1CE28, &H1E716, &H1DE68, &H1EF36, &H1DE64, &H1CE22, &H1DE62, &H18C28, &H1C616, &H19C68, &H18C24, &H1BCE8, &H19C64, &H18C22, &H1BCE4, &H19C62, &H1BCE2, &H10828, &H18416, &H11868, &H18C36, &H138E8, &H11864, &H10822, &H179E8, &H138E4, &H11862, &H179E4, &H138E2, &H179E2, &H11876, &H179F6, &H1EF12, &H1DE34, &H1DE32, &H19C34, &H1BC74, &H1BC72, &H11834, &H13874, &H178F4, &H178F2, &H10540, &H10520, &H18298, &H10510, &H10508, &H10504, &H105B0, &H10598, &H1058C, &H10586, &H105DC, &H105CE, &H186A0, &H18690, &H1C34C, &H18688, &H1C346, &H18684, &H18682, &H104A0, &H18258, &H10DA0, &H186D8, &H1824C, &H10D90, &H186CC, &H10D88, &H186C6, &H10D84, &H10482, &H10D82, &H104D8, &H1826E, &H10DD8, &H186EE, &H10DCC, &H104C6, &H10DC6, &H104EE, &H10DEE, &H1C750, &H1C748, &H1C744, &H1C742, &H18650, &H18ED0, &H1C76C, &H1C326, &H18EC8, &H1C766, &H18EC4, &H18642, &H18EC2, &H10450, &H10CD0, &H10448, &H18226, &H11DD0, &H10CC8, &H10444, &H11DC8, &H10CC4, &H10442, &H11DC4, &H10CC2, &H1046C, &H10CEC, &H10466, &H11DEC, &H10CE6, &H11DE6, &H1E7A8, &H1E7A4, &H1E7A2, &H1C728, &H1CF68, &H1E7B6, &H1CF64, &H1C722, &H1CF62, &H18628, &H1C316, &H18E68, &H1C736, &H19EE8, &H18E64, &H18622, &H19EE4, &H18E62, &H19EE2, &H10428, &H18216, &H10C68, &H18636, _
    &H11CE8, &H10C64, &H10422, &H13DE8, &H11CE4, &H10C62, &H13DE4, &H11CE2, &H10436, &H10C76, &H11CF6, &H13DF6, &H1F7D4, &H1F7D2, &H1E794, &H1EFB4, &H1E792, &H1EFB2, &H1C714, &H1CF34, &H1C712, &H1DF74, &H1CF32, &H1DF72, &H18614, &H18E34, &H18612, &H19E74, &H18E32, &H1BEF4}, New Integer() {&H1F560, &H1FAB8, &H1EA40, &H1F530, &H1FA9C, &H1EA20, &H1F518, &H1FA8E, &H1EA10, &H1F50C, &H1EA08, &H1F506, &H1EA04, &H1EB60, &H1F5B8, &H1FADE, &H1D640, &H1EB30, &H1F59C, &H1D620, &H1EB18, &H1F58E, &H1D610, &H1EB0C, &H1D608, &H1EB06, &H1D604, &H1D760, &H1EBB8, &H1F5DE, &H1AE40, &H1D730, &H1EB9C, &H1AE20, &H1D718, &H1EB8E, &H1AE10, &H1D70C, &H1AE08, &H1D706, &H1AE04, &H1AF60, &H1D7B8, &H1EBDE, &H15E40, &H1AF30, &H1D79C, &H15E20, &H1AF18, &H1D78E, &H15E10, &H1AF0C, &H15E08, &H1AF06, &H15F60, &H1AFB8, &H1D7DE, &H15F30, &H1AF9C, &H15F18, &H1AF8E, &H15F0C, &H15FB8, &H1AFDE, &H15F9C, &H15F8E, &H1E940, &H1F4B0, &H1FA5C, &H1E920, &H1F498, &H1FA4E, &H1E910, &H1F48C, &H1E908, &H1F486, &H1E904, &H1E902, &H1D340, &H1E9B0, &H1F4DC, &H1D320, &H1E998, &H1F4CE, &H1D310, &H1E98C, &H1D308, &H1E986, &H1D304, &H1D302, &H1A740, &H1D3B0, &H1E9DC, &H1A720, &H1D398, &H1E9CE, &H1A710, &H1D38C, &H1A708, &H1D386, &H1A704, &H1A702, &H14F40, &H1A7B0, &H1D3DC, &H14F20, &H1A798, &H1D3CE, &H14F10, &H1A78C, &H14F08, &H1A786, &H14F04, &H14FB0, &H1A7DC, &H14F98, &H1A7CE, &H14F8C, &H14F86, &H14FDC, &H14FCE, &H1E8A0, &H1F458, &H1FA2E, &H1E890, &H1F44C, &H1E888, &H1F446, &H1E884, &H1E882, &H1D1A0, &H1E8D8, &H1F46E, &H1D190, &H1E8CC, &H1D188, &H1E8C6, &H1D184, &H1D182, &H1A3A0, &H1D1D8, &H1E8EE, &H1A390, &H1D1CC, &H1A388, &H1D1C6, &H1A384, &H1A382, &H147A0, &H1A3D8, &H1D1EE, &H14790, &H1A3CC, &H14788, &H1A3C6, &H14784, &H14782, &H147D8, &H1A3EE, &H147CC, &H147C6, &H147EE, &H1E850, &H1F42C, &H1E848, &H1F426, &H1E844, &H1E842, &H1D0D0, &H1E86C, &H1D0C8, &H1E866, &H1D0C4, &H1D0C2, &H1A1D0, &H1D0EC, &H1A1C8, &H1D0E6, &H1A1C4, &H1A1C2, &H143D0, &H1A1EC, &H143C8, &H1A1E6, &H143C4, &H143C2, &H143EC, &H143E6, &H1E828, &H1F416, &H1E824, &H1E822, &H1D068, &H1E836, &H1D064, _
    &H1D062, &H1A0E8, &H1D076, &H1A0E4, &H1A0E2, &H141E8, &H1A0F6, &H141E4, &H141E2, &H1E814, &H1E812, &H1D034, &H1D032, &H1A074, &H1A072, &H1E540, &H1F2B0, &H1F95C, &H1E520, &H1F298, &H1F94E, &H1E510, &H1F28C, &H1E508, &H1F286, &H1E504, &H1E502, &H1CB40, &H1E5B0, &H1F2DC, &H1CB20, &H1E598, &H1F2CE, &H1CB10, &H1E58C, &H1CB08, &H1E586, &H1CB04, &H1CB02, &H19740, &H1CBB0, &H1E5DC, &H19720, &H1CB98, &H1E5CE, &H19710, &H1CB8C, &H19708, &H1CB86, &H19704, &H19702, &H12F40, &H197B0, &H1CBDC, &H12F20, &H19798, &H1CBCE, &H12F10, &H1978C, &H12F08, &H19786, &H12F04, &H12FB0, &H197DC, &H12F98, &H197CE, &H12F8C, &H12F86, &H12FDC, &H12FCE, &H1F6A0, &H1FB58, &H16BF0, &H1F690, &H1FB4C, &H169F8, &H1F688, &H1FB46, &H168FC, &H1F684, &H1F682, &H1E4A0, &H1F258, &H1F92E, &H1EDA0, &H1E490, &H1FB6E, &H1ED90, &H1F6CC, &H1F246, &H1ED88, &H1E484, &H1ED84, &H1E482, &H1ED82, &H1C9A0, &H1E4D8, &H1F26E, &H1DBA0, &H1C990, &H1E4CC, &H1DB90, &H1EDCC, &H1E4C6, &H1DB88, &H1C984, &H1DB84, &H1C982, &H1DB82, &H193A0, &H1C9D8, &H1E4EE, &H1B7A0, &H19390, &H1C9CC, &H1B790, &H1DBCC, &H1C9C6, &H1B788, &H19384, &H1B784, &H19382, &H1B782, &H127A0, &H193D8, &H1C9EE, &H16FA0, &H12790, &H193CC, &H16F90, &H1B7CC, &H193C6, &H16F88, &H12784, &H16F84, &H12782, &H127D8, &H193EE, &H16FD8, &H127CC, &H16FCC, &H127C6, &H16FC6, &H127EE, &H1F650, &H1FB2C, &H165F8, &H1F648, &H1FB26, &H164FC, &H1F644, &H1647E, &H1F642, &H1E450, &H1F22C, &H1ECD0, &H1E448, &H1F226, &H1ECC8, &H1F666, &H1ECC4, &H1E442, &H1ECC2, &H1C8D0, &H1E46C, &H1D9D0, &H1C8C8, &H1E466, &H1D9C8, &H1ECE6, &H1D9C4, &H1C8C2, &H1D9C2, &H191D0, &H1C8EC, &H1B3D0, &H191C8, &H1C8E6, &H1B3C8, &H1D9E6, &H1B3C4, &H191C2, &H1B3C2, &H123D0, &H191EC, &H167D0, &H123C8, &H191E6, &H167C8, &H1B3E6, &H167C4, &H123C2, &H167C2, &H123EC, &H167EC, &H123E6, &H167E6, &H1F628, &H1FB16, &H162FC, &H1F624, &H1627E, &H1F622, &H1E428, &H1F216, &H1EC68, &H1F636, &H1EC64, &H1E422, &H1EC62, &H1C868, &H1E436, &H1D8E8, &H1C864, &H1D8E4, &H1C862, &H1D8E2, &H190E8, &H1C876, &H1B1E8, &H1D8F6, &H1B1E4, &H190E2, &H1B1E2, &H121E8, &H190F6, _
    &H163E8, &H121E4, &H163E4, &H121E2, &H163E2, &H121F6, &H163F6, &H1F614, &H1617E, &H1F612, &H1E414, &H1EC34, &H1E412, &H1EC32, &H1C834, &H1D874, &H1C832, &H1D872, &H19074, &H1B0F4, &H19072, &H1B0F2, &H120F4, &H161F4, &H120F2, &H161F2, &H1F60A, &H1E40A, &H1EC1A, &H1C81A, &H1D83A, &H1903A, &H1B07A, &H1E2A0, &H1F158, &H1F8AE, &H1E290, &H1F14C, &H1E288, &H1F146, &H1E284, &H1E282, &H1C5A0, &H1E2D8, &H1F16E, &H1C590, &H1E2CC, &H1C588, &H1E2C6, &H1C584, &H1C582, &H18BA0, &H1C5D8, &H1E2EE, &H18B90, &H1C5CC, &H18B88, &H1C5C6, &H18B84, &H18B82, &H117A0, &H18BD8, &H1C5EE, &H11790, &H18BCC, &H11788, &H18BC6, &H11784, &H11782, &H117D8, &H18BEE, &H117CC, &H117C6, &H117EE, &H1F350, &H1F9AC, &H135F8, &H1F348, &H1F9A6, &H134FC, &H1F344, &H1347E, &H1F342, &H1E250, &H1F12C, &H1E6D0, &H1E248, &H1F126, &H1E6C8, &H1F366, &H1E6C4, &H1E242, &H1E6C2, &H1C4D0, &H1E26C, &H1CDD0, &H1C4C8, &H1E266, &H1CDC8, &H1E6E6, &H1CDC4, &H1C4C2, &H1CDC2, &H189D0, &H1C4EC, &H19BD0, &H189C8, &H1C4E6, &H19BC8, &H1CDE6, &H19BC4, &H189C2, &H19BC2, &H113D0, &H189EC, &H137D0, &H113C8, &H189E6, &H137C8, &H19BE6, &H137C4, &H113C2, &H137C2, &H113EC, &H137EC, &H113E6, &H137E6, &H1FBA8, &H175F0, &H1BAFC, &H1FBA4, &H174F8, &H1BA7E, &H1FBA2, &H1747C, &H1743E, &H1F328, &H1F996, &H132FC, &H1F768, &H1FBB6, &H176FC, &H1327E, &H1F764, &H1F322, &H1767E, &H1F762, &H1E228, &H1F116, &H1E668, &H1E224, &H1EEE8, &H1F776, &H1E222, &H1EEE4, &H1E662, &H1EEE2, &H1C468, &H1E236, &H1CCE8, &H1C464, &H1DDE8, &H1CCE4, &H1C462, &H1DDE4, &H1CCE2, &H1DDE2, &H188E8, &H1C476, &H199E8, &H188E4, &H1BBE8, &H199E4, &H188E2, &H1BBE4, &H199E2, &H1BBE2, &H111E8, &H188F6, &H133E8, &H111E4, &H177E8, &H133E4, &H111E2, &H177E4, &H133E2, &H177E2, &H111F6, &H133F6, &H1FB94, &H172F8, &H1B97E, &H1FB92, &H1727C, &H1723E, &H1F314, &H1317E, &H1F734, &H1F312, &H1737E, &H1F732, &H1E214, &H1E634, &H1E212, &H1EE74, &H1E632, &H1EE72, &H1C434, &H1CC74, &H1C432, &H1DCF4, &H1CC72, &H1DCF2, &H18874, &H198F4, &H18872, &H1B9F4, &H198F2, &H1B9F2, &H110F4, &H131F4, &H110F2, &H173F4, &H131F2, &H173F2, &H1FB8A, _
    &H1717C, &H1713E, &H1F30A, &H1F71A, &H1E20A, &H1E61A, &H1EE3A, &H1C41A, &H1CC3A, &H1DC7A, &H1883A, &H1987A, &H1B8FA, &H1107A, &H130FA, &H171FA, &H170BE, &H1E150, &H1F0AC, &H1E148, &H1F0A6, &H1E144, &H1E142, &H1C2D0, &H1E16C, &H1C2C8, &H1E166, &H1C2C4, &H1C2C2, &H185D0, &H1C2EC, &H185C8, &H1C2E6, &H185C4, &H185C2, &H10BD0, &H185EC, &H10BC8, &H185E6, &H10BC4, &H10BC2, &H10BEC, &H10BE6, &H1F1A8, &H1F8D6, &H11AFC, &H1F1A4, &H11A7E, &H1F1A2, &H1E128, &H1F096, &H1E368, &H1E124, &H1E364, &H1E122, &H1E362, &H1C268, &H1E136, &H1C6E8, &H1C264, &H1C6E4, &H1C262, &H1C6E2, &H184E8, &H1C276, &H18DE8, &H184E4, &H18DE4, &H184E2, &H18DE2, &H109E8, &H184F6, &H11BE8, &H109E4, &H11BE4, &H109E2, &H11BE2, &H109F6, &H11BF6, &H1F9D4, &H13AF8, &H19D7E, &H1F9D2, &H13A7C, &H13A3E, &H1F194, &H1197E, &H1F3B4, &H1F192, &H13B7E, &H1F3B2, &H1E114, &H1E334, &H1E112, &H1E774, &H1E332, &H1E772, &H1C234, &H1C674, &H1C232, &H1CEF4, &H1C672, &H1CEF2, &H18474, &H18CF4, &H18472, &H19DF4, &H18CF2, &H19DF2, &H108F4, &H119F4, &H108F2, &H13BF4, &H119F2, &H13BF2, &H17AF0, &H1BD7C, &H17A78, &H1BD3E, &H17A3C, &H17A1E, &H1F9CA, &H1397C, &H1FBDA, &H17B7C, &H1393E, &H17B3E, &H1F18A, &H1F39A, &H1F7BA, &H1E10A, &H1E31A, &H1E73A, &H1EF7A, &H1C21A, &H1C63A, &H1CE7A, &H1DEFA, &H1843A, &H18C7A, &H19CFA, &H1BDFA, &H1087A, &H118FA, &H139FA, &H17978, &H1BCBE, &H1793C, &H1791E, &H138BE, &H179BE, &H178BC, &H1789E, &H1785E, &H1E0A8, &H1E0A4, &H1E0A2, &H1C168, &H1E0B6, &H1C164, &H1C162, &H182E8, &H1C176, &H182E4, &H182E2, &H105E8, &H182F6, &H105E4, &H105E2, &H105F6, &H1F0D4, &H10D7E, &H1F0D2, &H1E094, &H1E1B4, &H1E092, &H1E1B2, &H1C134, &H1C374, &H1C132, &H1C372, &H18274, &H186F4, &H18272, &H186F2, &H104F4, &H10DF4, &H104F2, &H10DF2, &H1F8EA, &H11D7C, &H11D3E, &H1F0CA, &H1F1DA, &H1E08A, &H1E19A, &H1E3BA, &H1C11A, &H1C33A, &H1C77A, &H1823A, &H1867A, &H18EFA, &H1047A, &H10CFA, &H11DFA, &H13D78, &H19EBE, &H13D3C, &H13D1E, &H11CBE, &H13DBE, &H17D70, &H1BEBC, &H17D38, &H1BE9E, &H17D1C, &H17D0E, &H13CBC, &H17DBC, &H13C9E, &H17D9E, &H17CB8, &H1BE5E, &H17C9C, &H17C8E, _
    &H13C5E, &H17CDE, &H17C5C, &H17C4E, &H17C2E, &H1C0B4, &H1C0B2, &H18174, &H18172, &H102F4, &H102F2, &H1E0DA, &H1C09A, &H1C1BA, &H1813A, &H1837A, &H1027A, &H106FA, &H10EBE, &H11EBC, &H11E9E, &H13EB8, &H19F5E, &H13E9C, &H13E8E, &H11E5E, &H13EDE, &H17EB0, &H1BF5C, &H17E98, &H1BF4E, &H17E8C, &H17E86, &H13E5C, &H17EDC, &H13E4E, &H17ECE, &H17E58, &H1BF2E, &H17E4C, &H17E46, &H13E2E, &H17E6E, &H17E2C, &H17E26, &H10F5E, &H11F5C, &H11F4E, &H13F58, &H19FAE, &H13F4C, &H13F46, &H11F2E, &H13F6E, &H13F2C, &H13F26}, New Integer() {&H1ABE0, &H1D5F8, &H153C0, &H1A9F0, &H1D4FC, &H151E0, &H1A8F8, &H1D47E, &H150F0, &H1A87C, &H15078, &H1FAD0, &H15BE0, &H1ADF8, &H1FAC8, &H159F0, &H1ACFC, &H1FAC4, &H158F8, &H1AC7E, &H1FAC2, &H1587C, &H1F5D0, &H1FAEC, &H15DF8, &H1F5C8, &H1FAE6, &H15CFC, &H1F5C4, &H15C7E, &H1F5C2, &H1EBD0, &H1F5EC, &H1EBC8, &H1F5E6, &H1EBC4, &H1EBC2, &H1D7D0, &H1EBEC, &H1D7C8, &H1EBE6, &H1D7C4, &H1D7C2, &H1AFD0, &H1D7EC, &H1AFC8, &H1D7E6, &H1AFC4, &H14BC0, &H1A5F0, &H1D2FC, &H149E0, &H1A4F8, &H1D27E, &H148F0, &H1A47C, &H14878, &H1A43E, &H1483C, &H1FA68, &H14DF0, &H1A6FC, &H1FA64, &H14CF8, &H1A67E, &H1FA62, &H14C7C, &H14C3E, &H1F4E8, &H1FA76, &H14EFC, &H1F4E4, &H14E7E, &H1F4E2, &H1E9E8, &H1F4F6, &H1E9E4, &H1E9E2, &H1D3E8, &H1E9F6, &H1D3E4, &H1D3E2, &H1A7E8, &H1D3F6, &H1A7E4, &H1A7E2, &H145E0, &H1A2F8, &H1D17E, &H144F0, &H1A27C, &H14478, &H1A23E, &H1443C, &H1441E, &H1FA34, &H146F8, &H1A37E, &H1FA32, &H1467C, &H1463E, &H1F474, &H1477E, &H1F472, &H1E8F4, &H1E8F2, &H1D1F4, &H1D1F2, &H1A3F4, &H1A3F2, &H142F0, &H1A17C, &H14278, &H1A13E, &H1423C, &H1421E, &H1FA1A, &H1437C, &H1433E, &H1F43A, &H1E87A, &H1D0FA, &H14178, &H1A0BE, &H1413C, &H1411E, &H141BE, &H140BC, &H1409E, &H12BC0, &H195F0, &H1CAFC, &H129E0, &H194F8, &H1CA7E, &H128F0, &H1947C, &H12878, &H1943E, &H1283C, &H1F968, &H12DF0, &H196FC, &H1F964, &H12CF8, &H1967E, &H1F962, &H12C7C, &H12C3E, &H1F2E8, &H1F976, &H12EFC, &H1F2E4, &H12E7E, &H1F2E2, &H1E5E8, &H1F2F6, &H1E5E4, &H1E5E2, &H1CBE8, &H1E5F6, &H1CBE4, &H1CBE2, &H197E8, &H1CBF6, &H197E4, &H197E2, &H1B5E0, &H1DAF8, _
    &H1ED7E, &H169C0, &H1B4F0, &H1DA7C, &H168E0, &H1B478, &H1DA3E, &H16870, &H1B43C, &H16838, &H1B41E, &H1681C, &H125E0, &H192F8, &H1C97E, &H16DE0, &H124F0, &H1927C, &H16CF0, &H1B67C, &H1923E, &H16C78, &H1243C, &H16C3C, &H1241E, &H16C1E, &H1F934, &H126F8, &H1937E, &H1FB74, &H1F932, &H16EF8, &H1267C, &H1FB72, &H16E7C, &H1263E, &H16E3E, &H1F274, &H1277E, &H1F6F4, &H1F272, &H16F7E, &H1F6F2, &H1E4F4, &H1EDF4, &H1E4F2, &H1EDF2, &H1C9F4, &H1DBF4, &H1C9F2, &H1DBF2, &H193F4, &H193F2, &H165C0, &H1B2F0, &H1D97C, &H164E0, &H1B278, &H1D93E, &H16470, &H1B23C, &H16438, &H1B21E, &H1641C, &H1640E, &H122F0, &H1917C, &H166F0, &H12278, &H1913E, &H16678, &H1B33E, &H1663C, &H1221E, &H1661E, &H1F91A, &H1237C, &H1FB3A, &H1677C, &H1233E, &H1673E, &H1F23A, &H1F67A, &H1E47A, &H1ECFA, &H1C8FA, &H1D9FA, &H191FA, &H162E0, &H1B178, &H1D8BE, &H16270, &H1B13C, &H16238, &H1B11E, &H1621C, &H1620E, &H12178, &H190BE, &H16378, &H1213C, &H1633C, &H1211E, &H1631E, &H121BE, &H163BE, &H16170, &H1B0BC, &H16138, &H1B09E, &H1611C, &H1610E, &H120BC, &H161BC, &H1209E, &H1619E, &H160B8, &H1B05E, &H1609C, &H1608E, &H1205E, &H160DE, &H1605C, &H1604E, &H115E0, &H18AF8, &H1C57E, &H114F0, &H18A7C, &H11478, &H18A3E, &H1143C, &H1141E, &H1F8B4, &H116F8, &H18B7E, &H1F8B2, &H1167C, &H1163E, &H1F174, &H1177E, &H1F172, &H1E2F4, &H1E2F2, &H1C5F4, &H1C5F2, &H18BF4, &H18BF2, &H135C0, &H19AF0, &H1CD7C, &H134E0, &H19A78, &H1CD3E, &H13470, &H19A3C, &H13438, &H19A1E, &H1341C, &H1340E, &H112F0, &H1897C, &H136F0, &H11278, &H1893E, &H13678, &H19B3E, &H1363C, &H1121E, &H1361E, &H1F89A, &H1137C, &H1F9BA, &H1377C, &H1133E, &H1373E, &H1F13A, &H1F37A, &H1E27A, &H1E6FA, &H1C4FA, &H1CDFA, &H189FA, &H1BAE0, &H1DD78, &H1EEBE, &H174C0, &H1BA70, &H1DD3C, &H17460, &H1BA38, &H1DD1E, &H17430, &H1BA1C, &H17418, &H1BA0E, &H1740C, &H132E0, &H19978, &H1CCBE, &H176E0, &H13270, &H1993C, &H17670, &H1BB3C, &H1991E, &H17638, &H1321C, &H1761C, &H1320E, &H1760E, &H11178, &H188BE, &H13378, &H1113C, &H17778, &H1333C, &H1111E, &H1773C, &H1331E, &H1771E, &H111BE, &H133BE, &H177BE, &H172C0, &H1B970, _
    &H1DCBC, &H17260, &H1B938, &H1DC9E, &H17230, &H1B91C, &H17218, &H1B90E, &H1720C, &H17206, &H13170, &H198BC, &H17370, &H13138, &H1989E, &H17338, &H1B99E, &H1731C, &H1310E, &H1730E, &H110BC, &H131BC, &H1109E, &H173BC, &H1319E, &H1739E, &H17160, &H1B8B8, &H1DC5E, &H17130, &H1B89C, &H17118, &H1B88E, &H1710C, &H17106, &H130B8, &H1985E, &H171B8, &H1309C, &H1719C, &H1308E, &H1718E, &H1105E, &H130DE, &H171DE, &H170B0, &H1B85C, &H17098, &H1B84E, &H1708C, &H17086, &H1305C, &H170DC, &H1304E, &H170CE, &H17058, &H1B82E, &H1704C, &H17046, &H1302E, &H1706E, &H1702C, &H17026, &H10AF0, &H1857C, &H10A78, &H1853E, &H10A3C, &H10A1E, &H10B7C, &H10B3E, &H1F0BA, &H1E17A, &H1C2FA, &H185FA, &H11AE0, &H18D78, &H1C6BE, &H11A70, &H18D3C, &H11A38, &H18D1E, &H11A1C, &H11A0E, &H10978, &H184BE, &H11B78, &H1093C, &H11B3C, &H1091E, &H11B1E, &H109BE, &H11BBE, &H13AC0, &H19D70, &H1CEBC, &H13A60, &H19D38, &H1CE9E, &H13A30, &H19D1C, &H13A18, &H19D0E, &H13A0C, &H13A06, &H11970, &H18CBC, &H13B70, &H11938, &H18C9E, &H13B38, &H1191C, &H13B1C, &H1190E, &H13B0E, &H108BC, &H119BC, &H1089E, &H13BBC, &H1199E, &H13B9E, &H1BD60, &H1DEB8, &H1EF5E, &H17A40, &H1BD30, &H1DE9C, &H17A20, &H1BD18, &H1DE8E, &H17A10, &H1BD0C, &H17A08, &H1BD06, &H17A04, &H13960, &H19CB8, &H1CE5E, &H17B60, &H13930, &H19C9C, &H17B30, &H1BD9C, &H19C8E, &H17B18, &H1390C, &H17B0C, &H13906, &H17B06, &H118B8, &H18C5E, &H139B8, &H1189C, &H17BB8, &H1399C, &H1188E, &H17B9C, &H1398E, &H17B8E, &H1085E, &H118DE, &H139DE, &H17BDE, &H17940, &H1BCB0, &H1DE5C, &H17920, &H1BC98, &H1DE4E, &H17910, &H1BC8C, &H17908, &H1BC86, &H17904, &H17902, &H138B0, &H19C5C, &H179B0, &H13898, &H19C4E, &H17998, &H1BCCE, &H1798C, &H13886, &H17986, &H1185C, &H138DC, &H1184E, &H179DC, &H138CE, &H179CE, &H178A0, &H1BC58, &H1DE2E, &H17890, &H1BC4C, &H17888, &H1BC46, &H17884, &H17882, &H13858, &H19C2E, &H178D8, &H1384C, &H178CC, &H13846, &H178C6, &H1182E, &H1386E, &H178EE, &H17850, &H1BC2C, &H17848, &H1BC26, &H17844, &H17842, &H1382C, &H1786C, &H13826, &H17866, &H17828, &H1BC16, &H17824, &H17822, &H13816, &H17836, _
    &H10578, &H182BE, &H1053C, &H1051E, &H105BE, &H10D70, &H186BC, &H10D38, &H1869E, &H10D1C, &H10D0E, &H104BC, &H10DBC, &H1049E, &H10D9E, &H11D60, &H18EB8, &H1C75E, &H11D30, &H18E9C, &H11D18, &H18E8E, &H11D0C, &H11D06, &H10CB8, &H1865E, &H11DB8, &H10C9C, &H11D9C, &H10C8E, &H11D8E, &H1045E, &H10CDE, &H11DDE, &H13D40, &H19EB0, &H1CF5C, &H13D20, &H19E98, &H1CF4E, &H13D10, &H19E8C, &H13D08, &H19E86, &H13D04, &H13D02, &H11CB0, &H18E5C, &H13DB0, &H11C98, &H18E4E, &H13D98, &H19ECE, &H13D8C, &H11C86, &H13D86, &H10C5C, &H11CDC, &H10C4E, &H13DDC, &H11CCE, &H13DCE, &H1BEA0, &H1DF58, &H1EFAE, &H1BE90, &H1DF4C, &H1BE88, &H1DF46, &H1BE84, &H1BE82, &H13CA0, &H19E58, &H1CF2E, &H17DA0, &H13C90, &H19E4C, &H17D90, &H1BECC, &H19E46, &H17D88, &H13C84, &H17D84, &H13C82, &H17D82, &H11C58, &H18E2E, &H13CD8, &H11C4C, &H17DD8, &H13CCC, &H11C46, &H17DCC, &H13CC6, &H17DC6, &H10C2E, &H11C6E, &H13CEE, &H17DEE, &H1BE50, &H1DF2C, &H1BE48, &H1DF26, &H1BE44, &H1BE42, &H13C50, &H19E2C, &H17CD0, &H13C48, &H19E26, &H17CC8, &H1BE66, &H17CC4, &H13C42, &H17CC2, &H11C2C, &H13C6C, &H11C26, &H17CEC, &H13C66, &H17CE6, &H1BE28, &H1DF16, &H1BE24, &H1BE22, &H13C28, &H19E16, &H17C68, &H13C24, &H17C64, &H13C22, &H17C62, &H11C16, &H13C36, &H17C76, &H1BE14, &H1BE12, &H13C14, &H17C34, &H13C12, &H17C32, &H102BC, &H1029E, &H106B8, &H1835E, &H1069C, &H1068E, &H1025E, &H106DE, &H10EB0, &H1875C, &H10E98, &H1874E, &H10E8C, &H10E86, &H1065C, &H10EDC, &H1064E, &H10ECE, &H11EA0, &H18F58, &H1C7AE, &H11E90, &H18F4C, &H11E88, &H18F46, &H11E84, &H11E82, &H10E58, &H1872E, &H11ED8, &H18F6E, &H11ECC, &H10E46, &H11EC6, &H1062E, &H10E6E, &H11EEE, &H19F50, &H1CFAC, &H19F48, &H1CFA6, &H19F44, &H19F42, &H11E50, &H18F2C, &H13ED0, &H19F6C, &H18F26, &H13EC8, &H11E44, &H13EC4, &H11E42, &H13EC2, &H10E2C, &H11E6C, &H10E26, &H13EEC, &H11E66, &H13EE6, &H1DFA8, &H1EFD6, &H1DFA4, &H1DFA2, &H19F28, &H1CF96, &H1BF68, &H19F24, &H1BF64, &H19F22, &H1BF62, &H11E28, &H18F16, &H13E68, &H11E24, &H17EE8, &H13E64, &H11E22, &H17EE4, &H13E62, &H17EE2, &H10E16, &H11E36, &H13E76, &H17EF6, &H1DF94, _
    &H1DF92, &H19F14, &H1BF34, &H19F12, &H1BF32, &H11E14, &H13E34, &H11E12, &H17E74, &H13E32, &H17E72, &H1DF8A, &H19F0A, &H1BF1A, &H11E0A, &H13E1A, &H17E3A, &H1035C, &H1034E, &H10758, &H183AE, &H1074C, &H10746, &H1032E, &H1076E, &H10F50, &H187AC, &H10F48, &H187A6, &H10F44, &H10F42, &H1072C, &H10F6C, &H10726, &H10F66, &H18FA8, &H1C7D6, &H18FA4, &H18FA2, &H10F28, &H18796, &H11F68, &H18FB6, &H11F64, &H10F22, &H11F62, &H10716, &H10F36, &H11F76, &H1CFD4, &H1CFD2, &H18F94, &H19FB4, &H18F92, &H19FB2, &H10F14, &H11F34, &H10F12, &H13F74, &H11F32, &H13F72, &H1CFCA, &H18F8A, &H19F9A, &H10F0A, &H11F1A, &H13F3A, &H103AC, &H103A6, &H107A8, &H183D6, &H107A4, &H107A2, &H10396, &H107B6, &H187D4, &H187D2, &H10794, &H10FB4, &H10792, &H10FB2, &H1C7EA}}

    Private ERROR_LEVEL()() As Integer = {New Integer() {27, 917}, New Integer() {522, 568, 723, 809}, New Integer() {237, 308, 436, 284, 646, 653, 428, 379}, New Integer() {274, 562, 232, 755, 599, 524, 801, 132, 295, 116, 442, 428, 295, 42, 176, 65}, New Integer() {361, 575, 922, 525, 176, 586, 640, 321, 536, 742, 677, 742, 687, 284, 193, 517, 273, 494, 263, 147, 593, 800, 571, 320, 803, 133, 231, 390, 685, 330, 63, 410}, New Integer() {539, 422, 6, 93, 862, 771, 453, 106, 610, 287, 107, 505, 733, 877, 381, 612, 723, 476, 462, 172, 430, 609, 858, 822, 543, 376, 511, 400, 672, 762, 283, 184, 440, 35, 519, 31, 460, 594, 225, 535, 517, 352, 605, 158, 651, 201, 488, 502, 648, 733, 717, 83, 404, 97, 280, 771, 840, 629, 4, 381, 843, 623, 264, 543}, New Integer() {521, 310, 864, 547, 858, 580, 296, 379, 53, 779, 897, 444, 400, 925, 749, 415, 822, 93, 217, 208, 928, 244, 583, 620, 246, 148, 447, 631, 292, 908, 490, 704, 516, 258, 457, 907, 594, 723, 674, 292, 272, 96, 684, 432, 686, 606, 860, 569, 193, 219, 129, 186, 236, 287, 192, 775, 278, 173, 40, 379, 712, 463, 646, 776, 171, 491, 297, 763, 156, 732, 95, 270, 447, 90, 507, 48, 228, 821, 808, 898, 784, 663, 627, 378, 382, 262, 380, 602, 754, 336, 89, 614, 87, 432, 670, 616, 157, 374, 242, 726, 600, 269, 375, 898, 845, 454, 354, 130, 814, 587, 804, 34, 211, 330, 539, 297, 827, 865, 37, 517, 834, 315, 550, 86, 801, 4, 108, 539}, New Integer() {524, 894, 75, 766, 882, 857, 74, 204, 82, 586, 708, 250, 905, 786, 138, 720, 858, 194, 311, 913, 275, 190, 375, 850, 438, 733, 194, 280, 201, 280, 828, 757, 710, 814, 919, 89, 68, 569, 11, 204, 796, 605, 540, 913, 801, 700, 799, 137, 439, 418, 592, 668, 353, 859, 370, 694, 325, 240, 216, 257, 284, 549, 209, 884, 315, 70, 329, 793, 490, 274, 877, 162, 749, 812, 684, 461, 334, 376, 849, 521, 307, 291, 803, 712, 19, 358, 399, 908, 103, 511, 51, 8, 517, 225, 289, 470, 637, 731, 66, 255, 917, 269, 463, 830, 730, 433, 848, 585, 136, 538, 906, 90, 2, 290, 743, 199, 655, 903, 329, 49, 802, 580, 355, 588, 188, 462, 10, 134, 628, 320, 479, 130, 739, 71, 263, 318, 374, 601, _
     192, 605, 142, 673, 687, 234, 722, 384, 177, 752, 607, 640, 455, 193, 689, 707, 805, 641, 48, 60, 732, 621, 895, 544, 261, 852, 655, 309, 697, 755, 756, 60, 231, 773, 434, 421, 726, 528, 503, 118, 49, 795, 32, 144, 500, 238, 836, 394, 280, 566, 319, 9, 647, 550, 73, 914, 342, 126, 32, 681, 331, 792, 620, 60, 609, 441, 180, 791, 893, 754, 605, 383, 228, 749, 760, 213, 54, 297, 134, 54, 834, 299, 922, 191, 910, 532, 609, 829, 189, 20, 167, 29, 872, 449, 83, 402, 41, 656, 505, 579, 481, 173, 404, 251, 688, 95, 497, 555, 642, 543, 307, 159, 924, 558, 648, 55, 497, 10}, New Integer() {352, 77, 373, 504, 35, 599, 428, 207, 409, 574, 118, 498, 285, 380, 350, 492, 197, 265, 920, 155, 914, 299, 229, 643, 294, 871, 306, 88, 87, 193, 352, 781, 846, 75, 327, 520, 435, 543, 203, 666, 249, 346, 781, 621, 640, 268, 794, 534, 539, 781, 408, 390, 644, 102, 476, 499, 290, 632, 545, 37, 858, 916, 552, 41, 542, 289, 122, 272, 383, 800, 485, 98, 752, 472, 761, 107, 784, 860, 658, 741, 290, 204, 681, 407, 855, 85, 99, 62, 482, 180, 20, 297, 451, 593, 913, 142, 808, 684, 287, 536, 561, 76, 653, 899, 729, 567, 744, 390, 513, 192, 516, 258, 240, 518, 794, 395, 768, 848, 51, 610, 384, 168, 190, 826, 328, 596, 786, 303, 570, 381, 415, 641, 156, 237, 151, 429, 531, 207, 676, 710, 89, 168, 304, 402, 40, 708, 575, 162, 864, 229, 65, 861, 841, 512, 164, 477, 221, 92, 358, 785, 288, 357, 850, 836, 827, 736, 707, 94, 8, 494, 114, 521, 2, 499, 851, 543, 152, 729, 771, 95, 248, 361, 578, 323, 856, 797, 289, 51, 684, 466, 533, 820, 669, 45, 902, 452, 167, 342, 244, 173, 35, 463, 651, 51, 699, 591, 452, 578, 37, 124, 298, 332, 552, 43, 427, 119, 662, 777, 475, 850, 764, 364, 578, 911, 283, 711, 472, 420, 245, 288, 594, 394, 511, 327, 589, 777, 699, 688, 43, 408, 842, 383, 721, 521, 560, 644, 714, 559, 62, 145, 873, 663, 713, 159, 672, 729, 624, 59, 193, 417, 158, 209, 563, 564, 343, 693, 109, 608, 563, 365, 181, 772, 677, 310, 248, 353, 708, 410, 579, 870, 617, 841, 632, 860, 289, 536, 35, 777, 618, 586, 424, 833, 77, 597, 346, 269, 757, 632, _
     695, 751, 331, 247, 184, 45, 787, 680, 18, 66, 407, 369, 54, 492, 228, 613, 830, 922, 437, 519, 644, 905, 789, 420, 305, 441, 207, 300, 892, 827, 141, 537, 381, 662, 513, 56, 252, 341, 242, 797, 838, 837, 720, 224, 307, 631, 61, 87, 560, 310, 756, 665, 397, 808, 851, 309, 473, 795, 378, 31, 647, 915, 459, 806, 590, 731, 425, 216, 548, 249, 321, 881, 699, 535, 673, 782, 210, 815, 905, 303, 843, 922, 281, 73, 469, 791, 660, 162, 498, 308, 155, 422, 907, 817, 187, 62, 16, 425, 535, 336, 286, 437, 375, 273, 610, 296, 183, 923, 116, 667, 751, 353, 62, 366, 691, 379, 687, 842, 37, 357, 720, 742, 330, 5, 39, 923, 311, 424, 242, 749, 321, 54, 669, 316, 342, 299, 534, 105, 667, 488, 640, 672, 576, 540, 316, 486, 721, 610, 46, 656, 447, 171, 616, 464, 190, 531, 297, 321, 762, 752, 533, 175, 134, 14, 381, 433, 717, 45, 111, 20, 596, 284, 736, 138, 646, 411, 877, 669, 141, 919, 45, 780, 407, 164, 332, 899, 165, 726, 600, 325, 498, 655, 357, 752, 768, 223, 849, 647, 63, 310, 863, 251, 366, 304, 282, 738, 675, 410, 389, 244, 31, 121, 303, 263}}

    ' <summary>Holds value of property outBits. </summary>
    Private mOutBits() As SByte

    ' <summary>Holds value of property bitColumns. </summary>
    Private mBitColumns As Integer

    ' <summary>Holds value of property codeRows. </summary>
    Private mCodeRows As Integer

    ' <summary>Holds value of property codeColumns. </summary>
    Private mCodeColumns As Integer

    ' <summary>Holds value of property codewords. </summary>
    Private mCodewords() As Integer

    ' <summary>Holds value of property lenCodewords. </summary>
    Private mLenCodewords As Integer

    ' <summary>Holds value of property errorLevel. </summary>
    Private mErrorLevel As Integer

    ' <summary>Holds value of property text. </summary>
    Private mText() As SByte

    ' <summary>Holds value of property options. </summary>
    Private mOptions As Integer

    ' <summary>Holds value of property aspectRatio. </summary>
    Private mAspectRatio As Double

    ' <summary>Holds value of property yHeight. </summary>
    Private mYHeight As Double

    Public Class Segment

      Public Sub New()
        MyBase.New()
      End Sub

      Public Sub New(ByVal enclosingInstance As Pdf417lib, ByVal type As Char, ByVal start As Integer, ByVal _end As Integer)
        InitBlock(enclosingInstance)
        Me.type = type
        Me.start = start
        Me.iEnd = _end
      End Sub

      Private Sub InitBlock(ByVal enclosingInstance As Pdf417lib)
        Me.enclosingInstance = enclosingInstance
      End Sub

      Private enclosingInstance As Pdf417lib
      Public ReadOnly Property Enclosing_Instance() As Pdf417lib
        Get
          Return enclosingInstance
        End Get
      End Property

      Public type As Char
      Public start As Integer
      Public iEnd As Integer

    End Class

    Protected Friend Class cSegmentList

      Public Sub New(ByVal enclosingInstance As Pdf417lib)
        InitBlock(enclosingInstance)
      End Sub

      Private Sub InitBlock(ByVal enclosingInstance As Pdf417lib)
        Me.enclosingInstance = enclosingInstance
        list = New System.Collections.ArrayList()
      End Sub

      Private enclosingInstance As Pdf417lib
      Public ReadOnly Property Enclosing_Instance() As Pdf417lib
        Get
          Return enclosingInstance
        End Get
      End Property

      Protected Friend list As System.Collections.ArrayList

      Public Overridable Sub add(ByVal type As Char, ByVal start As Integer, ByVal _end As Integer)
        list.Add(New Segment(enclosingInstance, type, start, _end))
      End Sub

      Public Overridable Function get_Renamed(ByVal idx As Integer) As Segment
        If (idx < 0 Or idx >= list.Count) Then Return Nothing
        Return CType(list(idx), Segment)
      End Function

      Public Overridable Sub remove(ByVal idx As Integer)
        If (idx < 0 Or idx >= list.Count) Then Return
        list.RemoveAt(idx)
      End Sub

      Public Overridable Function size() As Integer
        Return list.Count
      End Function
    End Class
  End Class



End Namespace

