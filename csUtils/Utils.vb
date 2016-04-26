Imports System.IO
Imports System.Drawing.Printing
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Drawing
Imports System.Net
Imports System.Net.Dns
Imports System.Globalization

Public Class Utils

#Region " Funcions DNI "
    Shared Function NifIsOk(ByVal NIF As String, ByRef CharControl As String) As Boolean
        If NIF.Substring(0, 1) = "X" Then
            ' Es un extranger
            NIF = NIF.Substring(1)
        End If
        If NIF.Substring(0, 1) = "Y" Then
            ' Es un extranger
            NIF = "1" + NIF.Substring(1)
        End If

        If Char.IsLetter(NIF.Chars(0)) Then
            If Char.IsLetter(Right(NIF, 1), 0) Then
                ' Es pública
                NIF = NIF.Substring(1)
                CharControl = NifGetNifDigit(NIF)
                NifIsOk = True
            Else
                ' Es persona juridica
                NIF = NIF.Substring(1)
                CharControl = NifGetNifDigit(NIF)
                NifIsOk = (CharControl = NIF.Substring(NIF.Length - 1, 1))
            End If
        Else
            ' Es una persona fisica
            If Char.IsLetter(Right(NIF, 1), 0) Then
                CharControl = (NifGetDniLetter(NIF.Substring(0, NIF.Length - 1)))
                NifIsOk = (CharControl = NIF.Substring(NIF.Length - 1, 1))
            Else
                CharControl = (NifGetDniLetter(NIF.Substring(0, NIF.Length)))
                NifIsOk = False
            End If
        End If
    End Function

    Shared Function NifIsOk(ByVal NIF As String) As Boolean
        Dim dummy As String
        Return NifIsOk(NIF, dummy)
    End Function


    Shared Function NifGetDniLetter(ByVal NIf As String) As String
        If Not IsDigit(NIf) Then
            NIf = ExtractDigits(NIf)
        End If
        Return "TRWAGMYFPDXBNJZSQVHLCKE".Chars(CInt(NIf) Mod 23).ToString
    End Function

    Shared Function NifGetNifDigit(ByVal NIf As String) As String
        Dim DniDigit As Integer = 0
        Dim i, nDigit As Integer

        If Not IsDigit(NIf) Then
            NIf = ExtractDigits(NIf)
        End If

        NIf = NIf.PadLeft(7, "0"c)
        For i = 1 To 7
            nDigit = CInt(NIf.Substring(i - 1, 1))
            If i Mod 2 = 0 Then
                DniDigit += nDigit * 9
            Else
                DniDigit += nDigit * 8
                If nDigit > 4 Then
                    DniDigit -= 1
                End If
            End If
        Next i
        Return CStr(DniDigit Mod 10)
    End Function
#End Region

#Region " Funcions STRING "
    Shared Function IsDigit(ByVal Value As String) As Boolean
        Dim c As Char
        If String.IsNullOrEmpty(Value) Then
            Return False
        End If
        For Each c In Value.ToCharArray
            If Not Char.IsDigit(c) Then
                Return False
            End If
        Next
        Return True
    End Function

    Shared Function ExtractDigits(ByVal Value As String) As String
        Dim Result As New StringBuilder(15)
        Dim c As Char
        For Each c In Value.ToCharArray
            If Char.IsDigit(c) Then
                Result.Append(c)
            End If
        Next
        Return Result.ToString
    End Function

    Shared Function Transform(ByVal Value As String, ByVal FormatString As String) As String
        Dim ndxValue, ndxFmt As Integer
        Dim sb As New System.Text.StringBuilder(FormatString.Length)
        ndxValue = 0
        For ndxFmt = 0 To FormatString.Length - 1
            If ndxValue = Value.Length Then Exit For
            If FormatString.Substring(ndxFmt, 1).IndexOfAny("xX9".ToCharArray) > -1 Then
                sb.Append(Value.Substring(ndxValue, 1))
                ndxValue += 1
            Else
                sb.Append(FormatString.Substring(ndxFmt, 1))
            End If
        Next
        Return sb.ToString
    End Function

    'Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Object) As Object
    '  If Value Is DBNull.Value OrElse Value Is Nothing Then
    '    Return DefaultValue
    '  End If
    '  Return Value
    'End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As DateTime) As DateTime
        If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
            Return DefaultValue
        End If
        Return Convert.ToDateTime(Value)
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Decimal) As Decimal
        If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
            Return DefaultValue
        End If
        Return Convert.ToDecimal(Value)
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Integer) As Integer
        If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Trim = "") Then
            Return DefaultValue
        End If
        Return Convert.ToInt32(Value)
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As String) As String
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return DefaultValue
        End If
        Return CStr(Value)
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Object) As Object
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return DefaultValue
        End If
        Return Value
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Guid) As Guid
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return DefaultValue
        End If
        Return CType(Value, Guid)
    End Function

    Shared Function CNull(ByVal Value As Object, ByVal DefaultValue As Boolean) As Boolean
        If Value Is DBNull.Value OrElse Value Is Nothing OrElse String.IsNullOrEmpty(CNull(Value, "")) Then
            Return DefaultValue
        End If
        Return CBool(Value)
    End Function

    Shared Function CNull(ByVal Value As Object) As String
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return ""
        End If
        Return CStr(Value)
    End Function

    Shared Function CDNull(ByVal Value As Object, Optional ByVal FormatString As String = "dd/MM/yyyy") As String
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return ""
        End If
        Return CDate(Value).ToString(FormatString)
    End Function

    Shared Function AddBackslash(ByVal Path As String) As String
        If Not Path.EndsWith("\") Then
            Return Path & "\"
        Else
            Return Path
        End If
    End Function

    Shared Function IsEmptyStr(ByVal Value As Object) As Boolean
        If Value Is DBNull.Value OrElse Value Is Nothing OrElse (CStr(Value).Length = 0) Then
            Return True
        End If
        Return False
    End Function

    Shared Function CompareList(ByVal Arg As Object, ByVal ParamArray Values() As Object) _
        As Integer
        Return Array.IndexOf(Values, Arg) + 1
    End Function

    Shared Function ReplicateString(ByVal Source As String, ByVal Times As Integer) As _
      String
        Dim i As Integer
        Dim sb As New System.Text.StringBuilder(Source.Length * Times)
        For i = 1 To Times
            sb.Append(Source)
        Next
        Return sb.ToString
    End Function

    Shared Function IsInList(ByVal Value As Integer, ByVal ParamArray List() As Integer) As Boolean
        Return (InList(Value, List) <> -1)
    End Function

    Shared Function InList(ByVal Value As Integer, ByVal ParamArray List() As Integer) As Integer
        InList = -1
        For i As Integer = 0 To List.Length - 1
            If List(i) = Value Then
                InList = i
                Exit For
            End If
        Next
    End Function

    Shared Function CFoxToDate(ByVal aDate As String) As DateTime
        Dim dat As DateTime
        dat = Nothing
        If Not String.IsNullOrEmpty(aDate.Replace("/", "").Trim) Then
            dat = New Date(CInt(aDate.Substring(6)), CInt(aDate.Substring(3, 2)), CInt(aDate.Substring(0, 2)))
        End If
        Return dat
    End Function

    Shared Function IsInList(ByVal Value As String, ByVal ParamArray List() As String) As Boolean
        Return (InList(Value, List) <> -1)
    End Function

    Shared Function InList(ByVal Value As String, ByVal ParamArray List() As String) As Integer
        InList = -1
        For i As Integer = 0 To List.Length - 1
            If List(i) = Value Then
                InList = i
                Exit For
            End If
        Next
    End Function

    Public Shared Function NumeroALetras(ByVal NumeroAConvertir As Decimal) As String
        '********Declara variables de tipo cadena************
        Dim palabras, entero, dec, flag, Numero As String

        '********Declara variables de tipo entero***********
        Dim num, x, y As Integer

        flag = "N"
        Numero = NumeroAConvertir.ToString

        '**********Número Negativo***********
        If Mid(Numero, 1, 1) = "-" Then
            Numero = Mid(Numero, 2, Numero.ToString.Length - 1).ToString
            palabras = "menos "
        End If

        '**********Si tiene ceros a la izquierda*************
        For x = 1 To Numero.ToString.Length
            If Mid(Numero, 1, 1) = "0" Then
                Numero = Trim(Mid(Numero, 2, Numero.ToString.Length).ToString)
                If Numero.ToString.Length = 0 Then palabras = ""
            Else
                Exit For
            End If
        Next

        '*********Dividir parte entera y decimal************
        For y = 1 To Len(Numero)
            If Mid(Numero, y, 1) = "." Then
                flag = "S"
            Else
                If flag = "N" Then
                    entero = entero + Mid(Numero, y, 1)
                Else
                    dec = dec + Mid(Numero, y, 1)
                End If
            End If
        Next y

        If Len(dec) = 1 Then dec = dec & "0"

        '**********proceso de conversión***********
        flag = "N"

        If Val(Numero) <= 999999999 Then
            For y = Len(entero) To 1 Step -1
                num = Len(entero) - (y - 1)
                Select Case y
                    Case 3, 6, 9
                        '**********Asigna las palabras para las centenas***********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" And Mid(entero, num + 2, 1) = "0" Then
                                    palabras = palabras & "cien "
                                Else
                                    palabras = palabras & "ciento "
                                End If
                            Case "2"
                                palabras = palabras & "doscientos "
                            Case "3"
                                palabras = palabras & "trescientos "
                            Case "4"
                                palabras = palabras & "cuatrocientos "
                            Case "5"
                                palabras = palabras & "quinientos "
                            Case "6"
                                palabras = palabras & "seiscientos "
                            Case "7"
                                palabras = palabras & "setecientos "
                            Case "8"
                                palabras = palabras & "ochocientos "
                            Case "9"
                                palabras = palabras & "novecientos "
                        End Select
                    Case 2, 5, 8
                        '*********Asigna las palabras para las decenas************
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    flag = "S"
                                    palabras = palabras & "diez "
                                End If
                                If Mid(entero, num + 1, 1) = "1" Then
                                    flag = "S"
                                    palabras = palabras & "once "
                                End If
                                If Mid(entero, num + 1, 1) = "2" Then
                                    flag = "S"
                                    palabras = palabras & "doce "
                                End If
                                If Mid(entero, num + 1, 1) = "3" Then
                                    flag = "S"
                                    palabras = palabras & "trece "
                                End If
                                If Mid(entero, num + 1, 1) = "4" Then
                                    flag = "S"
                                    palabras = palabras & "catorce "
                                End If
                                If Mid(entero, num + 1, 1) = "5" Then
                                    flag = "S"
                                    palabras = palabras & "quince "
                                End If
                                If Mid(entero, num + 1, 1) > "5" Then
                                    flag = "N"
                                    palabras = palabras & "dieci"
                                End If
                            Case "2"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "veinte "
                                    flag = "S"
                                Else
                                    palabras = palabras & "veinti"
                                    flag = "N"
                                End If
                            Case "3"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "treinta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "treinta y "
                                    flag = "N"
                                End If
                            Case "4"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cuarenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cuarenta y "
                                    flag = "N"
                                End If
                            Case "5"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "cincuenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "cincuenta y "
                                    flag = "N"
                                End If
                            Case "6"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "sesenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "sesenta y "
                                    flag = "N"
                                End If
                            Case "7"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "setenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "setenta y "
                                    flag = "N"
                                End If
                            Case "8"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "ochenta "
                                    flag = "S"
                                Else
                                    palabras = palabras & "ochenta y "
                                    flag = "N"
                                End If
                            Case "9"
                                If Mid(entero, num + 1, 1) = "0" Then
                                    palabras = palabras & "noventa "
                                    flag = "S"
                                Else
                                    palabras = palabras & "noventa y "
                                    flag = "N"
                                End If
                        End Select
                    Case 1, 4, 7
                        '*********Asigna las palabras para las unidades*********
                        Select Case Mid(entero, num, 1)
                            Case "1"
                                If flag = "N" Then
                                    If y = 1 Then
                                        palabras = palabras & "uno "
                                    Else
                                        palabras = palabras & "un "
                                    End If
                                End If
                            Case "2"
                                If flag = "N" Then palabras = palabras & "dos "
                            Case "3"
                                If flag = "N" Then palabras = palabras & "tres "
                            Case "4"
                                If flag = "N" Then palabras = palabras & "cuatro "
                            Case "5"
                                If flag = "N" Then palabras = palabras & "cinco "
                            Case "6"
                                If flag = "N" Then palabras = palabras & "seis "
                            Case "7"
                                If flag = "N" Then palabras = palabras & "siete "
                            Case "8"
                                If flag = "N" Then palabras = palabras & "ocho "
                            Case "9"
                                If flag = "N" Then palabras = palabras & "nueve "
                        End Select
                End Select

                '***********Asigna la palabra mil***************
                If y = 4 Then
                    If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or
                    (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And
                    Len(entero) <= 6) Then palabras = palabras & "mil "
                End If

                '**********Asigna la palabra millón*************
                If y = 7 Then
                    If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                        palabras = palabras & "millón "
                    Else
                        palabras = palabras & "millones "
                    End If
                End If
            Next y

            '**********Une la parte entera y la parte decimal*************
            If dec <> "" Then
                NumeroALetras = palabras & "con " & dec
            Else
                NumeroALetras = palabras
            End If
        Else
            NumeroALetras = ""
        End If
    End Function

    Public Shared Function NumToWords(ByVal Amount As Decimal, ByVal Language As enLanguageEnum, ByVal Genere As enGenreEnum, ByVal Format As enFormatEnum) As String
        Dim words As String
        If Math.Truncate(Amount) = Amount Then
            If Language = enLanguageEnum.laCatala Then
                words = IntToWordsCAT(CInt(Amount), Genere, Format)
            Else
                words = IntToWordsESP(CInt(Amount), Genere, Format)
            End If
        Else
            If Language = enLanguageEnum.laCatala Then
                words = IntToWordsCAT(CInt(Math.Truncate(Amount)), Genere, Format) + " euros i " + IntToWordsCAT(CInt((Amount - Math.Truncate(Amount)) * 100), Genere, Format) + " cèntims."
            Else
                words = IntToWordsESP(CInt(Math.Truncate(Amount)), Genere, Format) + " euros y " + IntToWordsESP(CInt((Amount - Math.Truncate(Amount)) * 100), Genere, Format) + " centimos."
            End If
        End If
        Return words
    End Function

    Public Shared Function IntToWordsESP(ByVal Amount As Integer, ByVal Genere As enGenreEnum, ByVal Format As enFormatEnum) As String
        Dim workNumber As Integer
        Dim Centenes As Integer
        Dim workTexte As String
        Dim fiCentenes As String
        Dim runningNumber As Integer

        fiCentenes = IIf(Genere = enGenreEnum.geFemeni, "tas", "tos").ToString

        Dim texteCentenes() As String = {"", "cien", "doscien", "trescien", "cuatrocien", "quinien", "seiscien", "setecien", "ochocien", "novecien"}

        Dim texteDecenes() As String = {"", "", "", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"}
        Dim texteUnitats() As String = {"", "una", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve", "diez", "once", "doce", "trece", "catorce", "quince", "dieciseis", "diecisiete", "dieciocho", "diecinueve", "veinte", "veintiuna", "veintidos", "veintitres", "veinticuatro", "veinticinco", "veintiseis", "veintisiete", "veintiocho", "veintinueve"}

        If Genere = enGenreEnum.geMasculi Then
            texteUnitats(0) = "uno"
            texteUnitats(20) = "veintiuno"
        End If

        runningNumber = Amount
        workTexte = ""
        If runningNumber > 0 Then
            ' millones
            Centenes = 0
            workNumber = runningNumber \ 100000000
            If workNumber > 0 Then
                workTexte = texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber = runningNumber - workNumber * 100000000
            End If
            workNumber = runningNumber \ 1000000
            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "to", "tos").ToString
                End If
                If workNumber > 29 Then
                    workTexte = workTexte + " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 1 Then
                        workTexte = workTexte + " y " + texteUnitats(workNumber Mod 10)
                    Else
                        workTexte = workTexte + " y un"
                    End If
                Else
                    If workNumber = 1 Then
                        workTexte = workTexte + " un"
                    Else
                        workTexte = workTexte + " " + IIf(workNumber = 21, "veintiun", texteUnitats(workNumber)).ToString
                    End If
                End If
                If Centenes = 0 And workNumber = 1 Then
                    workTexte = "un millon"
                Else
                    workTexte = workTexte + " millones"
                End If
            Else
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "", "tos").ToString + " millones"
                End If
            End If
            '
            ' miles
            '
            Centenes = 0
            runningNumber = runningNumber - workNumber * 1000000
            workNumber = (runningNumber \ 100000)
            If workNumber > 0 Then
                workTexte = workTexte + " " + texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber = runningNumber - workNumber * 100000
            End If
            workNumber = (runningNumber \ 1000)
            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "to", fiCentenes).ToString
                End If
                If workNumber > 29 Then
                    workTexte = workTexte + " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 0 Then
                        workTexte = workTexte + " y " + texteUnitats(workNumber Mod 10)
                    End If
                Else
                    If workNumber > 0 Then
                        texteUnitats(1) = ""
                        workTexte = workTexte + " " + texteUnitats(workNumber)
                    End If
                End If
            Else
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "", fiCentenes).ToString
                End If
            End If
            If Centenes > 0 Or workNumber > 0 Then
                workTexte = workTexte + " mil"
            End If
            '
            ' Cents
            '
            Centenes = 0
            runningNumber = runningNumber - workNumber * 1000
            workNumber = (runningNumber \ 100)
            If workNumber > 0 Then
                workTexte = workTexte + " " + texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber = runningNumber - workNumber * 100
            End If
            workNumber = runningNumber

            If Genere = enGenreEnum.geFemeni Then
                texteUnitats(1) = "una"
            Else
                If Format = enFormatEnum.fmOrdinal Then
                    texteUnitats(1) = "uno"

                Else
                    texteUnitats(1) = "un"
                End If
            End If

            ' texteUnitats(1) = IIf(Genere = 0, "una", "uno").ToString

            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "to", fiCentenes).ToString
                End If
                If workNumber > 29 Then
                    workTexte = workTexte + " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 0 Then
                        workTexte = workTexte + " y " + texteUnitats(workNumber Mod 10)
                    End If
                Else
                    workTexte = workTexte + " " + texteUnitats(workNumber)
                End If
            Else
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "", fiCentenes).ToString
                End If
            End If
        End If

        Return workTexte.TrimEnd

    End Function

    Public Shared Function IntToWordsCAT(ByVal Amount As Integer, ByVal Genere As enGenreEnum, ByVal Format As enFormatEnum) As String
        Dim workNumber As Integer
        Dim Centenes As Integer
        Dim workTexte As String
        Dim runningNumber As Integer

        Dim texteCentenes() As String = {"", "cent", "dos-cent", "tres-cent", "quatre-cent", "cinc-cent", "sis-cent", "set-cent", "vuit-cent", "nou-cent"}
        Dim texteDecenes() As String = {"", "", "", "trenta", "quaranta", "cinquanta", "seixanta", "setanta", "vuitanta", "noranta"}
        Dim texteUnitats() As String = {"", "un", "dos", "tres", "quatre", "cinc", "sis", "set", "vuit", "nou", "deu", "onze", "dotze", "tretze", "catorze", "quinze", "setze", "disset", "divuit", "dinou", "vint", "vint-i-un", "vint-i-dos", "vint-i-tres", "vint-i-quatre", "vint-i-cinc", "vint-i-sis", "vint-i-set", "vint-i-vuit", "vint-i-nou"}

        runningNumber = Amount
        workTexte = ""
        If runningNumber > 0 Then
            ' milions
            Centenes = 0
            workNumber = runningNumber \ 100000000
            If workNumber > 0 Then
                workTexte = texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber -= workNumber * 100000000
            End If
            workNumber = runningNumber \ 1000000
            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte += IIf(Centenes = 1, "", "s").ToString
                End If
                If workNumber > 29 Then
                    workTexte += " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 0 Then
                        workTexte += "-" + texteUnitats(workNumber Mod 10)
                    End If
                Else
                    workTexte += texteUnitats(workNumber)
                End If
                If Centenes = 0 And workNumber = 1 Then
                    workTexte = "un milió"
                Else
                    workTexte += " milions"
                End If
            Else
                If Centenes > 0 Then
                    workTexte += IIf(Centenes = 1, "", "s").ToString + " milions"
                End If
            End If
            '
            ' miles
            '
            If Genere = enGenreEnum.geFemeni Then
                texteCentenes(2) = "dues-cent"
                texteUnitats(2) = "dues"
                texteUnitats(21) = "vint-i-una"
                texteUnitats(22) = "vint-i-dues"
            Else
                texteCentenes(2) = "dos-cent"
                texteUnitats(2) = "dos"
                texteUnitats(21) = "vint-i-un"
                texteUnitats(22) = "vint-i-dos"
            End If
            Centenes = 0
            runningNumber -= workNumber * 1000000
            workNumber = (runningNumber \ 100000)
            If workNumber > 0 Then
                workTexte += " " + texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber = runningNumber - workNumber * 100000
            End If
            workNumber = (runningNumber \ 1000)
            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte += IIf(Centenes = 1, "", "s").ToString
                End If
                If workNumber > 29 Then
                    workTexte += " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 0 Then
                        If (workNumber Mod 10) > 1 Then
                            workTexte += "-" + texteUnitats(workNumber Mod 10)
                        Else
                            If Genere = enGenreEnum.geFemeni Then
                                workTexte += " una"
                            Else
                                workTexte += " un"
                            End If
                        End If
                    End If
                Else
                    If workNumber > 1 Then
                        workTexte += " " + texteUnitats(workNumber)
                    Else
                        If Centenes > 0 Then
                            If Genere = enGenreEnum.geFemeni Then
                                workTexte += " una"
                            Else
                                workTexte += " un"
                            End If
                        End If
                    End If
                End If
            Else
                If Centenes > 0 Then
                    workTexte += IIf(Centenes = 1, "", "s").ToString
                End If
            End If
            If Centenes > 0 Or workNumber > 0 Then
                workTexte = workTexte + " mil"
            End If
            '
            ' Cents
            '
            Centenes = 0
            If Genere = enGenreEnum.geFemeni Then
                texteUnitats(1) = "una"
            Else
                If Format = enFormatEnum.fmOrdinal Then
                    texteUnitats(1) = "u"

                Else
                    texteUnitats(1) = "un"
                End If
            End If
            'texteUnitats(1) = IIf(Genere = enGenreEnum.geFemeni, "una", "un").ToString
            runningNumber = runningNumber - workNumber * 1000
            workNumber = (runningNumber \ 100)
            If workNumber > 0 Then
                workTexte += " " + texteCentenes(workNumber)
                Centenes = workNumber
                runningNumber -= workNumber * 100
            End If
            workNumber = runningNumber
            If workNumber > 0 Then
                If Centenes > 0 Then
                    workTexte = workTexte + IIf(Centenes = 1, "", "s").ToString
                End If
                If workNumber > 29 Then
                    workTexte += " " + texteDecenes(workNumber \ 10)
                    If (workNumber Mod 10) > 0 Then
                        workTexte += "-" + texteUnitats(workNumber Mod 10)
                    End If
                Else
                    workTexte = workTexte + " " + texteUnitats(workNumber)
                End If
            Else
                If Centenes > 0 Then
                    workTexte += IIf(Centenes = 1, "", "s").ToString
                End If
            End If
        End If

        Return workTexte.TrimEnd

    End Function



    Public Shared Function IndexOfOcurrence(ByVal Text As String, ByVal Lookup As String, ByVal Occurrence As Integer) As Integer
        Dim index As Integer
        Dim txt As Char() = Text.ToCharArray
        If Occurrence < 1 Or Occurrence > txt.Length Then
            Return -1
        End If
        For index = 0 To txt.Length
            If txt(index) = Lookup Then
                Occurrence -= 1
            End If
            If Occurrence = 0 Then
                Exit For
            End If
        Next
        Return index
    End Function

    Public Shared Function IsNullValue(ByVal Value As Object) As Boolean
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return True
        End If
        If Value.GetType Is System.Type.GetType("System.Guid") Then
            If CType(Value, Guid) = Guid.Empty Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Shared Function IsNullOrEmptyValue(ByVal Value As Object) As Boolean
        If Value Is DBNull.Value OrElse Value Is Nothing Then
            Return True
        End If

        If Value.GetType Is System.Type.GetType("System.Guid") Then
            If CType(Value, Guid) = Guid.Empty Then
                Return True
            End If
        Else
            If Value.GetType Is System.Type.GetType("System.String") Then
                If String.IsNullOrEmpty(Value.ToString.Trim) Then
                    Return True
                End If
            End If
            If Value.GetType Is System.Type.GetType("System.Decimal") Then
                If CDec(Value) = 0D Then
                    Return True
                End If
            End If
            If Value.GetType Is System.Type.GetType("System.Int32") Then
                If CInt(Value) = 0 Then
                    Return True
                End If
            End If
            If Value.GetType Is System.Type.GetType("System.DateTime") Then
                If CDate(Value) = CDate(Nothing) Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Public Shared Function InputBox(ByVal Title As String, ByVal Prompt As String, ByVal DefaultResult As String) As String
        Dim ib As New csUtils.csInputBox(Title, Prompt, DefaultResult)
        ib.ShowDialog()
        DefaultResult = ib.Result
        ib.Dispose()
        Return DefaultResult
    End Function

    Public Shared Function StringToDate(ByVal Data As String, ByVal Format As String) As Date
        Return DateTime.ParseExact(Data, Format, New System.Globalization.DateTimeFormatInfo)
    End Function

    Public Shared Function CVal(ByVal Value As String) As Integer
        If IsNothing(Value) OrElse IsDBNull(Value) Then
            Return 0
        Else
            Dim ValidValue As String = String.Empty
            Dim pos As Integer = 0

            Value = Value.Trim
            While pos < Value.Length
                If Char.IsDigit(Value.Chars(pos)) Then
                    ValidValue += Value.Chars(pos)
                Else
                    Exit While
                End If
                pos += 1
            End While
            If ValidValue = String.Empty Then
                ValidValue = "0"
            End If
            Return CInt(ValidValue)
        End If
    End Function

#End Region

#Region " Funcions BANCA "
    Shared Function csbIsCompteOK(ByVal Compte As String) As Boolean
        Dim DigitsControl As String
        If Compte.Length <> 20 Then
            Return False
        End If
        If Not IsDigit(Compte) Then
            Return False
        End If
        DigitsControl = csbGetControlDigits(Compte.Substring(0, 8), Compte.Substring(10, 10))
        Return (DigitsControl = Compte.Substring(8, 2))
    End Function

    Shared Function csbIsCompteOK(ByVal Compte As String, ByRef DigitsControl As String) As Boolean
        If Compte.Length <> 20 Then
            Return False
        End If
        If Not IsDigit(Compte) Then
            Return False
        End If
        DigitsControl = csbGetControlDigits(Compte.Substring(0, 8), Compte.Substring(10, 10))
        Return (DigitsControl = Compte.Substring(8, 2))
    End Function

    Shared Function csbGetControlDigits(ByVal EntitatOficina As String, ByVal Compte As String) As String
        Return (csbGetControlDigit(EntitatOficina) & csbGetControlDigit(Compte))
    End Function

    Shared Function csbGetControlDigits(ByVal Compte As String) As String
        Return csbGetControlDigits(Compte.Substring(0, 8), Compte.Substring(10, 10))
    End Function

    Shared Function csbGetControlDigits(ByVal Entitat As String, ByVal Oficina As String, ByVal Compte As String) As String
        Return (csbGetControlDigit(Entitat & Oficina) & csbGetControlDigit(Compte))
    End Function

    Shared Function csbGetControlDigit(ByVal Value As String) As String
        Dim Pesos() As Integer = {6, 3, 7, 9, 10, 5, 8, 4, 2, 1}
        Dim Suma As Integer = 0
        Dim Reste As Integer = 0
        Dim i As Integer

        For i = 0 To Value.Length - 1
            Suma += CInt(Value.Substring(Value.Length - i - 1, 1)) * Pesos(i)
        Next
        Reste = 11 - (Suma Mod 11)
        If Reste > 9 Then
            Reste = 11 - Reste
        End If
        Return Reste.ToString
    End Function

    Shared Function GetClauRIB(ByVal RIB As String) As String

        Dim Banc As String
        Dim Oficina As String
        Dim Compte As String
        Dim Control As String
        Dim CompteDgt As String

        If RIB.Length < 21 Then
            RIB = RIB.PadLeft(21, "0"c)
        End If
        RIB = RIB.Substring(0, 21)

        Banc = RIB.Substring(0, 5)
        Oficina = RIB.Substring(5, 5)
        Compte = RIB.Substring(10, 11)
        CompteDgt = ""

        Dim tbFrom As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
        Dim tbTo As String = "123456789123456789234567890123456789"

        For Each c As Char In Compte.ToUpper
            CompteDgt += tbTo.Substring(tbFrom.IndexOf(c), 1)
        Next

        Dim ribn As String = Banc + Oficina + CompteDgt + "00"

        Dim ctrl As Integer

        Dim r As Integer
        For Each c As Char In ribn
            r = CInt(r.ToString + c) Mod 97
        Next

        ctrl = 97 - r

        Control = ctrl.ToString("D2")

        Return Control

    End Function

    Shared Function IsValidRIB(ByVal RIB As String) As Boolean
        If RIB.Length <> 23 Then
            Return False
        End If
        Return (GetClauRIB(RIB) = RIB.Substring(21, 2))
    End Function

    Shared Function IsValidRIB(ByVal RIB As String, ByRef Digits As String) As Boolean
        Digits = ""
        If RIB.Length <> 23 Then
            Return False
        End If
        Digits = GetClauRIB(RIB)
        Return (Digits = RIB.Substring(21, 2))
    End Function

    Shared Function GetClauIBAN(ByVal Pais As String, ByVal Compte As String) As String
        Dim IBAN As String
        Dim IBANdgt As String
        Dim tbFrom As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

        If Pais.Length <> 2 Then
            Return ""
        End If
        ' Pasem el pais i el compte al format de calcul: Compte + país + '00'
        Pais = Pais.ToUpper
        IBAN = Compte + Pais + "00"

        ' pasem els chars alfanumeric a equivalencia numerica
        IBANdgt = ""
        For Each c As Char In IBAN
            IBANdgt += tbFrom.IndexOf(c).ToString
        Next

        'Calculem el modulus 97
        Dim r As Integer
        For Each c As Char In IBANdgt
            r = CInt(r.ToString + c) Mod 97
        Next

        'Restem el valor obtingut de 97 + 1

        r = 98 - r

        Return r.ToString("D2")

    End Function

    Shared Function GetIBAN(ByVal Pais As String, ByVal Compte As String) As String
        Return Pais + GetClauIBAN(Pais, Compte) + Compte
    End Function

    Shared Function IsValidIBAN(ByVal IBAN As String) As Boolean
        If IBAN.Length < 5 Then
            Return False
        End If
        Return (IBAN.Substring(2, 2) = GetClauIBAN(IBAN.Substring(0, 2), IBAN.Substring(4)))
    End Function

    Shared Function IsValidIBAN(ByVal IBAN As String, ByRef ControlDigits As String) As Boolean
        If IBAN.Length < 5 Then
            ControlDigits = ""
            Return False
        End If
        ControlDigits = GetClauIBAN(IBAN.Substring(0, 2), IBAN.Substring(4))
        Return (IBAN.Substring(2, 2) = ControlDigits)
    End Function

#End Region

#Region " Funcions FITXERS "
    Shared Function File2ArrayList(ByVal FilePath As String, ByVal AL As ArrayList) As Boolean
        Dim sr As System.IO.StreamReader
        Try
            sr = New System.IO.StreamReader(FilePath)
            Do Until sr.Peek = -1
                AL.Add(sr.ReadLine)
            Loop
        Catch ex As Exception
            Return False
        Finally
            If Not sr Is Nothing Then sr.Close()
        End Try
        Return True
    End Function

    Shared Function ArrayList2File(ByVal FilePath As String, ByVal AL As ArrayList, Optional ByVal Append As Boolean = False) As Boolean
        Dim sw As System.IO.StreamWriter
        Try
            sw = New System.IO.StreamWriter(FilePath, Append)
            For Each st As String In AL
                sw.WriteLine(st)
            Next
        Catch
            Return False
        Finally
            If Not sw Is Nothing Then sw.Close()
        End Try
        Return True
    End Function

    Shared Function SaveTextFile(ByVal filePath As String, ByVal fileContent As String, Optional ByVal append As Boolean = False) As Boolean
        Dim sw As System.IO.StreamWriter
        Try
            sw = New System.IO.StreamWriter(filePath, append)
            sw.Write(fileContent)
            Return True
        Catch e As Exception
            Return False
        Finally
            If Not sw Is Nothing Then sw.Close()
        End Try
    End Function

    Shared Function LoadTextFile(ByVal filePath As String) As String
        Dim sr As System.IO.StreamReader
        Try
            sr = New System.IO.StreamReader(filePath)
            LoadTextFile = sr.ReadToEnd()
        Finally
            If Not sr Is Nothing Then sr.Close()
        End Try
    End Function

    Shared Function LoadTextFile(ByVal filePath As String, ByVal CodePage As Integer) As String
        Dim sr As System.IO.StreamReader
        Try
            sr = New System.IO.StreamReader(filePath, System.Text.Encoding.GetEncoding(CodePage))
            LoadTextFile = sr.ReadToEnd()
        Finally
            If Not sr Is Nothing Then sr.Close()
        End Try
    End Function

    Shared Function GetApplicationPath() As String
        Return System.IO.Path.GetDirectoryName _
            (System.Reflection.Assembly.GetExecutingAssembly().Location())
    End Function

    Shared Function DebugLog(ByVal LogText As String) As Boolean
        Dim fileLog As String
        fileLog = System.IO.Path.Combine(GetApplicationPath, "DebugInfo.log")
        SaveTextFile(fileLog, String.Format("{0:dd/MM/yy HH:mm:ss}: {1}{2}", Date.Now, LogText, vbCrLf), True)
        'MsgBox(fileLog)
        Return True
    End Function

    Shared Function DebugLog(ByVal WriteLog As Boolean, ByVal LogText As String) As Boolean
        If Not WriteLog Then
            Return False
        End If
        DebugLog(LogText)
        Return True
    End Function

#End Region

#Region " Data "

    Shared Function FmtDataIdioma(ByVal Data As Date, ByVal IdiomaID As enIdiomaEnum) As String
        Dim CultureInfo As System.Globalization.CultureInfo
        Dim DataReturn As String
        Dim Culture As String = ""

        Select Case IdiomaID
            Case enIdiomaEnum.idCastella
                Culture = "es-ES"
            Case enIdiomaEnum.idAngles
                Culture = "en-US" 'poso "en-US" perquè "en-GB" no retorna el dia de la setmana traduit'
            Case enIdiomaEnum.idFrances
                Culture = "fr-FR"
            Case enIdiomaEnum.idAlemany
                Culture = "de-DE"
            Case enIdiomaEnum.idItalia
                Culture = "it-IT"
            Case Else
                IdiomaID = enIdiomaEnum.idCatala
        End Select

        If IdiomaID = enIdiomaEnum.idCatala Then
            DataReturn = FmtDataCatala(Data)
        Else
            CultureInfo = New System.Globalization.CultureInfo(Culture, True)
            DataReturn = Date.Now.ToString("D", CultureInfo)
        End If

        Return DataReturn

    End Function

    Shared Function FmtDataCatala(ByVal Data As Date) As String
        Return String.Format("{0} {1} de {2}", Microsoft.VisualBasic.DateAndTime.Day(Data), GetMesCatala(Data, True), Year(Data))
    End Function

    Shared Function GetMesCatala(ByVal NumeroMes As Integer, ByVal Prefix As Boolean) As String
        Dim Mesos() As String = {"Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre"}
        Dim Prefixes() As String = {"de ", "de ", "de ", "d'", "de ", "de ", "de ", "d'", "de ", "d'", "de ", "de "}
        Dim mes As String
        mes = Mesos(NumeroMes - 1)
        If Prefix Then mes = Prefixes(NumeroMes - 1) + mes
        Return mes
    End Function

    Shared Function Mesos() As String()
        Mesos = New String() {"Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre"}
    End Function

    Shared Function GetMesCatala(ByVal Data As Date, ByVal Prefix As Boolean) As String
        Return GetMesCatala(Month(Data), Prefix)
    End Function

    Shared Function FmtDataFrances(ByVal Data As Date) As String
        Return String.Format("{0} {1} {2}", Microsoft.VisualBasic.DateAndTime.Day(Data), GetMesFrances(Data), Year(Data))
    End Function

    Shared Function GetMesFrances(ByVal Data As Date) As String
        Dim Mesos() As String = {"Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Aout", "Septembre", "Octobre", "Novembre", "Décembre"}
        Return Mesos(Month(Data) - 1)
    End Function


    Shared Function SemanaSanta(ByVal year As Integer) As DateTime
        Dim g, c, h, i, j, l, month, day As Integer

        g = year Mod 19
        c = year \ 100
        h = ((c - (c \ 4) - ((8 * c + 13) \ 25) + (19 * g) + 15) Mod 30)
        i = h - ((h \ 28) * (1 - (h \ 28) * (29 \ (h + 1)) * ((21 - g) \ 11)))
        j = ((year + (year \ 4) + i + 2 - c + (c \ 4)) Mod 7)
        l = i - j

        month = 3 + ((l + 40) \ 44)
        day = l + 28 - (31 * (month \ 4))

        Return New DateTime(year, month, day)
    End Function

    Shared Function Age(ByVal birthDate As Date, Optional ByVal currentDate As Date = #1/1/1900#, Optional ByVal exactAge As Boolean = True) As Integer
        If currentDate = #1/1/1900# Then currentDate = Date.Today
        Age = currentDate.Year - birthDate.Year

        If exactAge Then
            ' subtract one if this year's birthday hasn't occurred yet
            If New Date(currentDate.Year, birthDate.Month,
                birthDate.Day) > currentDate Then
                Age -= 1
            End If
        End If
    End Function

    Shared Function GetPeriodeCatala(ByVal Data As Date) As String
        Return GetMesCatala(Data.Month, False) + " " + Data.Year.ToString
    End Function

    Shared Function GetPeriodeCatala(ByVal Data As String) As String
        Return GetMesCatala(CInt(Data.Substring(4, 2)), False) + " " + Data.Substring(0, 4)
    End Function

    Shared Sub GetFirstLastMonth(ByVal Data As Date, ByRef DataInicial As Date, ByRef DataFinal As Date)
        DataInicial = New Date(Data.Year, Data.Month, 1)
        DataFinal = DataInicial.AddMonths(1).AddDays(-1)
    End Sub

    Shared Function DateToYMD(ByVal Data As Date) As Integer
        Return Data.Year * 10000 + Data.Month * 100 + Data.Day
    End Function

    Public Class DateAndTime

        Public Shared Function GetYearWeek(ByVal inDate As Date) As Integer
            Const JAN As Integer = 1
            Const DEC As Integer = 12
            Const LASTDAYOFDEC As Integer = 31
            Const FIRSTDAYOFJAN As Integer = 1
            Const THURSDAY As Integer = 4
            Dim ThursdayFlag As Boolean = False
            Dim YearNumber As Integer

            ' Get the day number since the beginning of the year
            Dim DayOfYear As Integer = inDate.DayOfYear
            YearNumber = inDate.Year

            ' Get the numeric weekday of the first day of the
            ' year (using sunday as FirstDay)
            Dim StartWeekDayOfYear As Integer =
               DirectCast(New DateTime(inDate.Year, JAN, FIRSTDAYOFJAN).DayOfWeek, Integer)
            Dim EndWeekDayOfYear As Integer =
                DirectCast(New DateTime(inDate.Year, DEC, LASTDAYOFDEC).DayOfWeek, Integer)

            ' Compensate for the fact that we are using monday
            ' as the first day of the week
            If StartWeekDayOfYear = 0 Then
                StartWeekDayOfYear = 7
            End If
            If EndWeekDayOfYear = 0 Then
                EndWeekDayOfYear = 7
            End If

            ' Calculate the number of days in the first and last week
            Dim DaysInFirstWeek As Integer = 8 - StartWeekDayOfYear
            Dim DaysInLastWeek As Integer = 8 - EndWeekDayOfYear

            ' If the year either starts or ends on a thursday it will have a 53rd week
            If StartWeekDayOfYear = THURSDAY OrElse EndWeekDayOfYear = THURSDAY Then
                ThursdayFlag = True
            End If

            ' We begin by calculating the number of FULL weeks between the start of the year and
            ' our date. The number is rounded up, so the smallest possible value is 0.
            Dim FullWeeks As Integer =
                CType(Math.Ceiling((DayOfYear - DaysInFirstWeek) / 7), Integer)

            Dim WeekNumber As Integer = FullWeeks

            ' If the first week of the year has at least four days, then the actual week number for our date
            ' can be incremented by one.
            If DaysInFirstWeek >= THURSDAY Then
                WeekNumber = WeekNumber + 1
            End If

            ' If week number is larger than week 52 (and the year doesn't either start or end on a thursday)
            ' then the correct week number is 1.
            If WeekNumber > 52 AndAlso Not ThursdayFlag Then
                WeekNumber = 1
                YearNumber += 1
            End If

            'If week number is still 0, it means that we are trying to evaluate the week number for a
            'week that belongs in the previous year (since that week has 3 days or less in our date's year).
            'We therefore make a recursive call using the last day of the previous year.
            If WeekNumber = 0 Then
                WeekNumber = GetYearWeek(New DateTime(inDate.Year - 1, DEC, LASTDAYOFDEC)) Mod 100
                YearNumber -= 1
            End If
            Return YearNumber * 100 + WeekNumber
        End Function

        Public Shared Function GetFirstDayOfWeek(ByVal InDate As Date) As Date
            Return GetFirstDayOfWeek(GetYearWeek(InDate))
        End Function

        Public Shared Function GetFirstDayOfWeek(ByVal YearWeek As Integer) As Date
            Return GetFirstDayOfWeek(YearWeek \ 100, YearWeek Mod 100)
        End Function

        Public Shared Function GetFirstDayOfWeek(ByVal year As Integer, ByVal week As Integer) As Date
            Return GetFirstDayOfFirstWeekOfYear(New Date(year, 1, 1)).AddDays((week - 1) * 7)
        End Function

        Public Shared Function GetLastDayOfWeek(ByVal Data As Date) As Date
            Return GetLastDayOfWeek(GetYearWeek(Data))
        End Function

        Public Shared Function GetLastDayOfWeek(ByVal yearWeek As Integer) As Date
            Return GetLastDayOfWeek(yearWeek \ 100, yearWeek Mod 100)
        End Function

        Public Shared Function GetLastDayOfWeek(ByVal year As Integer, ByVal week As Integer) As Date
            Return GetFirstDayOfWeek(year, week).AddDays(6)
        End Function

        Public Shared Function AddWeeks(ByVal YearWeek As Integer, ByVal Weeks As Integer) As Integer
            Return GetYearWeek(GetFirstDayOfWeek(YearWeek \ 100, YearWeek Mod 100).AddDays(Weeks * 7))
        End Function

        Public Shared Function GetFirstDayOfFirstWeekOfYear(ByVal inDate As Date) As Date

            Dim StartDate As Date = New DateTime(inDate.Year, 1, 1)
            Dim StartWeekDayOfYear As Integer = DirectCast(StartDate.DayOfWeek, Integer)
            If StartWeekDayOfYear = 0 Then
                StartWeekDayOfYear = 7
            End If
            If StartWeekDayOfYear > 4 Then
                StartDate = StartDate.AddDays(8 - StartWeekDayOfYear)
            Else
                StartDate = StartDate.AddDays(1 - StartWeekDayOfYear)
            End If

            Return StartDate

        End Function

        'Public Shared Function GetFirstWeekDayOfYear(ByVal year As Integer) As Date
        '  Return GetFirstWeekDayOfMonth(year, 1)
        'End Function

        'Public Shared Function GetFirstWeekDayOfMonth(ByVal year As Integer, ByVal month As Integer) As Date
        '  Dim d As Date = New Date(year, month, 1)
        '  Dim offset As Integer
        '  Select Case d.DayOfWeek
        '    Case DayOfWeek.Monday
        '      offset = 0
        '    Case DayOfWeek.Tuesday
        '      offset = 6
        '    Case DayOfWeek.Wednesday
        '      offset = 5
        '    Case DayOfWeek.Thursday
        '      offset = 4
        '    Case DayOfWeek.Friday
        '      offset = 3
        '    Case DayOfWeek.Saturday
        '      offset = 2
        '    Case DayOfWeek.Sunday
        '      offset = 1
        '  End Select
        '  Return d.AddDays(offset)
        'End Function

        'Public Shared Function GetWeeksInMonth(ByVal Data As Date) As Integer
        '  Return GetWeeksInMonth(Data.Year, Data.Month)
        'End Function

        'Public Shared Function GetWeeksInMonths(ByVal year As Integer, ByVal month As Integer) As Integer
        '  Dim d As Date = GetFirstWeekDayOfMonth(year, month)
        '  Dim weeks As Integer = 1
        '  While d.AddDays(7).Month = month
        '    weeks += 1
        '  End While
        '  Return weeks
        'End Function

        Public Shared Function GetSerialDate(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer) As Date
            'Return Date.ParseExact(String.Format("{0:D4}{1:D2}{2:D2}", year, month, day), "yyyyMMdd", System.Globalization.DateTimeFormatInfo.InvariantInfo)
            Return New Date(year, month, day)
        End Function

        Public Shared Function FmtMinToDHM(ByVal Minuts As Integer) As String
            Dim dhm As String
            Dim dies, hores As Integer
            dies = Minuts \ (24 * 60)
            hores = (Minuts - (dies * 24 * 60)) \ 60
            Minuts = Minuts - (dies * 24 * 60) - hores * 60
            dhm = String.Format("{0:D2}:{1:D2}:{2:D2}", dies, hores, Minuts)
            Return dhm
        End Function

        Public Shared Function GetLastDayOfMonth(ByVal Data As DateTime) As DateTime
            Return Data.AddDays(1 - Data.Day).AddMonths(1).AddDays(-1)
        End Function

        Public Shared Function GetFirstDayOfMonth(ByVal Data As DateTime) As DateTime
            Return Data.AddDays(1 - Data.Day)
        End Function

    End Class

#End Region

#Region " Internet "
    Shared Function IsValidEmail(ByVal Value As String, Optional ByVal MaxLength As _
      Integer = 255, Optional ByVal IsRequired As Boolean = True) As Boolean
        If Value Is Nothing OrElse Value.Length = 0 Then
            ' rule out the null string case
            Return Not IsRequired
        ElseIf Value.Length > MaxLength Then
            ' rule out values that are longer than allowed
            Return False
        End If

        ' search invalid chars
        If Not System.Text.RegularExpressions.Regex.IsMatch(Value,
            "^[-A-Za-z0-9_@.]+$") Then Return False

        ' search the @ char
        Dim i As Integer = Value.IndexOf("@"c)
        ' there must be at least three chars after the @
        If i <= 0 Or i >= Value.Length - 3 Then Return False
        ' ensure there is only one @ char
        If Value.IndexOf("@"c, i + 1) >= 0 Then Return False

        ' check that the domain portion contains at least one dot
        Dim j As Integer = Value.LastIndexOf("."c)
        ' it can't be before or immediately after the @ char
        If j < 0 Or j <= i + 1 Then Return False

        ' if we get here the address if validated
        Return True
    End Function

    Shared Function GetHostIPAddress(ByVal mStrHost As String) As String
        Dim mIpHostEntry As IPHostEntry = GetHostEntry(mStrHost)
        Dim mIpAddLst As IPAddress() = mIpHostEntry.AddressList()
        ' para efecto de este ejemplo y reducir codigo
        ' se devolvera la primera direccion IP y no se
        ' incluira manejo de excepciones
        Return mIpAddLst(0).ToString
    End Function

    Shared Function MailNotification(ByVal SMTPserver As String, ByVal SMTPuser As String, ByVal SMTPpassword As String, ByVal MailFrom As String, ByVal MailTo As String, ByVal Subject As String, ByVal Body As String) As Boolean
        Dim eMail As System.Net.Mail.MailMessage
        Dim objSmtp As Net.Mail.SmtpClient
        Dim MailSentOK As Boolean

        eMail = New System.Net.Mail.MailMessage

        Try

            eMail.From = New Net.Mail.MailAddress(MailFrom)
            eMail.To.Add(MailTo)


            eMail.Subject = Subject
            eMail.Body = Body
            eMail.IsBodyHtml = False

            objSmtp = New Net.Mail.SmtpClient(SMTPserver)

            objSmtp.Credentials = New System.Net.NetworkCredential(SMTPuser, SMTPpassword)
            objSmtp.Send(eMail)

            eMail.Dispose()

            MailSentOK = True

        Catch ex As Exception

            MailSentOK = False

        End Try

        eMail = Nothing
        objSmtp = Nothing

        Return MailSentOK

    End Function

    Shared Function SimpleMail(ByVal SMTPserver As String, ByVal SMTPuser As String, ByVal SMTPpassword As String, ByVal MailFrom As String, ByVal MailTo As String, ByVal Subject As String, ByVal Body As String, ByVal Attachment As String) As Boolean
        Dim eMail As System.Net.Mail.MailMessage
        Dim objSmtp As Net.Mail.SmtpClient
        Dim MailSentOK As Boolean

        eMail = New System.Net.Mail.MailMessage

        Try

            eMail.From = New Net.Mail.MailAddress(MailFrom)
            eMail.To.Add(MailTo)


            eMail.Subject = Subject
            eMail.Body = Body
            eMail.IsBodyHtml = False

            If Not String.IsNullOrEmpty(Attachment) Then
                eMail.Attachments.Add(New Mail.Attachment(Attachment))
            End If

            objSmtp = New Net.Mail.SmtpClient(SMTPserver)

            objSmtp.Credentials = New System.Net.NetworkCredential(SMTPuser, SMTPpassword)
            objSmtp.Send(eMail)

            eMail.Dispose()

            MailSentOK = True

        Catch ex As Exception

            MailSentOK = False

        End Try

        eMail = Nothing
        objSmtp = Nothing

        Return MailSentOK

    End Function

#End Region

#Region " C1 ComponentOne "

    Shared Function isKeyValid(ByVal e As KeyEventArgs) As Boolean
        If (e.KeyCode < Keys.D0 Or e.KeyCode > Keys.D9) And e.KeyCode <> Keys.Back And e.KeyCode <> Keys.Tab _
             And e.KeyCode <> Keys.Delete And e.KeyCode <> Keys.Left And e.KeyCode <> Keys.Up And e.KeyCode <> Keys.Down And e.KeyCode <> Keys.Right Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

#Region " BAR CODES "

#Region " EAN13 "
    Shared Function EAN13_IsValid(ByVal Value As String) As Boolean
        If Value.Length <> 13 Then
            Return False
        End If
        For Each c As Char In Value
            If Not Char.IsDigit(c) Then
                Return False
            End If
        Next
        If Value.Substring(12, 1) <> EAN13_GetDC(Value) Then
            Return False
        End If
        Return True
    End Function

    Shared Function EAN13_GetDC(ByVal Value As String) As String
        Dim Suma As Integer = 0
        Dim Pes As Integer
        Dim DigitControl As Integer

        If Value.Length < 12 Then
            Value = Value.PadLeft(12, "0"c)
        End If

        Pes = 1
        For i As Integer = 0 To 11
            'Pes = (i Mod 2) * 2 + 1
            Suma = Suma + CInt(Value.Substring(i, 1)) * Pes
            ' Multiplica = IIf(Multiplica = 1, 3, 1)
            Pes = 4 - Pes
        Next
        DigitControl = (Suma Mod 10)
        If DigitControl > 0 Then DigitControl = 10 - DigitControl
        Return DigitControl.ToString
    End Function

    Shared Function EAN13_GetEAN13(ByVal Value As String) As String
        If Value.Length >= 13 Then
            Value = Value.Substring(0, 12)
        Else
            Value = Value.PadLeft(12, "0"c)
        End If
        Value = Value + EAN13_GetDC(Value)
        Return Value
    End Function

#End Region

#Region " DUN14 "

    Shared Function DUN14_GetDC(ByVal Value As String) As String
        Dim Suma As Integer = 0
        Dim Pes As Integer
        Dim DigitControl As Integer

        If Value.Length < 13 Then
            Value = Value.PadLeft(13, "0"c)
        End If

        Pes = 1
        For i As Integer = 0 To 12
            'Pes = (i Mod 2) * 2 + 1
            Suma = Suma + CInt(Value.Substring(i, 1)) * Pes
            ' Multiplica = IIf(Multiplica = 1, 3, 1)
            Pes = 4 - Pes
        Next
        DigitControl = (Suma Mod 10)
        If DigitControl > 0 Then DigitControl = 10 - DigitControl
        Return DigitControl.ToString
    End Function

    Shared Function DUN14_GetDUN14(ByVal VariableLogistica As String, ByVal EAN13 As String) As String
        Dim DUN14 As String
        If EAN13.Length >= 12 Then
            EAN13 = EAN13.Substring(0, 12)
        Else
            EAN13 = EAN13.PadLeft(12, "0"c)
        End If
        DUN14 = VariableLogistica + EAN13
        DUN14 += DUN14_GetDC(DUN14)
        Return DUN14
    End Function

    Shared Function DUN14_IsValid(ByVal Value As String) As Boolean
        If Value.Length <> 14 Then
            Return False
        End If
        For Each c As Char In Value
            If Not Char.IsDigit(c) Then
                Return False
            End If
        Next
        If Value.Substring(13, 1) <> DUN14_GetDC(Value) Then
            Return False
        End If
        Return True
    End Function

#End Region

    Shared Function SSCC_GetDC(ByVal Data As String) As String
        Dim sum As Integer
        Dim digit As Integer

        sum = 0
        For i As Integer = 0 To Data.Length - 1
            sum += CInt(Data(Data.Length - 1 - i).ToString) * ((1 - (i Mod 2)) * 2 + 1)
        Next
        digit = CInt(Math.Ceiling(sum / 10)) * 10 - sum

        Return digit.ToString

    End Function

    Shared Function SSCC_GetSSCC(ByVal SSCCData As String) As String
        Dim sscc As String

        sscc = SSCCData + SSCC_GetDC(SSCCData)


        Return sscc
    End Function

    Shared Function SSCC_IsValid(ByVal Value As String) As Boolean
        For Each c As Char In Value
            If Not Char.IsDigit(c) Then
                Return False
            End If
        Next
        If Value.Substring(Value.Length - 1, 1) <> SSCC_GetDC(Value) Then
            Return False
        End If
        Return True
    End Function

#End Region

#Region " WIN32 "
    <System.Runtime.InteropServices.DllImport("user32")> Shared Function _
  GetSystemMenu(ByVal hWnd As IntPtr, ByVal bRevert As Boolean) As Integer
    End Function

    <System.Runtime.InteropServices.DllImport("user32")> Shared Function _
  EnableMenuItem(ByVal hMenu As Integer, ByVal uIDEnableItem As Integer, ByVal uEnable As Integer) As Boolean
    End Function

    Private Const SC_CLOSE As Integer = &HF060
    Private Const MF_BYCOMMAND As Integer = &H0
    Private Const MF_GRAYED As Integer = &H1
    Private Const MF_ENABLED As Integer = &H0

    Public Shared Sub DisableCloseButton(ByVal Form As Windows.Forms.Form)

        Try
            EnableMenuItem(GetSystemMenu(Form.Handle, False), SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED)

        Catch ex As Exception
        End Try
    End Sub

    Public Shared Function PrevInstance() As Boolean
        If (UBound(Diagnostics.Process.GetProcessesByName(
        Diagnostics.Process.GetCurrentProcess.ProcessName)) > 0) Then
            PrevInstance = True
        End If
        PrevInstance = False
    End Function

    Public Class Win32
        ' Declare Windows' API functions
        Declare Auto Function RegisterWindowMessage Lib "user32.dll" _
                (ByVal lpString As String) As Integer

        Declare Auto Function FindWindow Lib "user32.dll" _
                (ByVal lpClassName As String,
                ByVal lpWindowName As String) As Integer

        Declare Auto Function SendMessage Lib "user32.dll" _
                (ByVal hwnd As Integer, ByVal wMsg As Integer,
                ByVal wParam As Integer, ByVal lParam As Integer) As Integer

        Declare Auto Function IsWindow Lib "user32.dll" _
                (ByVal hwnd As Integer) As Boolean

    End Class

    Public Class ClickYes
        Private wnd As Integer
        Private uClickYes As Integer
        Private Res As Integer

        Public Sub New()
            uClickYes = Win32.RegisterWindowMessage("CLICKYES_SUSPEND_RESUME")
            wnd = Win32.FindWindow("EXCLICKYES_WND", "Express ClickYes 1.2") 'System.String.Empty
        End Sub

        Public Sub Enable()
            ' Send the message to Enable ClickYes
            If Win32.IsWindow(wnd) Then
                Res = Win32.SendMessage(wnd, uClickYes, 1, 0)
            End If
        End Sub

        Public Sub Disable()
            ' Send the message to Suspend ClickYes
            If Win32.IsWindow(wnd) Then
                Res = Win32.SendMessage(wnd, uClickYes, 0, 0)
            End If
        End Sub

    End Class

    Declare Function GetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Integer) As Integer
    Public Shared Function RunningOnTabletPC() As Boolean
        Const SM_TABLETPC As Integer = 86
        If GetSystemMetrics(SM_TABLETPC) <> 0 Then
            Return True
        End If
        Return False
    End Function

#End Region

#Region " RS-232 "
    Shared Sub SetSerialOptions(ByVal CommPortCtrl As System.IO.Ports.SerialPort, ByVal CustomSerialSettings As String)
        Dim DefaultSerialSettings As String = "Port=COM1;Velocitat=9600;Paritat=Ninguna;BitsDades=8;BitsStop=1;DTR=True;RTS=True;DsrDtr=False;XonXoff=False;CtsRts=True;Rs485=False"
        Dim Paritats() As String = {"Ninguna", "Impar", "Par", "Marca", "Espai"}
        Dim BitsParada() As String = {"0", "1", "1.5", "2"}

        If String.IsNullOrEmpty(CustomSerialSettings) Then
            CustomSerialSettings = DefaultSerialSettings
        End If
        Dim items() As String
        items = Split(CustomSerialSettings, ";")
        If items.Length <> 11 Then
            items = Split(DefaultSerialSettings, ";")
        End If

        With CommPortCtrl
            .PortName = CStr(Split(items(0), "=")(1))
            .BaudRate = CInt(Split(items(1), "=")(1))
            .Parity = CType(Array.IndexOf(Paritats, CStr(Split(items(2), "=")(1))), System.IO.Ports.Parity)
            .DataBits = CInt(Split(items(3), "=")(1))
            .StopBits = CType(Array.IndexOf(BitsParada, CStr(Split(items(4), "=")(1))), System.IO.Ports.StopBits)

            .DtrEnable = CBool(Split(items(5), "=")(1))
            .RtsEnable = CBool(Split(items(6), "=")(1))
            If CBool(Split(items(7), "=")(1)) Then
                .Handshake = IO.Ports.Handshake.None
            End If
            If CBool(Split(items(8), "=")(1)) Then
                .Handshake = IO.Ports.Handshake.XOnXOff
            End If
            If CBool(Split(items(9), "=")(1)) Then
                .Handshake = IO.Ports.Handshake.RequestToSend
            End If
        End With
        Return
    End Sub

#End Region

#Region " Registry "

    Public Shared Function GetRegString(ByVal RegKey As String, ByVal KeyName As String) As String
        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegKey)
        GetRegString = Convert.ToString(key.GetValue(KeyName, ""))
        key.Close()
    End Function

    Public Shared Function GetRegInteger(ByVal RegKey As String, ByVal KeyName As String) As Integer
        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegKey)
        GetRegInteger = Convert.ToInt32(key.GetValue(KeyName, 0))
        key.Close()
    End Function

    Public Shared Sub SetRegValue(ByVal RegKey As String, ByVal KeyName As String, ByVal Value As Object)
        Dim key As Microsoft.Win32.RegistryKey
        key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegKey)
        key.SetValue(KeyName, Value)
        key.Close()
    End Sub

#End Region

#Region " Graphics "

    Public Shared Sub DrawRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal Width As Integer,
      ByVal Height As Integer,
      ByVal Diameter As Integer,
      ByVal Stroke As Integer)

        DrawRoundedRectangle(objGraphics, x, y, Width, Height, Diameter, Stroke, System.Drawing.Color.Black)

    End Sub

    Public Shared Sub DrawRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics,
    ByVal x As Integer,
    ByVal y As Integer,
    ByVal Width As Integer,
    ByVal Height As Integer,
    ByVal Diameter As Integer,
    ByVal Stroke As Integer,
    ByVal PenColor As System.Drawing.Color)

        Dim p As New System.Drawing.Pen(PenColor, Stroke)
        'Dim g As Graphics
        Dim BaseRect As New System.Drawing.RectangleF(x, y, Width, Height)
        Dim ArcRect As New System.Drawing.RectangleF(BaseRect.Location, New System.Drawing.SizeF(Diameter, Diameter))
        'top left Arc
        objGraphics.DrawArc(p, ArcRect, 180, 90)
        objGraphics.DrawLine(p, x + CInt(Diameter / 2), y, x + Width - CInt(Diameter / 2), y)

        ' top right arc
        ArcRect.X = BaseRect.Right - Diameter
        objGraphics.DrawArc(p, ArcRect, 270, 90)
        objGraphics.DrawLine(p, x + Width, y + CInt(Diameter / 2), x + Width, y + Height - CInt(Diameter / 2))

        ' bottom right arc
        ArcRect.Y = BaseRect.Bottom - Diameter
        objGraphics.DrawArc(p, ArcRect, 0, 90)
        objGraphics.DrawLine(p, x + CInt(Diameter / 2), y + Height, x + Width - CInt(Diameter / 2), y + Height)

        ' bottom left arc
        ArcRect.X = BaseRect.Left
        objGraphics.DrawArc(p, ArcRect, 90, 90)
        objGraphics.DrawLine(p, x, y + CInt(Diameter / 2), x, y + Height - CInt(Diameter / 2))

        p.Dispose()

    End Sub

    Public Shared Sub FillRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics,
      ByVal brush As System.Drawing.Brush,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal Width As Integer,
      ByVal Height As Integer,
      ByVal Diameter As Integer)

        'Dim g As Graphics
        Dim BaseRect As New System.Drawing.RectangleF(x, y, Width, Height)
        Dim ArcRect As New System.Drawing.RectangleF(BaseRect.Location, New System.Drawing.SizeF(Diameter, Diameter))
        Dim SmallRect As New System.Drawing.RectangleF(x, y + Diameter \ 2, Width, Height - Diameter)

        'top left Arc
        objGraphics.FillEllipse(brush, ArcRect)

        ' top right arc
        ArcRect.X = BaseRect.Right - Diameter
        objGraphics.FillEllipse(brush, ArcRect)

        ' bottom right arc
        ArcRect.Y = BaseRect.Bottom - Diameter
        objGraphics.FillEllipse(brush, ArcRect)

        ' bottom left arc
        ArcRect.X = BaseRect.Left
        objGraphics.FillEllipse(brush, ArcRect)

        ' center rectangle
        objGraphics.FillRectangle(brush, SmallRect)

        ' top rectangle
        SmallRect.X = x + Diameter \ 2
        SmallRect.Y = y
        SmallRect.Width = Width - Diameter
        SmallRect.Height = Diameter \ 2
        objGraphics.FillRectangle(brush, SmallRect)

        ' bottop rectangle
        SmallRect.X = x + Diameter \ 2
        SmallRect.Y = y + Height - Diameter \ 2
        SmallRect.Width = Width - Diameter
        SmallRect.Height = Diameter \ 2
        objGraphics.FillRectangle(brush, SmallRect)

    End Sub


    Public Shared Sub FillTopRoundedRectangle(ByVal objGraphics As System.Drawing.Graphics,
      ByVal brush As System.Drawing.Brush,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal Width As Integer,
      ByVal Height As Integer,
      ByVal Diameter As Integer)

        'Dim g As Graphics
        Dim BaseRect As New System.Drawing.RectangleF(x, y, Width, Height)
        Dim ArcRect As New System.Drawing.RectangleF(BaseRect.Location, New System.Drawing.SizeF(Diameter, Diameter))
        Dim RectArea As New System.Drawing.RectangleF

        'top left Arc
        objGraphics.FillEllipse(brush, ArcRect)

        ' top right arc
        ArcRect.X = BaseRect.Right - Diameter
        objGraphics.FillEllipse(brush, ArcRect)

        ' bottom right arc
        ArcRect.Y = BaseRect.Bottom - Diameter
        objGraphics.FillEllipse(brush, ArcRect)

        ' bottom left arc
        ArcRect.X = BaseRect.Left
        objGraphics.FillEllipse(brush, ArcRect)

        ' top rectangle
        RectArea.X = x + Diameter \ 2
        RectArea.Y = y
        RectArea.Width = Width - Diameter
        RectArea.Height = Diameter \ 2
        objGraphics.FillRectangle(brush, RectArea)

        ' bottom rectangle
        RectArea.X = x
        RectArea.Y = y + Diameter \ 2
        RectArea.Width = Width
        RectArea.Height = Height - Diameter \ 2
        objGraphics.FillRectangle(brush, RectArea)


    End Sub

    Public Shared Sub DrawRotateText(ByVal objGraphics As System.Drawing.Graphics,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal Angle As Integer,
      ByVal Text As String,
      ByVal Fnt As System.Drawing.Font,
      ByVal brsh As System.Drawing.Brush)

        '  Rotating

        'Another interesting transformation of the coordinate system is its ability to rotate. 
        'This allows for fancy tricks such as rendering text at an angle: 

        objGraphics.TranslateTransform(x, y) ' -> desplaça l'orige on començara a escriure
        objGraphics.RotateTransform(Angle) ' -> aplica un rotacio en el sentit de les agulles del relotge de 35º
        objGraphics.DrawString(Text, Fnt, brsh, 0, 0)
        objGraphics.ResetTransform() ' Desfà les trasformacions. les coordenades tornen al seu lloc.

    End Sub

    Public Shared Sub DrawRotateImage(ByVal gr As System.Drawing.Graphics, ByVal bmp As System.Drawing.Bitmap, ByVal x As Integer, ByVal y As Integer, ByVal angle As Single)
        angle = CSng(angle / 180 / Math.PI)
        Dim x1 As Integer = CInt(x + bmp.Width * Math.Cos(angle))
        Dim y1 As Integer = CInt(y + bmp.Width * Math.Sin(angle))
        Dim x2 As Integer = CInt(x - bmp.Height * Math.Sin(angle))
        Dim y2 As Integer = CInt(y + bmp.Height * Math.Cos(angle))
        Dim points() As Point = {New Point(x, y), New Point(x1, y1), New Point(x2, y2)}
        gr.DrawImage(bmp, points)
    End Sub

    Public Shared Sub DrawRotateImage(ByVal objGraphics As System.Drawing.Graphics,
      ByVal Image As System.Drawing.Image,
      ByVal Angle As Integer,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal width As Integer,
      ByVal height As Integer)

        '  Rotating

        'Another interesting transformation of the coordinate system is its ability to rotate. 
        'This allows for fancy tricks such as rendering text at an angle: 
        Dim gs As System.Drawing.Drawing2D.GraphicsState = objGraphics.Save()

        objGraphics.TranslateTransform(x, y) ' -> desplaça l'orige on començara a escriure
        objGraphics.RotateTransform(Angle) ' -> aplica un rotacio en el sentit de les agulles del relotge 
        objGraphics.DrawImage(Image, 0, 0, width, height)
        objGraphics.ResetTransform() ' Desfà les trasformacions. les coordenades tornen al seu lloc.
        objGraphics.Restore(gs)

    End Sub

    Public Shared Sub DrawFittedText(ByVal objGraphics As System.Drawing.Graphics,
      ByVal x As Integer,
      ByVal y As Integer,
      ByVal width As Integer,
      ByVal height As Integer,
      ByVal Text As String,
      ByVal Fnt As System.Drawing.Font,
      ByVal foreBrush As System.Drawing.Brush,
      ByVal backBrush As System.Drawing.Brush,
      ByVal Stroke As Integer,
      ByVal Diameter As Integer)

        '  Rotating

        'Another interesting transformation of the coordinate system is its ability to rotate. 
        'This allows for fancy tricks such as rendering text at an angle: 

        Using Fnt

            Dim ds As SizeF = objGraphics.MeasureString(Text, Fnt, New Point(0, 0), Nothing)
            ' Netejem la superficie
            objGraphics.FillRectangle(Brushes.White, x, y, width, height)

            If Diameter > 0 Then
                FillRoundedRectangle(objGraphics, backBrush, x, y, width, height, Diameter)
            Else
                objGraphics.FillRectangle(backBrush, x, y, width, height)
            End If

            ' Si stroke es mes gran que 0 va una rebora:
            If Stroke > 0 Then
                ' si el diametre es mes gran que 0 es redondejat:
                If Diameter > 0 Then
                    DrawRoundedRectangle(objGraphics, x, y, width, height, Diameter, Stroke)
                Else
                    Dim p As New System.Drawing.Pen(System.Drawing.Color.Black, Stroke)
                    objGraphics.DrawRectangle(p, x, y, width, height)
                End If
            End If

            Dim gs As System.Drawing.Drawing2D.GraphicsState = objGraphics.Save()
            objGraphics.TranslateTransform(x, y)
            objGraphics.ScaleTransform(CSng(width / ds.Width), CSng(height / ds.Height))
            objGraphics.DrawString(Text, Fnt, foreBrush, 0, 0)
            objGraphics.Restore(gs)

        End Using


    End Sub



    '  Public Shared Function createMetafile() As System.Drawing.Imaging.Metafile

    '  private Metafile createMetafile( Page page )
    '  Metafile metafile = null; 

    '   // create a Metafile object that is compatible with the surface of this 
    '   // form
    '   using ( Graphics graphics = this.CreateGraphics() )
    '   { 
    '      System.IntPtr hdc = graphics.GetHdc(); 
    '      metafile = new Metafile( hdc, new Rectangle( 0, 0, 
    '           (int) page.Width, (int) page.Height ), MetafileFrameUnit.Point ); 
    '      graphics.ReleaseHdc( hdc );
    '   }

    '   // draw to the metafile
    '   using ( Graphics metafileGraphics = Graphics.FromImage( metafile ) )
    '   {
    '      metafileGraphics.SmoothingMode = SmoothingMode.AntiAlias; // smooth the 
    '                                                                // output
    '      page.Draw( metafileGraphics );
    '   }

    '   return metafile;

    '  End Function


    Shared Sub SaveJPG(ByVal path As String, ByVal Image As Image, ByVal quality As Integer)
        Dim ePs As System.Drawing.Imaging.EncoderParameters = New System.Drawing.Imaging.EncoderParameters(1)
        Dim _jpeg As System.Drawing.Imaging.ImageCodecInfo = Nothing
        ' L'argument 'quality' doit être compris entre 0 et 100.

        Try
            For Each codec As System.Drawing.Imaging.ImageCodecInfo In System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders
                If codec.MimeType = "image/jpeg" Then
                    _jpeg = codec
                End If
            Next
        Catch ex As Exception
        End Try

        If IsNothing(Image) Then
            Throw New ArgumentNullException("image")
        End If
        If quality < 0 Or quality > 100 Then
            Throw New ArgumentOutOfRangeException("quality")
        End If
        If IsNothing(_jpeg) Then
            Throw New InvalidOperationException("Impossible de trouver un encodeur Jpeg.")
        End If

        ePs.Param(0) = New System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, CLng(quality))
        Image.Save(path, _jpeg, ePs)

    End Sub

#End Region

#Region " INI files "

    Public Class IniFile
        ' API functions
        Private Declare Ansi Function GetPrivateProfileString _
          Lib "kernel32.dll" Alias "GetPrivateProfileStringA" _
          (ByVal lpApplicationName As String,
          ByVal lpKeyName As String, ByVal lpDefault As String,
          ByVal lpReturnedString As System.Text.StringBuilder,
          ByVal nSize As Integer, ByVal lpFileName As String) _
          As Integer
        Private Declare Ansi Function WritePrivateProfileString _
          Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
          (ByVal lpApplicationName As String,
          ByVal lpKeyName As String, ByVal lpString As String,
          ByVal lpFileName As String) As Integer
        Private Declare Ansi Function GetPrivateProfileInt _
          Lib "kernel32.dll" Alias "GetPrivateProfileIntA" _
          (ByVal lpApplicationName As String,
          ByVal lpKeyName As String, ByVal nDefault As Integer,
          ByVal lpFileName As String) As Integer
        Private Declare Ansi Function FlushPrivateProfileString _
          Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
          (ByVal lpApplicationName As Integer,
          ByVal lpKeyName As Integer, ByVal lpString As Integer,
          ByVal lpFileName As String) As Integer
        Dim strFilename As String

        ' Constructor, accepting a filename
        Public Sub New(ByVal Filename As String)
            strFilename = Filename
        End Sub

        ' Read-only filename property
        ReadOnly Property FileName() As String
            Get
                Return strFilename
            End Get
        End Property

        Public Function GetString(ByVal Section As String,
          ByVal Key As String, ByVal [Default] As String) As String
            ' Returns a string from your INI file
            Dim intCharCount As Integer
            Dim objResult As New System.Text.StringBuilder(256)
            intCharCount = GetPrivateProfileString(Section, Key,
               [Default], objResult, objResult.Capacity, strFilename)
            If intCharCount > 0 Then
                GetString = Left(objResult.ToString, intCharCount)
            Else
                GetString = ""
            End If
        End Function

        Public Function GetInteger(ByVal Section As String,
          ByVal Key As String, ByVal [Default] As Integer) As Integer
            ' Returns an integer from your INI file
            Return GetPrivateProfileInt(Section, Key,
               [Default], strFilename)
        End Function

        Public Function GetBoolean(ByVal Section As String,
          ByVal Key As String, ByVal [Default] As Boolean) As Boolean
            ' Returns a boolean from your INI file
            Return (GetPrivateProfileInt(Section, Key,
               CInt([Default]), strFilename) = 1)
        End Function

        Public Sub WriteString(ByVal Section As String,
          ByVal Key As String, ByVal Value As String)
            ' Writes a string to your INI file
            WritePrivateProfileString(Section, Key, Value, strFilename)
            Flush()
        End Sub

        Public Sub WriteInteger(ByVal Section As String,
          ByVal Key As String, ByVal Value As Integer)
            ' Writes an integer to your INI file
            WriteString(Section, Key, CStr(Value))
            Flush()
        End Sub

        Public Sub WriteBoolean(ByVal Section As String,
          ByVal Key As String, ByVal Value As Boolean)
            ' Writes a boolean to your INI file
            WriteString(Section, Key, CStr(CInt(Value)))
            Flush()
        End Sub

        Private Sub Flush()
            ' Stores all the cached changes to your INI file
            FlushPrivateProfileString(0, 0, 0, strFilename)
        End Sub

    End Class

#End Region

#Region " Controls "

    Public Shared Sub FillComboBoxFromArray(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal Items() As String)

        Dim oData As DataTable = Nothing
        Dim oRow As DataRow = Nothing
        Dim oColumn As DataColumn = Nothing

        '-------------------------------------------------------------
        ' Create the DataTable
        oData = New DataTable

        oColumn = New DataColumn("Key", GetType(System.Int32))
        oData.Columns.Add("Key")

        oColumn = New DataColumn("Value", GetType(System.String))
        oData.Columns.Add("Value")
        '-------------------------------------------------------------

        '-------------------------------------------------------------
        ' Add the enum items to the datatable
        For Each iItem As String In Items
            Try
                oRow = oData.NewRow()
                oRow("Key") = CType(iItem.Split(","c)(0).Trim, Int32)
                oRow("Value") = iItem.Split(","c)(1).Trim
                oData.Rows.Add(oRow)
            Catch ex As Exception
            End Try
        Next
        '-------------------------------------------------------------

        ComboBoxControl.DataSource = oData
        ComboBoxControl.ValueMember = "Key"
        ComboBoxControl.DisplayMember = "Value"

    End Sub

    Public Shared Sub FillComboBoxFromArrayList(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal Items As ArrayList)

        Dim oData As DataTable = Nothing
        Dim oRow As DataRow = Nothing
        Dim oColumn As DataColumn = Nothing

        '-------------------------------------------------------------
        ' Create the DataTable
        oData = New DataTable

        oColumn = New DataColumn("Key", GetType(System.Int32))
        oData.Columns.Add("Key")

        oColumn = New DataColumn("Value", GetType(System.String))
        oData.Columns.Add("Value")
        '-------------------------------------------------------------

        '-------------------------------------------------------------
        ' Add the enum items to the datatable
        For Each iItem() As String In Items
            Try
                oRow = oData.NewRow()
                oRow("Key") = CType(iItem(0).Trim, Int32)
                oRow("Value") = iItem(1).ToString
                oData.Rows.Add(oRow)
            Catch ex As Exception
            End Try
        Next
        '-------------------------------------------------------------

        ComboBoxControl.DataSource = oData
        ComboBoxControl.ValueMember = "Key"
        ComboBoxControl.DisplayMember = "Value"

    End Sub

    Public Shared Sub FillComboBoxFromEnum(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal EnumType As Type)
        Dim oData As DataTable

        ' Notice that we must use 'GetType(Enumeration)'
        oData = EnumToDataTable(EnumType, "Key", "Value")
        ComboBoxControl.DataSource = oData
        ComboBoxControl.ValueMember = "Key"
        ComboBoxControl.DisplayMember = "Value"

    End Sub

    Public Shared Sub FillComboBoxFromEnum(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal EnumType As Type, ByVal ValuesToExclude As ArrayList)
        Dim oData As DataTable

        ' Notice that we must use 'GetType(Enumeration)'
        oData = EnumToDataTable(EnumType, "Key", "Value", ValuesToExclude)
        ComboBoxControl.DataSource = oData
        ComboBoxControl.ValueMember = "Key"
        ComboBoxControl.DisplayMember = "Value"

    End Sub

    Public Shared Sub FillComboBoxFromTable(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal Table As Data.DataTable)

        ' Notice that we must use 'GetType(Enumeration)'
        ComboBoxControl.DataSource = Table
        ComboBoxControl.ValueMember = Table.Columns(0).ColumnName
        ComboBoxControl.DisplayMember = Table.Columns(1).ColumnName

    End Sub

    Public Shared Sub FillComboBoxFromTable(ByVal ComboBoxControl As Windows.Forms.ComboBox, ByVal Table As Data.DataTable, ByVal ColumnNameValueMember As String, ByVal ColumnNameDisplayMember As String)

        ' Notice that we must use 'GetType(Enumeration)'
        ComboBoxControl.DataSource = Table
        ComboBoxControl.ValueMember = ColumnNameValueMember
        ComboBoxControl.DisplayMember = ColumnNameDisplayMember

    End Sub

    Public Shared Function EnumToDataTable(ByVal EnumObject As Type,
       ByVal KeyField As String, ByVal ValueField As String) As DataTable
        Return EnumToDataTable(EnumObject, KeyField, ValueField, Nothing)
    End Function

    Public Shared Function EnumToDataTable(ByVal EnumObject As Type,
       ByVal KeyField As String, ByVal ValueField As String, ByVal ValuesToExclude As ArrayList) As DataTable

        Dim oData As DataTable = Nothing
        Dim oRow As DataRow = Nothing
        Dim oColumn As DataColumn = Nothing

        '-------------------------------------------------------------
        ' Sanity check
        If KeyField.Trim() = String.Empty Then
            KeyField = "KEY"
        End If

        If ValueField.Trim() = String.Empty Then
            ValueField = "VALUE"
        End If
        '-------------------------------------------------------------

        '-------------------------------------------------------------
        ' Create the DataTable
        oData = New DataTable

        oColumn = New DataColumn(KeyField, GetType(System.Int32))
        oData.Columns.Add(KeyField)

        oColumn = New DataColumn(ValueField, GetType(System.String))
        oData.Columns.Add(ValueField)
        '-------------------------------------------------------------

        '-------------------------------------------------------------
        ' Add the enum items to the datatable
        For Each iEnumItem As Object In [Enum].GetValues(EnumObject)
            If Not IsNothing(ValuesToExclude) Then
                If ValuesToExclude.Contains(CObj(CType(iEnumItem, Int32))) Then
                    Continue For
                End If
            End If
            oRow = oData.NewRow()
            oRow(KeyField) = CType(iEnumItem, Int32)
            oRow(ValueField) = StrConv(Replace(iEnumItem.ToString(), "_", " "),
                  VbStrConv.ProperCase)
            oData.Rows.Add(oRow)
        Next
        '-------------------------------------------------------------

        Return oData

    End Function

#End Region

#Region " Math "

    Public Shared Function Arrodonir(ByVal Value As Decimal, ByVal Dec As Integer) As Decimal
        Return CDec(Math.Truncate(Value * 10 ^ Dec + 0.5) / 10 ^ Dec)
    End Function

    Public Class LinearRegression
        ' Y = mX + n
        Private MustCalculate As Boolean
        Private nX() As Double
        Private nY() As Double
        Private nM As Double
        Private nN As Double
        Private nYmax As Double
        Private nYmin As Double

        Public WriteOnly Property X() As Double()
            Set(ByVal value As Double())
                nX = value
                MustCalculate = True
            End Set
        End Property

        Public WriteOnly Property Y() As Double()
            Set(ByVal value As Double())
                nY = value
                MustCalculate = True
            End Set
        End Property

        Public ReadOnly Property Pdte() As Double
            Get
                If MustCalculate Then
                    Calculate()
                End If
                Return nM
            End Get
        End Property

        Public ReadOnly Property Abcs() As Double
            Get
                If MustCalculate Then
                    Calculate()
                End If
                Return nN
            End Get
        End Property

        Public ReadOnly Property Ymax() As Double
            Get
                If MustCalculate Then
                    Calculate()
                End If
                Return nYmax
            End Get
        End Property

        Public ReadOnly Property Ymin() As Double
            Get
                If MustCalculate Then
                    Calculate()
                End If
                Return nYmin
            End Get
        End Property

        Public Sub Calculate()
            Dim Xavg As Double
            Dim Yavg As Double
            Dim A As Double
            Dim B As Double

            A = 0
            B = 0
            Xavg = 0
            Yavg = 0

            nYmax = nY(0)
            nYmin = nY(0)

            For i As Integer = 0 To nX.Length - 1
                Xavg += nX(i)
                Yavg += nY(i)
                If nY(i) > nYmax Then
                    nYmax = nY(i)
                End If
                If nY(i) < nYmin Then
                    nYmin = nY(i)
                End If
            Next

            Xavg /= nX.Length
            Yavg /= nY.Length

            For i As Integer = 0 To nX.Length - 1
                A += ((nX(i) - Xavg) * (nY(i) - Yavg))
                B += (nX(i) - Xavg) ^ 2
            Next

            Try
                nM = A / B
                nN = Yavg - nM * Xavg
            Catch ex As Exception
                nM = 0
                nN = 0
            End Try

            MustCalculate = False

        End Sub

        Public Function GetY(ByVal X As Double) As Double
            If MustCalculate Then
                Calculate()
            End If
            Return nM * X + nN
        End Function

    End Class

#End Region

#Region " Printing "

#Region " Structure and API declarions "
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Structure DOCINFOW
        <MarshalAs(UnmanagedType.LPWStr)> Public pDocName As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pOutputFile As String
        <MarshalAs(UnmanagedType.LPWStr)> Public pDataType As String
    End Structure
    <DllImport("winspool.Drv", EntryPoint:="OpenPrinterW",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function OpenPrinter(ByVal src As String, ByRef hPrinter As IntPtr, ByVal pd As Integer) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="ClosePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function ClosePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="StartDocPrinterW",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function StartDocPrinter(ByVal hPrinter As IntPtr, ByVal level As Int32, ByRef pDI As DOCINFOW) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="EndDocPrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function EndDocPrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="StartPagePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function StartPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="EndPagePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function EndPagePrinter(ByVal hPrinter As IntPtr) As Boolean
    End Function
    <DllImport("winspool.Drv", EntryPoint:="WritePrinter",
       SetLastError:=True, CharSet:=CharSet.Unicode,
       ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function WritePrinter(ByVal hPrinter As IntPtr, ByVal pBytes As IntPtr, ByVal dwCount As Int32, ByRef dwWritten As Int32) As Boolean
    End Function
#End Region

    ' SendBytesToPrinter()
    ' When the function is given a printer name and an unmanaged array of  
    ' bytes, the function sends those bytes to the print queue.
    ' Returns True on success or False on failure.
    Public Shared Function SendBytesToPrinter(ByVal szDocName As String, ByVal szPrinterName As String, ByVal pBytes As IntPtr, ByVal dwCount As Int32) As Boolean
        Dim hPrinter As IntPtr      ' The printer handle.
        Dim dwError As Int32        ' Last error - in case there was trouble.
        Dim di As DOCINFOW          ' Describes your document (name, port, data type).
        Dim dwWritten As Int32      ' The number of bytes written by WritePrinter().
        Dim bSuccess As Boolean     ' Your success code.

        ' Set up the DOCINFO structure.
        With di
            .pDocName = szDocName
            .pDataType = "RAW"
        End With
        ' Assume failure unless you specifically succeed.
        bSuccess = False
        If OpenPrinter(szPrinterName, hPrinter, 0) Then
            If StartDocPrinter(hPrinter, 1, di) Then
                If StartPagePrinter(hPrinter) Then
                    ' Write your printer-specific bytes to the printer.
                    bSuccess = WritePrinter(hPrinter, pBytes, dwCount, dwWritten)
                    EndPagePrinter(hPrinter)
                End If
                EndDocPrinter(hPrinter)
            End If
            ClosePrinter(hPrinter)
        End If
        ' If you did not succeed, GetLastError may give more information
        ' about why not.
        If bSuccess = False Then
            dwError = Marshal.GetLastWin32Error()
        End If
        Return bSuccess
    End Function ' SendBytesToPrinter()

    ' SendFileToPrinter()
    ' When the function is given a file name and a printer name, 
    ' the function reads the contents of the file and sends the
    ' contents to the printer.
    ' Presumes that the file contains printer-ready data.
    ' Shows how to use the SendBytesToPrinter function.
    ' Returns True on success or False on failure.
    Public Shared Function SendFileToPrinter(ByVal szDocName As String, ByVal szPrinterName As String, ByVal szFileName As String) As Boolean
        ' Open the file.
        Dim fs As New FileStream(szFileName, FileMode.Open)
        ' Create a BinaryReader on the file.
        Dim br As New BinaryReader(fs)
        ' Dim an array of bytes large enough to hold the file's contents.
        Dim bytes(CInt(fs.Length)) As Byte
        Dim bSuccess As Boolean
        ' Your unmanaged pointer.
        Dim pUnmanagedBytes As IntPtr

        ' Read the contents of the file into the array.
        bytes = br.ReadBytes(CInt(fs.Length))
        ' Allocate some unmanaged memory for those bytes.
        pUnmanagedBytes = Marshal.AllocCoTaskMem(CInt(fs.Length))
        ' Copy the managed byte array into the unmanaged array.
        Marshal.Copy(bytes, 0, pUnmanagedBytes, CInt(fs.Length))
        ' Send the unmanaged bytes to the printer.
        bSuccess = SendBytesToPrinter(szDocName, szPrinterName, pUnmanagedBytes, CInt(fs.Length))
        ' Free the unmanaged memory that you allocated earlier.
        Marshal.FreeCoTaskMem(pUnmanagedBytes)
        Return bSuccess
    End Function ' SendFileToPrinter()

    ' When the function is given a string and a printer name,
    ' the function sends the string to the printer as raw bytes.
    Public Shared Function SendStringToPrinter(ByVal szDocName As String, ByVal szPrinterName As String, ByVal szString As String) As Boolean
        Dim pBytes As IntPtr
        Dim dwCount As Int32
        ' How many characters are in the string?
        dwCount = szString.Length()
        ' Assume that the printer is expecting ANSI text, and then convert
        ' the string to ANSI text.
        pBytes = Marshal.StringToCoTaskMemAnsi(szString)
        ' Send the converted ANSI string to the printer.
        SendBytesToPrinter(szDocName, szPrinterName, pBytes, dwCount)
        Marshal.FreeCoTaskMem(pBytes)
        Return True
    End Function

    Public Shared Function SendStringUnicodeToPrinter(ByVal szDocName As String, ByVal szPrinterName As String, ByVal szString As String) As Boolean
        Dim pBytes As IntPtr
        Dim dwCount As Int32
        ' How many characters are in the string?
        dwCount = szString.Length()
        ' Assume that the printer is expecting ANSI text, and then convert
        ' the string to ANSI text.
        pBytes = Marshal.StringToCoTaskMemUni(szString)
        ' Send the converted ANSI string to the printer.
        SendBytesToPrinter(szDocName, szPrinterName, pBytes, dwCount)
        Marshal.FreeCoTaskMem(pBytes)
        Return True
    End Function


    Public Shared Function SendStreamToPrinter(ByVal szDocName As String, ByVal szPrinterName As String, ByVal memStream As Stream) As Boolean
        Dim bytes() As Byte
        Dim pUnmanagedBytes As IntPtr
        Dim bSuccess As Boolean
        Dim Count As Integer
        Dim br As BinaryReader

        br = New BinaryReader(memStream)
        Count = CInt(memStream.Length)
        memStream.Seek(0, SeekOrigin.Begin)
        bytes = br.ReadBytes(Count)
        pUnmanagedBytes = Marshal.AllocCoTaskMem(Count)
        Marshal.Copy(bytes, 0, pUnmanagedBytes, Count)
        bSuccess = SendBytesToPrinter(szDocName, szPrinterName, pUnmanagedBytes, Count)
        Marshal.FreeCoTaskMem(pUnmanagedBytes)

        Return True
    End Function

#End Region

#Region " Excel "

    Public Class Export2Excel
        Public Reader2Export As SqlClient.SqlDataReader
        Private xls As New C1.C1Excel.C1XLBook()

        Public Sub AddColumnDef(ByVal Title As String, ByVal xDataType As DataTypeEnum, ByVal FieldName As String)
            Columns.Add(New Column(Title, xDataType, FieldName))
        End Sub

        Public Sub AddCaption(ByVal Caption As String, ByVal Text As String)
            Captions.Add(New Caption(Caption, Text))
        End Sub

        Public Sub AddTitle(ByVal Title As String)
            SheetTitle = Title
        End Sub

        Public Enum DataTypeEnum
            [String]
            [Integer]
            [Decimal]
            [DateTime]
            [Formula]
        End Enum

        Private Structure Caption
            Public Caption As String
            Public Text As String
            Public Sub New(ByVal Caption As String, ByVal Text As String)
                Me.Caption = Caption
                Me.Text = Text
            End Sub
        End Structure

        Private Structure Column
            Public Title As String
            Public xDataType As DataTypeEnum
            Public FieldName As String
            Public Sub New(ByVal Title As String, ByVal xDataType As DataTypeEnum, ByVal FieldName As String)
                Me.Title = Title
                Me.xDataType = xDataType
                Me.FieldName = FieldName
            End Sub
        End Structure

        Private Columns As New Collections.Generic.List(Of Column)
        Private Captions As New Collections.Generic.List(Of Caption)
        Private SheetTitle As String

        Public Sub DataReader2Excel(ByVal ReaderToExport As IDataReader, ByVal FileName As String, ByVal SheetIndex As Integer, ByVal SheetName As String)

            DataReader2Excel(ReaderToExport, SheetIndex, SheetName)
            Save(FileName, False)

        End Sub

        Public Sub DataReader2Excel(ByVal ReaderToExport As IDataReader, ByVal SheetIndex As Integer, ByVal SheetName As String)
            Dim sheetStyle As New C1.C1Excel.XLStyle(xls)
            Dim sheet As C1.C1Excel.XLSheet
            Dim colCount As Integer
            Dim rowCount As Integer

            While xls.Sheets.Count < SheetIndex + 1
                xls.Sheets.Add()
            End While

            sheet = xls.Sheets(SheetIndex)
            If Not String.IsNullOrEmpty(SheetName) Then
                sheet.Name = SheetName
            End If

            sheetStyle.Font = New Font("Arial", 20, FontStyle.Bold)

            rowCount = 0

            If Not String.IsNullOrEmpty(SheetTitle) Then
                sheet(0, 0).Value = SheetTitle
                sheet(0, 0).Style = sheetStyle
                rowCount += 1
            End If

            rowCount += 1
            For Each c As Caption In Captions
                sheet(rowCount, 0).Value = c.Caption
                sheet(rowCount, 1).Value = c.Text
                rowCount += 1
            Next

            rowCount += 1
            colCount = 0
            For Each c As Column In Columns
                sheet(rowCount, colCount).Value = c.Title
                colCount += 1
            Next

            Do While ReaderToExport.Read
                rowCount += 1
                colCount = 0
                For Each c As Column In Columns

                    Select Case c.xDataType
                        Case DataTypeEnum.Integer
                            If Not Utils.IsNullOrEmptyValue(ReaderToExport(c.FieldName)) Then
                                sheet(rowCount, colCount).Value = CInt(ReaderToExport(c.FieldName))
                            Else
                                sheet(rowCount, colCount).Value = 0
                            End If
                        Case DataTypeEnum.Decimal
                            If Not Utils.IsNullOrEmptyValue(ReaderToExport(c.FieldName)) Then
                                sheet(rowCount, colCount).Value = CDec(ReaderToExport(c.FieldName))
                            Else
                                sheet(rowCount, colCount).Value = 0D
                            End If
                        Case DataTypeEnum.DateTime
                            If Not Utils.IsNullOrEmptyValue(ReaderToExport(c.FieldName)) Then
                                sheet(rowCount, colCount).Value = CStr(Date.Parse(String.Format("{0}", ReaderToExport(c.FieldName))))
                            Else
                                sheet(rowCount, colCount).Value = ReaderToExport(c.FieldName)
                            End If
                        Case DataTypeEnum.Formula
                            'Cal actualitzar la versió de C1
                            'sheet(rowCount, colCount).Formula = c.FieldName
                            sheet(rowCount, colCount).Value = c.FieldName
                        Case Else
                            sheet(rowCount, colCount).Value = ReaderToExport(c.FieldName)
                    End Select

                    colCount += 1

                Next

            Loop

            ReaderToExport.Close()

        End Sub

        Public Sub DataTable2Excel(ByVal TableToExport As DataTable, ByVal FileName As String, ByVal SheetIndex As Integer, ByVal SheetName As String)
            DataReader2Excel(TableToExport.CreateDataReader, FileName, SheetIndex, SheetName)
        End Sub

        Public Sub DataTable2Excel(ByVal TableToExport As DataTable, ByVal SheetIndex As Integer, ByVal SheetName As String)
            DataReader2Excel(TableToExport.CreateDataReader, SheetIndex, SheetName)
        End Sub

        Public Sub Save(ByVal FileName As String, ByVal ShowSaveDialog As Boolean)
            If ShowSaveDialog Or String.IsNullOrEmpty(FileName) Then
                Dim sdlg As New SaveFileDialog
                sdlg.FileName = FileName
                sdlg.Filter = "Fitxers Excel (*.xls)|*.xls|" & "Tots els fitxer|*.*"
                sdlg.AddExtension = True
                sdlg.CheckPathExists = True
                sdlg.DefaultExt = "xls"
                sdlg.RestoreDirectory = True
                sdlg.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                If sdlg.ShowDialog = DialogResult.OK Then
                    FileName = sdlg.FileName
                Else
                    FileName = String.Empty
                End If
            End If
            If Not String.IsNullOrEmpty(FileName) Then
                xls.Save(FileName)
            End If
        End Sub

        Public Sub ClearColumns()
            Columns.Clear()
        End Sub

        Public Sub ClearCaptions()
            Captions.Clear()
        End Sub

    End Class

#End Region

#Region " Databases "
    Public Class csTableReader
        Private Pointer As Integer

        Private mInternalDataTable As DataTable
        Public Property InternalDataTable() As DataTable
            Get
                Return mInternalDataTable
            End Get
            Set(ByVal value As DataTable)
                mInternalDataTable = value
                Pointer = -1
            End Set
        End Property

        Default Public ReadOnly Property Row(ByVal field As String) As Object
            Get
                Dim oData As Object
                Try
                    oData = InternalDataTable.Rows(Pointer)(field)
                Catch ex As Exception
                    oData = Nothing
                End Try
                Return oData
            End Get
        End Property

        Public ReadOnly Property DataRow(ByVal field As String) As DataRow
            Get
                Dim oData As DataRow
                Try
                    oData = InternalDataTable.Rows(Pointer)
                Catch ex As Exception
                    oData = Nothing
                End Try
                Return oData
            End Get
        End Property

        Public ReadOnly Property HasRows() As Boolean
            Get
                If InternalDataTable Is Nothing Then
                    Return False
                Else
                    If InternalDataTable.Rows.Count > 0 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property

        Public Sub New()
            InternalDataTable = Nothing
            Pointer = -1
        End Sub

        Public Sub New(ByRef Table As DataTable)
            InternalDataTable = Table
            Pointer = -1
        End Sub

        Public Sub Rewind()
            Pointer = -1
        End Sub

        Public Function Read() As Boolean
            Pointer += 1
            If Pointer >= InternalDataTable.Rows.Count Then
                Return False
            End If
            Return True
        End Function

    End Class
#End Region

End Class

