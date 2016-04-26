Imports csAppData
Imports csUtils
Imports csUtils.Utils

Public Class R90FAC0002A_Factura
  Inherits ReportBaseClass
  Private dbaFac As New C00_gst_fac
  Private dbaCli As New C00_gst_cli

  Public Overrides Sub Execute(workInfo As String)
    Dim rpt As New R90FAC0002C_Factura
    Dim tbFacs As DataTable

    Dim PrinterID As Integer
    Dim origenDades As String
    Dim serie As String
    Dim any_ As String
    Dim Numero As Integer
    Dim numCopies As Integer
    Dim Output As String
    Dim eMail As String
    Dim mail As csFaxMail
    Dim NumFactura As String

    Dim ReportOutput As TipusEnviamentDocumentEnum

    Dim params As FetchXmlParameter

    Try

      rpt.CustomID = My.Settings.ClientCustom
      rpt.ShowForm = False
      rpt.PageNumbering = csRpt.PageNumberEnum.PageNofM

      If String.IsNullOrEmpty(workInfo) Then
        rpt.Print()
        Return
      End If

      params = New FetchXmlParameter(workInfo)
      PrinterID = CNull(params.GetValue("PrinterID"), 0)
      origenDades = CNull(params.GetValue("OrigenDades"))
      serie = CNull(params.GetValue("Serie"))
      any_ = CNull(params.GetValue("Any"))
      Numero = CVal(params.GetValue("Numero"))
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


      tbFacs = dbaFac.GetFactures(origenDades, serie, "", "", any_, any_, Numero, Numero, Nothing, Nothing, "t")

      If tbFacs.Rows.Count = 0 Then
        Return
      End If

      If ReportOutput = TipusEnviamentDocumentEnum.Pdf Then
        If tbFacs.Rows.Count > 1 Then
          rpt.pdfNumberOfJobs = tbFacs.Rows.Count
        End If
      End If

      Select Case ReportOutput

        Case TipusEnviamentDocumentEnum.Pdf, TipusEnviamentDocumentEnum.Email
          rpt.Destinacio = csRpt.ReportDestinationEnum.PDF
          rpt.pdfShowSaveDialog = False
          rpt.pdfPathAndFileName = String.Empty
          rpt.pdfDirectori = My.Settings.OutputDirPDF
          rpt.pdfNomFitxer = String.Format("Factura {0}-{1} - {2}.PDF", CInt(tbFacs.Rows(0)("fc_any")), CInt(tbFacs.Rows(0)("fc_Numero")), CNull(tbFacs.Rows(0)("fc_nifcli")).Trim)

        Case TipusEnviamentDocumentEnum.Impressora
          rpt.Destinacio = csRpt.ReportDestinationEnum.Printer
          rpt.ShowPrintDialog = False
          rpt.SetDefaultPrinter()

      End Select

      For Each r As DataRow In tbFacs.Rows

        rpt.OrigenDades = origenDades
        rpt.SerieFactura = r("fc_Serie").ToString
        rpt.AnyFactura = r("fc_any").ToString
        rpt.NumeroFactura = CNull(r("fc_numero"), 0)

        NumFactura = r("fc_any").ToString + "-" + r("fc_numero").ToString

        rpt.Print()

        If ReportOutput = TipusEnviamentDocumentEnum.Email Then

          mail = New csFaxMail

          mail.fmDestination = FaxEmailFormActorEmum.actorEmail

          eMail = dbaCli.Lookup(CNull(r("fc_codcli")), "fc_email").ToString

          mail.fmSubject = rpt.pdfNomFitxer

          If IsNullOrEmptyValue(eMail) Then
            ' Si no te correo l'enviem al administrador (Enric)
            mail.fmMailTo = My.Settings.mailMailUserSMTP.Trim
            mail.fmSubject = "SENSE MAIL: Factura " + NumFactura
          Else
            mail.fmMailTo = eMail
          End If

          mail.fmAttachment = rpt.pdfPathAndFileName

          mail.fmShowForm = False
          mail.fmCanAttachFiles = False

          mail.fmDeleteSentFiles = True

          mail.fmMailFeedBack = My.Settings.mailMailFeedback
          mail.fmMailFrom = My.Settings.mailMailUserSMTP
          mail.fmMailReplyTo = Nothing
          mail.fmSmtpLogin = My.Settings.mailUserSMTP
          mail.fmSmtpPassword = My.Settings.mailPasswordSMTP
          mail.fmSmtpServer = My.Settings.mailServerSMTP
          mail.fmBody = String.Format("Estimat col·laborador: {0}{0}" + _
                                      "Us adjuntem document PDF de la còpia de la factura so·licitada.{0}{0}" + _
                                      "Sense cap altre particular, rebeu una cordial salutació{0}{0}" + _
                                      "{1}{0}" + _
                                      "{2}{0}" + _
                                      "{3}", vbCr, My.Settings.PersonaQueSaluda, My.Settings.NomEmpresa, My.Settings.emailPersonaQueSaluda)

          mail.Send()

          If Not mail.fmMailSentOK Then
            Utils.MailNotification( _
              My.Settings.mailServerSMTP, _
              My.Settings.mailUserSMTP, _
              My.Settings.mailPasswordSMTP, _
              My.Settings.mailMailUserSMTP, _
              My.Settings.mailMailUserSMTP.Trim, _
              "ERROR enviament correo", _
              "ERROR MAIL: EXP - Factura " + NumFactura)
          End If

          mail = Nothing

        End If

      Next

    Catch ex As Exception

      DebugLog(AppData.Debug, Date.Now.ToString("dd/MM/yyyy HH:mm:ss") + " R90FAC0002A_Factura " + ex.Message)

    End Try

  End Sub

End Class
