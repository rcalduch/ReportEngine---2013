Imports System.Net.Mail
Imports System.Drawing

Public Class csFaxMail

  Public fmDestination As FaxEmailFormActorEmum
  Public fmContactes As DataTable
  Public fmCanAttachFiles As Boolean
  Public fmAttachment As String
  Public fmOFDFilter As String
  Public fmOFDInitialDir As String
  Public fmOFDDefaultExtension As String
  Public fmFaxLogo As Image
  Public fmFaxCustomerLogo As Image
  Public fmFaxLogoFile As String
  Public fmFaxCustomerLogoFile As String
  Public fmFaxNomUsuari As String
  Public fmFaxNumero As String
  Public fmFaxAlaAtencio As String
  Public fmFaxPaginesDocument As Integer
  Public fmMailFrom As String
  Public fmMailTo As String
  Public fmMailCC As String
  Public fmMailReplyTo As String
  Public fmMailFeedBack As String
  Public fmMailAccountFax As String
  Public fmMailSentOK As Boolean
  Public fmSubject As String
  Public fmBody As String
  Public fmSmtpServer As String
  Public fmSmtpLogin As String
  Public fmSmtpPassword As String
  Public fmShowForm As Boolean
  Public fmDeleteSentFiles As Boolean
  Public fmShowSentResult As Boolean

  Private ff As csFaxEmailForm
  Private email As MailMessage
  Private Att As Attachment
  Private objSmtp As Net.Mail.SmtpClient

  Private PortadaFaxFileName As String
  Private PortadaFax As csPortadaFax

  Private Function MustShowForm() As Boolean
    Return False
  End Function

  Public Sub AddContacte(ByVal Contacte As Object, ByVal Departament As Object, ByVal Telefon As Object, ByVal Fax As Object, ByVal eMail As Object)
    Dim rowContacte As DataRow
    If fmContactes Is Nothing Then
      fmContactes = New DataTable("Contactes")
      fmContactes.Columns.Add(New DataColumn("T2512_Contacte", GetType(String)))
      fmContactes.Columns.Add(New DataColumn("T2512_Departament", GetType(String)))
      fmContactes.Columns.Add(New DataColumn("T2512_Telefon", GetType(String)))
      fmContactes.Columns.Add(New DataColumn("T2512_Fax", GetType(String)))
      fmContactes.Columns.Add(New DataColumn("T2512_Email", GetType(String)))
    End If
    rowContacte = fmContactes.NewRow
    rowContacte("T2512_Contacte") = Contacte
    rowContacte("T2512_Departament") = Departament
    rowContacte("T2512_Telefon") = Telefon
    rowContacte("T2512_Fax") = Fax
    rowContacte("T2512_Email") = eMail
    fmContactes.Rows.Add(rowContacte)
  End Sub

  Public Overridable Sub Send()

    If fmShowForm Or MustShowForm() Then

      ff = New csFaxEmailForm

      ff.Actor = fmDestination
      ff.Contactes = fmContactes

      If fmDestination = FaxEmailFormActorEmum.actorEmail Then
        ff.MailTo = fmMailTo
      Else
        ff.NumerosDeFax = fmFaxNumero
        ff.De = fmFaxNomUsuari
        ff.AlaAtencio = fmFaxAlaAtencio
      End If

      ff.Assumpte = fmSubject
      ff.Missatge = fmBody

      ff.Attachment = fmAttachment
      ff.CanAttachFiles = fmCanAttachFiles
      ff.OpenFileFilter = fmOFDFilter
      ff.OpenFileInitialDir = fmOFDInitialDir
      ff.OpenFileDefaultExt = fmOFDDefaultExtension

      If ff.ShowDialog = Windows.Forms.DialogResult.Cancel Then
        ff.Dispose()
        Return
      End If

      fmAttachment = ff.Attachment
      fmBody = ff.Missatge
      fmSubject = ff.Assumpte
      fmFaxAlaAtencio = ff.AlaAtencio
      fmFaxNomUsuari = ff.De
      fmFaxNumero = ff.NumerosDeFax
      fmMailTo = ff.MailTo

      ff.Dispose()

    End If

    email = New System.Net.Mail.MailMessage

    Try

      email.From = New Net.Mail.MailAddress(Me.fmMailFrom)
      If Not String.IsNullOrEmpty(fmMailReplyTo) Then
        email.ReplyTo = New Net.Mail.MailAddress(fmMailReplyTo)
      End If

      If fmDestination = FaxEmailFormActorEmum.actorFax Then

        If fmFaxLogo Is Nothing Then
          If Not String.IsNullOrEmpty(fmFaxLogoFile) Then
            If IO.File.Exists(fmFaxLogoFile) Then
              fmFaxLogo = System.Drawing.Image.FromFile(fmFaxLogoFile)
            End If
          End If
        End If

        If fmFaxCustomerLogo Is Nothing Then
          If Not String.IsNullOrEmpty(fmFaxCustomerLogoFile) Then
            If IO.File.Exists(fmFaxCustomerLogoFile) Then
              fmFaxCustomerLogo = System.Drawing.Image.FromFile(fmFaxCustomerLogoFile)
            End If
          End If
        End If

        PortadaFax = New csPortadaFax

        PortadaFax.fmFaxCustomerLogo = fmFaxCustomerLogo
        PortadaFax.fmFaxLogo = fmFaxLogo
        PortadaFax.fmFaxNomUsuari = fmFaxNomUsuari
        PortadaFax.fmFaxAlaAtencio = fmFaxAlaAtencio
        PortadaFax.fmFaxNumero = fmFaxNumero
        PortadaFax.fmSubject = fmSubject
        PortadaFax.fmBody = fmBody
        PortadaFax.fmFaxPaginesDocument = fmFaxPaginesDocument

        PortadaFaxFileName = PortadaFax.PrintPortadaFax

        PortadaFax = Nothing

        If IO.File.Exists(PortadaFaxFileName) Then
          Att = New Attachment(PortadaFaxFileName)
          email.Attachments.Add(Att)
        End If

      End If

      For Each fileToSend As String In fmAttachment.Split(";"c)
        If Not String.IsNullOrEmpty(fileToSend) Then
          Att = New Attachment(fileToSend)
          email.Attachments.Add(Att)
        End If
      Next

      If fmDestination = FaxEmailFormActorEmum.actorEmail Then
        For Each mailToSend As String In fmMailTo.Split(";"c)
          If Not String.IsNullOrEmpty(mailToSend) Then
            email.To.Add(mailToSend)
          End If
        Next
      Else
        email.To.Add(fmMailAccountFax)
        For Each dest As String In fmFaxNumero.Split(";"c)
          email.To.Add(dest + fmMailAccountFax.Substring(fmMailAccountFax.IndexOf("@")))
        Next
      End If

      If Not String.IsNullOrEmpty(fmMailFeedBack) Then
        email.Bcc.Add(fmMailFeedBack)
      End If

      email.Subject = fmSubject
      email.Body = fmBody
      email.IsBodyHtml = False

      objSmtp = New Net.Mail.SmtpClient(fmSmtpServer)

      objSmtp.Credentials = New System.Net.NetworkCredential(fmSmtpLogin, fmSmtpPassword)
      objSmtp.Send(email)

      ' Clean up
      For Each Att As Attachment In email.Attachments
        Att.ContentStream.Close()
      Next
      email.Attachments.Clear()
      email.Dispose()

      email = Nothing
      objSmtp = Nothing

      If fmShowSentResult Then
        If fmDestination = FaxEmailFormActorEmum.actorEmail Then
          MsgBox("Correo enviat correctament.", MsgBoxStyle.Information, "Enviament correo")
        Else
          MsgBox("Fax enviat correctament.", MsgBoxStyle.Information, "Enviamnet fax")
        End If
      End If

      fmMailSentOK = True

    Catch ex As Exception

      If fmShowSentResult Then
        If fmDestination = FaxEmailFormActorEmum.actorEmail Then
          MsgBox("S'ha produït un error en l'enviament del correo. " + Chr(13) + ex.Message, MsgBoxStyle.Critical, "ERROR")
        Else
          MsgBox("S'ha produït un error en l'enviament del fax. " + Chr(13) + ex.Message, MsgBoxStyle.Critical, "ERROR")
        End If
      End If

      fmMailSentOK = False

    Finally

      ' Ja em enviat, borrem el fitxer enviat
      If fmDeleteSentFiles Then
        For Each fileSent As String In fmAttachment.Split(";"c)
          If Not String.IsNullOrEmpty(fileSent) Then
            If IO.File.Exists(fileSent) Then
              Try
                IO.File.Delete(fileSent)
              Catch ex As Exception
                'MsgBox(ex.ToString)
              End Try
            End If
          End If
        Next
      End If

    End Try

  End Sub


End Class
