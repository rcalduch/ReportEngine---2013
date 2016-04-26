Imports csUtils.Utils
Imports System.Windows.Forms


Public Class csFaxEmailForm
  Inherits System.Windows.Forms.Form

  Private mActor As FaxEmailFormActorEmum
  Property Actor() As FaxEmailFormActorEmum
    Get
      Return mActor
    End Get
    Set(ByVal value As FaxEmailFormActorEmum)
      mActor = value
      Select Case value
        Case FaxEmailFormActorEmum.actorFax
          Me.Text = "Enviament de fax"
          lblDestinatari.Visible = True
          lblDestinatari.Text = "Número de fax:"
          lblRetol.Text = "Enviament de FAX"
          lblDe.Visible = True
          txtDe.Visible = True
          lblAlaAtencio.Visible = True
          txtAlaAtencio.Visible = True
          lblAttachment.Visible = False
          txtAttachment.Visible = False

        Case FaxEmailFormActorEmum.actorEmail
          Me.Text = "Enviament de correu electrònic"
          lblDestinatari.Visible = True
          lblDestinatari.Text = "Destinatari:"
          lblRetol.Text = "Enviament de e-mail"
          lblDe.Visible = False
          txtDe.Visible = False
          lblAlaAtencio.Visible = False
          txtAlaAtencio.Visible = False
          lblAttachment.Visible = True
          txtAttachment.Visible = True

      End Select
      Me.cmdEnviar.Enabled = False
    End Set
  End Property

  Property MailTo() As String
    Get
      Return txtSendTo.Value.ToString.Trim.Replace(",", ";")
    End Get
    Set(ByVal value As String)
      txtSendTo.Value = value
    End Set
  End Property

  Property NumerosDeFax() As String
    Get
      Return txtSendTo.Value.ToString.Trim.Replace(",", ";")
    End Get
    Set(ByVal value As String)
      txtSendTo.Value = value
    End Set
  End Property

  Property Assumpte() As String
    Get
      Return txtAssumpte.Text
    End Get
    Set(ByVal value As String)
      txtAssumpte.Text = value
    End Set
  End Property

  Property Attachment() As String
    Get
      Return txtAttachment.Value.ToString.Trim.Replace(","c, ";"c)
    End Get
    Set(ByVal value As String)
      txtAttachment.Value = value
    End Set
  End Property

  Property CanAttachFiles() As Boolean
    Get
      Return txtAttachment.Enabled
    End Get
    Set(ByVal value As Boolean)
      txtAttachment.Enabled = value
    End Set
  End Property

  Property De() As String
    Get
      Return txtDe.Text
    End Get
    Set(ByVal value As String)
      txtDe.Text = value
    End Set
  End Property

  Property AlaAtencio() As String
    Get
      Return txtAlaAtencio.Text
    End Get
    Set(ByVal value As String)
      txtAlaAtencio.Text = value
    End Set
  End Property

  Private mContactes As DataTable
  Property Contactes() As DataTable
    Get
      Return mContactes
    End Get
    Set(ByVal value As DataTable)
      mContactes = value
    End Set
  End Property

  Property Missatge() As String
    Get
      Return txtMissatge.Text
    End Get
    Set(ByVal value As String)
      txtMissatge.Text = value
    End Set
  End Property

  Private mOpenFileInitialDir As String
  Property OpenFileInitialDir() As String
    Get
      If String.IsNullOrEmpty(mOpenFileInitialDir) Then
        Return My.Computer.FileSystem.SpecialDirectories.MyDocuments
      Else
        Return mOpenFileInitialDir
      End If
    End Get
    Set(ByVal value As String)
      mOpenFileInitialDir = value
    End Set
  End Property

  Private mOpenFileDefaultExt As String
  Property OpenFileDefaultExt() As String
    Get
      If String.IsNullOrEmpty(mOpenFileDefaultExt) Then
        Return "*.*"
      Else
        Return mOpenFileDefaultExt
      End If
    End Get
    Set(ByVal value As String)
      mOpenFileDefaultExt = value
    End Set
  End Property

  Private mOpenFileFilter As String
  Property OpenFileFilter() As String
    Get
      If String.IsNullOrEmpty(mOpenFileFilter) Then
        Return "Fitxers pdf (*.pdf)|*.pdf|Tots els fitxers (*.*)|*.*"
      Else
        Return mOpenFileFilter
      End If
    End Get
    Set(ByVal value As String)
      mOpenFileFilter = value
    End Set
  End Property

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.DialogResult = Windows.Forms.DialogResult.Cancel
  End Sub

  Private Sub cmdEnviar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEnviar.Click
    Me.DialogResult = Windows.Forms.DialogResult.OK
  End Sub

  Private Sub txtAttachment_ModalButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAttachment.ModalButtonClick

    Dim dlg As New System.Windows.Forms.OpenFileDialog
    dlg.InitialDirectory = OpenFileInitialDir
    dlg.Multiselect = True
    dlg.CheckFileExists = True
    dlg.ValidateNames = True
    dlg.DefaultExt = OpenFileDefaultExt
    dlg.Filter = OpenFileFilter
    If dlg.ShowDialog = Windows.Forms.DialogResult.OK Then
      For Each f As String In dlg.FileNames
        If Not String.IsNullOrEmpty(txtAttachment.Text) Then
          txtAttachment.Text += ";"
        End If
        txtAttachment.Text += f
      Next
    End If
  End Sub

  Private Sub txtSendTo_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSendTo.Enter
    If lvContactes.Visible Then
      lvContactes.Visible = False
      lvContactes.Tag = 0
    End If
  End Sub

  Private Sub txtTextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles txtSendTo.TextChanged
    Select Case Actor
      Case FaxEmailFormActorEmum.actorFax
        If txtSendTo.Text.Length >= 9 Then
          cmdEnviar.Enabled = True
        Else
          cmdEnviar.Enabled = False
        End If
      Case FaxEmailFormActorEmum.actorEmail
        If txtSendTo.Text.IndexOf("@"c) >= 0 And txtSendTo.Text.IndexOf("."c) >= 0 Then
          cmdEnviar.Enabled = True
        Else
          cmdEnviar.Enabled = False
        End If
    End Select
  End Sub

  Private Sub txtSendTo_ModalButtonClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSendTo.ModalButtonClick
    If Utils.CNull(lvContactes.Tag, 0) = 1 Then
      lvContactes.Tag = 0
      lvContactes.Visible = False
      txtSendTo.Focus()
    Else
      lvContactes.Tag = 1
      lvContactes.Visible = True
      lvContactes.Focus()
    End If
  End Sub

  Private Sub csFaxEmailForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    Dim ds As New DataSet
    Dim lv As listviewitem

    lvContactes.Visible = False

    lvContactes.Location = New System.Drawing.Point(txtSendTo.Location.X, txtSendTo.Location.Y + txtSendTo.Height)
    lvContactes.Height = 150
    lvContactes.Width = 500
   
    lvContactes.Columns.Clear()
    lvContactes.Columns.Add("Contacte", 125, HorizontalAlignment.Left)
    lvContactes.Columns.Add("Departament", 125, HorizontalAlignment.Left)
    lvContactes.Columns.Add("Telèfon", 100, HorizontalAlignment.Left)
    If Actor = FaxEmailFormActorEmum.actorFax Then
      lvContactes.Columns.Add("Fax", 150, HorizontalAlignment.Left)
    Else
      lvContactes.Columns.Add("e-mail", 150, HorizontalAlignment.Left)
    End If
    lvContactes.FullRowSelect = True
    lvContactes.View = View.Details

    lvContactes.Items.Clear()

    If Not Contactes Is Nothing Then

      For Each r As DataRow In Contactes.Rows
        lv = lvContactes.Items.Add(CNull(r("T2512_Contacte")).Trim)
        lv.SubItems.Add(CNull(r("T2512_Departament")))
        lv.SubItems.Add(CNull(r("T2512_Telefon")))
        If Actor = FaxEmailFormActorEmum.actorFax Then
          lv.SubItems.Add(CNull(r("T2512_Fax")))
        Else
          lv.SubItems.Add(CNull(r("T2512_Email")))
        End If

      Next

    End If

  End Sub

  Private Sub lvContactes_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles lvContactes.DoubleClick
    Dim Value As String
    lvContactes.MultiSelect = False

    For Each it As ListViewItem In lvContactes.Items
      If it.Selected Then
        Value = it.SubItems(3).Text
        If Not String.IsNullOrEmpty(Value) Then
          If My.Computer.Keyboard.CtrlKeyDown Then
            If String.IsNullOrEmpty(txtSendTo.Text) Then
              txtSendTo.Text = Value
            Else
              txtSendTo.Text += ", " + Value
            End If
          Else
            txtSendTo.Text = Value
          End If
        End If

      End If
    Next

    txtSendTo.Focus()
    lvContactes.Visible = False
    lvContactes.Tag = 0
  End Sub


End Class
