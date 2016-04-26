<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class csFaxEmailForm
  Inherits System.Windows.Forms.Form

  'Form overrides dispose to clean up the component list.
  <System.Diagnostics.DebuggerNonUserCode()> _
  Protected Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing AndAlso components IsNot Nothing Then
      components.Dispose()
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> _
  Private Sub InitializeComponent()
    Dim ListViewItem1 As System.Windows.Forms.ListViewItem = New System.Windows.Forms.ListViewItem(New String() {"", "", "", "", ""}, -1)
    Me.lblDestinatari = New System.Windows.Forms.Label()
    Me.pnlFax = New System.Windows.Forms.Panel()
    Me.lblRetol = New System.Windows.Forms.Label()
    Me.lblAssumpte = New System.Windows.Forms.Label()
    Me.Label4 = New System.Windows.Forms.Label()
    Me.txtAssumpte = New System.Windows.Forms.TextBox()
    Me.txtMissatge = New System.Windows.Forms.TextBox()
    Me.cmdEnviar = New System.Windows.Forms.Button()
    Me.cmdCancel = New System.Windows.Forms.Button()
    Me.txtAttachment = New C1.Win.C1Input.C1DropDownControl()
    Me.lblAttachment = New System.Windows.Forms.Label()
    Me.txtSendTo = New C1.Win.C1Input.C1DropDownControl()
    Me.txtDe = New System.Windows.Forms.TextBox()
    Me.txtAlaAtencio = New System.Windows.Forms.TextBox()
    Me.lblDe = New System.Windows.Forms.Label()
    Me.lblAlaAtencio = New System.Windows.Forms.Label()
    Me.lvContactes = New System.Windows.Forms.ListView()
    Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.ColumnHeader5 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
    Me.pnlFax.SuspendLayout()
    CType(Me.txtAttachment, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.txtSendTo, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.SuspendLayout()
    '
    'lblDestinatari
    '
    Me.lblDestinatari.AutoSize = True
    Me.lblDestinatari.Location = New System.Drawing.Point(18, 66)
    Me.lblDestinatari.Name = "lblDestinatari"
    Me.lblDestinatari.Size = New System.Drawing.Size(79, 13)
    Me.lblDestinatari.TabIndex = 0
    Me.lblDestinatari.Text = "Número de fax:"
    '
    'pnlFax
    '
    Me.pnlFax.BackColor = System.Drawing.SystemColors.ActiveCaptionText
    Me.pnlFax.Controls.Add(Me.lblRetol)
    Me.pnlFax.Dock = System.Windows.Forms.DockStyle.Top
    Me.pnlFax.Location = New System.Drawing.Point(0, 0)
    Me.pnlFax.Name = "pnlFax"
    Me.pnlFax.Size = New System.Drawing.Size(640, 47)
    Me.pnlFax.TabIndex = 1
    '
    'lblRetol
    '
    Me.lblRetol.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.lblRetol.Dock = System.Windows.Forms.DockStyle.Fill
    Me.lblRetol.Font = New System.Drawing.Font("Microsoft Sans Serif", 22.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblRetol.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
    Me.lblRetol.Location = New System.Drawing.Point(0, 0)
    Me.lblRetol.Name = "lblRetol"
    Me.lblRetol.Size = New System.Drawing.Size(640, 47)
    Me.lblRetol.TabIndex = 0
    Me.lblRetol.Text = "Enviament de FAX "
    Me.lblRetol.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblAssumpte
    '
    Me.lblAssumpte.AutoSize = True
    Me.lblAssumpte.Location = New System.Drawing.Point(18, 102)
    Me.lblAssumpte.Name = "lblAssumpte"
    Me.lblAssumpte.Size = New System.Drawing.Size(56, 13)
    Me.lblAssumpte.TabIndex = 2
    Me.lblAssumpte.Text = "Assumpte:"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(18, 174)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(49, 13)
    Me.Label4.TabIndex = 3
    Me.Label4.Text = "Missatge"
    '
    'txtAssumpte
    '
    Me.txtAssumpte.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtAssumpte.Location = New System.Drawing.Point(102, 96)
    Me.txtAssumpte.Name = "txtAssumpte"
    Me.txtAssumpte.Size = New System.Drawing.Size(516, 20)
    Me.txtAssumpte.TabIndex = 2
    '
    'txtMissatge
    '
    Me.txtMissatge.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtMissatge.Location = New System.Drawing.Point(102, 168)
    Me.txtMissatge.Multiline = True
    Me.txtMissatge.Name = "txtMissatge"
    Me.txtMissatge.Size = New System.Drawing.Size(516, 168)
    Me.txtMissatge.TabIndex = 6
    '
    'cmdEnviar
    '
    Me.cmdEnviar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdEnviar.Location = New System.Drawing.Point(468, 342)
    Me.cmdEnviar.Name = "cmdEnviar"
    Me.cmdEnviar.Size = New System.Drawing.Size(75, 23)
    Me.cmdEnviar.TabIndex = 7
    Me.cmdEnviar.Text = "Enviar"
    Me.cmdEnviar.UseVisualStyleBackColor = True
    '
    'cmdCancel
    '
    Me.cmdCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancel.Location = New System.Drawing.Point(546, 342)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancel.TabIndex = 8
    Me.cmdCancel.Text = "Cancelar"
    Me.cmdCancel.UseVisualStyleBackColor = True
    '
    'txtAttachment
    '
    Me.txtAttachment.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtAttachment.Location = New System.Drawing.Point(102, 132)
    Me.txtAttachment.Name = "txtAttachment"
    Me.txtAttachment.Size = New System.Drawing.Size(516, 20)
    Me.txtAttachment.TabIndex = 5
    Me.txtAttachment.Tag = Nothing
    Me.txtAttachment.Value = ""
    Me.txtAttachment.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.Modal
    '
    'lblAttachment
    '
    Me.lblAttachment.AutoSize = True
    Me.lblAttachment.Location = New System.Drawing.Point(18, 138)
    Me.lblAttachment.Name = "lblAttachment"
    Me.lblAttachment.Size = New System.Drawing.Size(77, 13)
    Me.lblAttachment.TabIndex = 10
    Me.lblAttachment.Text = "Fitxers adjunts:"
    '
    'txtSendTo
    '
    Me.txtSendTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtSendTo.Location = New System.Drawing.Point(102, 66)
    Me.txtSendTo.Name = "txtSendTo"
    Me.txtSendTo.Size = New System.Drawing.Size(516, 20)
    Me.txtSendTo.TabIndex = 0
    Me.txtSendTo.Tag = Nothing
    Me.txtSendTo.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.Modal
    '
    'txtDe
    '
    Me.txtDe.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtDe.Location = New System.Drawing.Point(102, 132)
    Me.txtDe.Name = "txtDe"
    Me.txtDe.Size = New System.Drawing.Size(204, 20)
    Me.txtDe.TabIndex = 3
    '
    'txtAlaAtencio
    '
    Me.txtAlaAtencio.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtAlaAtencio.Location = New System.Drawing.Point(408, 132)
    Me.txtAlaAtencio.Name = "txtAlaAtencio"
    Me.txtAlaAtencio.Size = New System.Drawing.Size(210, 20)
    Me.txtAlaAtencio.TabIndex = 4
    '
    'lblDe
    '
    Me.lblDe.AutoSize = True
    Me.lblDe.Location = New System.Drawing.Point(18, 138)
    Me.lblDe.Name = "lblDe"
    Me.lblDe.Size = New System.Drawing.Size(24, 13)
    Me.lblDe.TabIndex = 2
    Me.lblDe.Text = "De:"
    '
    'lblAlaAtencio
    '
    Me.lblAlaAtencio.AutoSize = True
    Me.lblAlaAtencio.Location = New System.Drawing.Point(342, 138)
    Me.lblAlaAtencio.Name = "lblAlaAtencio"
    Me.lblAlaAtencio.Size = New System.Drawing.Size(63, 13)
    Me.lblAlaAtencio.TabIndex = 2
    Me.lblAlaAtencio.Text = "A la atenció"
    '
    'lvContactes
    '
    Me.lvContactes.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4, Me.ColumnHeader5})
    Me.lvContactes.Items.AddRange(New System.Windows.Forms.ListViewItem() {ListViewItem1})
    Me.lvContactes.Location = New System.Drawing.Point(102, 48)
    Me.lvContactes.Name = "lvContactes"
    Me.lvContactes.Size = New System.Drawing.Size(156, 18)
    Me.lvContactes.TabIndex = 11
    Me.lvContactes.UseCompatibleStateImageBehavior = False
    '
    'csFaxEmailForm
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.ClientSize = New System.Drawing.Size(640, 373)
    Me.Controls.Add(Me.lvContactes)
    Me.Controls.Add(Me.txtSendTo)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdEnviar)
    Me.Controls.Add(Me.txtMissatge)
    Me.Controls.Add(Me.txtAssumpte)
    Me.Controls.Add(Me.Label4)
    Me.Controls.Add(Me.lblDe)
    Me.Controls.Add(Me.lblAssumpte)
    Me.Controls.Add(Me.pnlFax)
    Me.Controls.Add(Me.lblDestinatari)
    Me.Controls.Add(Me.lblAttachment)
    Me.Controls.Add(Me.lblAlaAtencio)
    Me.Controls.Add(Me.txtAlaAtencio)
    Me.Controls.Add(Me.txtDe)
    Me.Controls.Add(Me.txtAttachment)
    Me.Name = "csFaxEmailForm"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Enviament"
    Me.TopMost = True
    Me.pnlFax.ResumeLayout(False)
    CType(Me.txtAttachment, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.txtSendTo, System.ComponentModel.ISupportInitialize).EndInit()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub
  Friend WithEvents lblDestinatari As System.Windows.Forms.Label
  Friend WithEvents pnlFax As System.Windows.Forms.Panel
  Friend WithEvents lblRetol As System.Windows.Forms.Label
  Friend WithEvents lblAssumpte As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents txtAssumpte As System.Windows.Forms.TextBox
  Friend WithEvents txtMissatge As System.Windows.Forms.TextBox
  Friend WithEvents cmdEnviar As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents txtAttachment As C1.Win.C1Input.C1DropDownControl
  Friend WithEvents lblAttachment As System.Windows.Forms.Label
  Friend WithEvents txtSendTo As C1.Win.C1Input.C1DropDownControl
  Friend WithEvents txtDe As System.Windows.Forms.TextBox
  Friend WithEvents txtAlaAtencio As System.Windows.Forms.TextBox
  Friend WithEvents lblDe As System.Windows.Forms.Label
  Friend WithEvents lblAlaAtencio As System.Windows.Forms.Label
  Friend WithEvents lvContactes As System.Windows.Forms.ListView
  Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
  Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
End Class
