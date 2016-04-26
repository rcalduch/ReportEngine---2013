Imports csUtils
Imports System.Type

Public Class frmSerialDialog
  Inherits System.Windows.Forms.Form

  Private mSetDefaultValuesOnStartUp As Boolean = True
  Private mSerialSettings As String

  Property SerialSettings() As String
    ' "Port=COM1;Velocitat=9600;Paritat=Ninguna;BitsDades=8;BitsStop=1;DTR=True;RTS=True;DsrDtr=False;XonXoff=False;CtsRts=True;Rs485=False"
    Get
      Return String.Format("Port={0};Velocitat={1};Paritat={2};BitsDades={3};BitsStop={4};DTR={5};RTS={6};DsrDtr={7};XonXoff={8};CtsRts={9};Rs485={10}", cboPorts.SelectedItem, cboVelocitat.SelectedItem, cboParitat.SelectedItem, cboBitsDades.SelectedItem, cboBitsParada.SelectedItem, chkDTR.Checked, chkRTS.Checked, chkDsrDtr.Checked, chkXonXoff.Checked, chkCtsRts.Checked, chkRs485.Checked)
    End Get
    Set(ByVal Value As String)
      If Value Is Nothing Then
        Return
      End If
      Dim items() As String

      items = Split(Value, ";")

      If items.Length <> 11 Then
        Return
      End If

      cboPorts.SelectedItem = Split(items(0), "=")(1)
      cboVelocitat.SelectedItem = Split(items(1), "=")(1)
      cboParitat.SelectedItem = Split(items(2), "=")(1)
      cboBitsDades.SelectedItem = Split(items(3), "=")(1)
      cboBitsParada.SelectedItem = Split(items(4), "=")(1)

      chkDTR.Checked = CBool(Split(items(5), "=")(1))
      chkRTS.Checked = CBool(Split(items(6), "=")(1))
      chkDsrDtr.Checked = CBool(Split(items(7), "=")(1))
      chkXonXoff.Checked = CBool(Split(items(8), "=")(1))
      chkCtsRts.Checked = CBool(Split(items(9), "=")(1))
      chkRs485.Checked = CBool(Split(items(10), "=")(1))

      mSetDefaultValuesOnStartUp = False
    End Set
  End Property

#Region " Código generado por el Diseñador de Windows Forms "

  Public Sub New()
    MyBase.New()

    'El Diseñador de Windows Forms requiere esta llamada.
    InitializeComponent()

    'Agregar cualquier inicialización después de la llamada a InitializeComponent()

  End Sub

  'Form reemplaza a Dispose para limpiar la lista de componentes.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Requerido por el Diseñador de Windows Forms
  Private components As System.ComponentModel.IContainer

  'NOTA: el Diseñador de Windows Forms requiere el siguiente procedimiento
  'Puede modificarse utilizando el Diseñador de Windows Forms. 
  'No lo modifique con el editor de código.
  Friend WithEvents lblTitol As System.Windows.Forms.Label
  Friend WithEvents StatusBar As System.Windows.Forms.StatusBar
  Friend WithEvents sbpInfo As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpFormName As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpVersion As System.Windows.Forms.StatusBarPanel
  Friend WithEvents sbpDummy As System.Windows.Forms.StatusBarPanel
  Friend WithEvents HelpProvider As System.Windows.Forms.HelpProvider
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents cboPorts As System.Windows.Forms.ComboBox
  Friend WithEvents cboVelocitat As System.Windows.Forms.ComboBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents cboBitsDades As System.Windows.Forms.ComboBox
  Friend WithEvents cboParitat As System.Windows.Forms.ComboBox
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents cboBitsParada As System.Windows.Forms.ComboBox
  Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
  Friend WithEvents chkDTR As System.Windows.Forms.CheckBox
  Friend WithEvents chkRTS As System.Windows.Forms.CheckBox
  Friend WithEvents chkDsrDtr As System.Windows.Forms.CheckBox
  Friend WithEvents chkXonXoff As System.Windows.Forms.CheckBox
  Friend WithEvents chkCtsRts As System.Windows.Forms.CheckBox
  Friend WithEvents chkRs485 As System.Windows.Forms.CheckBox
  Friend WithEvents cmdDefault As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents cmdCancelar As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.lblTitol = New System.Windows.Forms.Label
    Me.StatusBar = New System.Windows.Forms.StatusBar
    Me.sbpInfo = New System.Windows.Forms.StatusBarPanel
    Me.sbpFormName = New System.Windows.Forms.StatusBarPanel
    Me.sbpVersion = New System.Windows.Forms.StatusBarPanel
    Me.sbpDummy = New System.Windows.Forms.StatusBarPanel
    Me.HelpProvider = New System.Windows.Forms.HelpProvider
    Me.cboPorts = New System.Windows.Forms.ComboBox
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label1 = New System.Windows.Forms.Label
    Me.cboVelocitat = New System.Windows.Forms.ComboBox
    Me.GroupBox2 = New System.Windows.Forms.GroupBox
    Me.Label5 = New System.Windows.Forms.Label
    Me.cboBitsParada = New System.Windows.Forms.ComboBox
    Me.Label3 = New System.Windows.Forms.Label
    Me.Label4 = New System.Windows.Forms.Label
    Me.cboBitsDades = New System.Windows.Forms.ComboBox
    Me.cboParitat = New System.Windows.Forms.ComboBox
    Me.GroupBox3 = New System.Windows.Forms.GroupBox
    Me.chkRs485 = New System.Windows.Forms.CheckBox
    Me.chkCtsRts = New System.Windows.Forms.CheckBox
    Me.chkXonXoff = New System.Windows.Forms.CheckBox
    Me.chkDsrDtr = New System.Windows.Forms.CheckBox
    Me.chkRTS = New System.Windows.Forms.CheckBox
    Me.chkDTR = New System.Windows.Forms.CheckBox
    Me.cmdDefault = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.cmdCancelar = New System.Windows.Forms.Button
    CType(Me.sbpInfo, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpFormName, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpVersion, System.ComponentModel.ISupportInitialize).BeginInit()
    CType(Me.sbpDummy, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.GroupBox1.SuspendLayout()
    Me.GroupBox2.SuspendLayout()
    Me.GroupBox3.SuspendLayout()
    Me.SuspendLayout()
    '
    'lblTitol
    '
    Me.lblTitol.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
    Me.lblTitol.Dock = System.Windows.Forms.DockStyle.Top
    Me.lblTitol.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lblTitol.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
    Me.lblTitol.Location = New System.Drawing.Point(0, 0)
    Me.lblTitol.Name = "lblTitol"
    Me.lblTitol.Size = New System.Drawing.Size(330, 32)
    Me.lblTitol.TabIndex = 22
    Me.lblTitol.Text = "  Paràmetres conexió sèrie"
    Me.lblTitol.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'StatusBar
    '
    Me.StatusBar.Location = New System.Drawing.Point(0, 464)
    Me.StatusBar.Name = "StatusBar"
    Me.StatusBar.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.sbpInfo, Me.sbpFormName, Me.sbpVersion, Me.sbpDummy})
    Me.StatusBar.ShowPanels = True
    Me.StatusBar.Size = New System.Drawing.Size(330, 22)
    Me.StatusBar.TabIndex = 20
    '
    'sbpInfo
    '
    Me.sbpInfo.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
    Me.sbpInfo.Name = "sbpInfo"
    Me.sbpInfo.Width = 185
    '
    'sbpFormName
    '
    Me.sbpFormName.Alignment = System.Windows.Forms.HorizontalAlignment.Center
    Me.sbpFormName.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpFormName.Name = "sbpFormName"
    Me.sbpFormName.Text = "pr_frmSerialDialog"
    Me.sbpFormName.Width = 108
    '
    'sbpVersion
    '
    Me.sbpVersion.Alignment = System.Windows.Forms.HorizontalAlignment.Center
    Me.sbpVersion.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpVersion.Name = "sbpVersion"
    Me.sbpVersion.Width = 10
    '
    'sbpDummy
    '
    Me.sbpDummy.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Contents
    Me.sbpDummy.Name = "sbpDummy"
    Me.sbpDummy.Width = 10
    '
    'cboPorts
    '
    Me.cboPorts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboPorts.Items.AddRange(New Object() {"COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM10", "COM11", "COM12", "COM13", "COM14", "COM15", "COM16", "COM17", "COM18", "COM19", "COM20", "COM21", "COM22", "COM23", "COM24", "COM25", "COM26", "COM27", "COM28", "COM29", "COM30", "COM31", "COM32", "COM33", "COM34", "COM35", "COM36", "COM37", "COM38", "COM39", "COM40", "COM41", "COM42", "COM43", "COM44", "COM45", "COM46", "COM47", "COM48", "COM49", "COM50", "COM51", "COM52", "COM53", "COM54", "COM55", "COM56", "COM57", "COM58", "COM59", "COM60", "COM61", "COM62", "COM63", "COM64", "COM65", "COM66", "COM67", "COM68", "COM69", "COM70", "COM71", "COM72", "COM73", "COM74", "COM75", "COM76", "COM77", "COM78", "COM79", "COM80", "COM81", "COM82", "COM83", "COM84", "COM85", "COM86", "COM87", "COM88", "COM89", "COM90", "COM91", "COM92", "COM93", "COM94", "COM95", "COM96", "COM97", "COM98", "COM99", "COM100", "COM101", "COM102", "COM103", "COM104", "COM105", "COM106", "COM107", "COM108", "COM109", "COM110", "COM111", "COM112", "COM113", "COM114", "COM115", "COM116", "COM117", "COM118", "COM119", "COM120", "COM121", "COM122", "COM123", "COM124", "COM125", "COM126", "COM127", "COM128", "COM129", "COM130", "COM131", "COM132", "COM133", "COM134", "COM135", "COM136", "COM137", "COM138", "COM139", "COM140", "COM141", "COM142", "COM143", "COM144", "COM145", "COM146", "COM147", "COM148", "COM149", "COM150", "COM151", "COM152", "COM153", "COM154", "COM155", "COM156", "COM157", "COM158", "COM159", "COM160", "COM161", "COM162", "COM163", "COM164", "COM165", "COM166", "COM167", "COM168", "COM169", "COM170", "COM171", "COM172", "COM173", "COM174", "COM175", "COM176", "COM177", "COM178", "COM179", "COM180", "COM181", "COM182", "COM183", "COM184", "COM185", "COM186", "COM187", "COM188", "COM189", "COM190", "COM191", "COM192", "COM193", "COM194", "COM195", "COM196", "COM197", "COM198", "COM199", "COM200", "COM201", "COM202", "COM203", "COM204", "COM205", "COM206", "COM207", "COM208", "COM209", "COM210", "COM211", "COM212", "COM213", "COM214", "COM215", "COM216", "COM217", "COM218", "COM219", "COM220", "COM221", "COM222", "COM223", "COM224", "COM225", "COM226", "COM227", "COM228", "COM229", "COM230", "COM231", "COM232", "COM233", "COM234", "COM235", "COM236", "COM237", "COM238", "COM239", "COM240", "COM241", "COM242", "COM243", "COM244", "COM245", "COM246", "COM247", "COM248", "COM249", "COM250", "COM251", "COM252", "COM253", "COM254", "COM255"})
    Me.cboPorts.Location = New System.Drawing.Point(112, 32)
    Me.cboPorts.Name = "cboPorts"
    Me.cboPorts.Size = New System.Drawing.Size(168, 21)
    Me.cboPorts.TabIndex = 0
    '
    'GroupBox1
    '
    Me.GroupBox1.Controls.Add(Me.Label2)
    Me.GroupBox1.Controls.Add(Me.Label1)
    Me.GroupBox1.Controls.Add(Me.cboVelocitat)
    Me.GroupBox1.Controls.Add(Me.cboPorts)
    Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.GroupBox1.Location = New System.Drawing.Point(16, 40)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(304, 104)
    Me.GroupBox1.TabIndex = 0
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Port comunicacions"
    '
    'Label2
    '
    Me.Label2.AutoSize = True
    Me.Label2.Location = New System.Drawing.Point(16, 72)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(48, 13)
    Me.Label2.TabIndex = 26
    Me.Label2.Text = "Velocitat"
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(16, 40)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(26, 13)
    Me.Label1.TabIndex = 25
    Me.Label1.Text = "Port"
    '
    'cboVelocitat
    '
    Me.cboVelocitat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboVelocitat.Items.AddRange(New Object() {"115200", "57600", "38400", "19200", "9600", "4800", "2400", "1200", "300"})
    Me.cboVelocitat.Location = New System.Drawing.Point(112, 64)
    Me.cboVelocitat.Name = "cboVelocitat"
    Me.cboVelocitat.Size = New System.Drawing.Size(168, 21)
    Me.cboVelocitat.TabIndex = 1
    '
    'GroupBox2
    '
    Me.GroupBox2.Controls.Add(Me.Label5)
    Me.GroupBox2.Controls.Add(Me.cboBitsParada)
    Me.GroupBox2.Controls.Add(Me.Label3)
    Me.GroupBox2.Controls.Add(Me.Label4)
    Me.GroupBox2.Controls.Add(Me.cboBitsDades)
    Me.GroupBox2.Controls.Add(Me.cboParitat)
    Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.GroupBox2.Location = New System.Drawing.Point(16, 152)
    Me.GroupBox2.Name = "GroupBox2"
    Me.GroupBox2.Size = New System.Drawing.Size(304, 136)
    Me.GroupBox2.TabIndex = 1
    Me.GroupBox2.TabStop = False
    Me.GroupBox2.Text = "Format dades"
    '
    'Label5
    '
    Me.Label5.AutoSize = True
    Me.Label5.Location = New System.Drawing.Point(16, 104)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(75, 13)
    Me.Label5.TabIndex = 28
    Me.Label5.Text = "Bits de parada"
    '
    'cboBitsParada
    '
    Me.cboBitsParada.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboBitsParada.Items.AddRange(New Object() {"0", "1", "2", "1.5"})
    Me.cboBitsParada.Location = New System.Drawing.Point(112, 96)
    Me.cboBitsParada.Name = "cboBitsParada"
    Me.cboBitsParada.Size = New System.Drawing.Size(168, 21)
    Me.cboBitsParada.TabIndex = 2
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(16, 72)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(71, 13)
    Me.Label3.TabIndex = 26
    Me.Label3.Text = "Bits de dades"
    '
    'Label4
    '
    Me.Label4.AutoSize = True
    Me.Label4.Location = New System.Drawing.Point(16, 40)
    Me.Label4.Name = "Label4"
    Me.Label4.Size = New System.Drawing.Size(37, 13)
    Me.Label4.TabIndex = 25
    Me.Label4.Text = "Paritat"
    '
    'cboBitsDades
    '
    Me.cboBitsDades.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboBitsDades.Items.AddRange(New Object() {"5", "6", "7", "8"})
    Me.cboBitsDades.Location = New System.Drawing.Point(112, 64)
    Me.cboBitsDades.Name = "cboBitsDades"
    Me.cboBitsDades.Size = New System.Drawing.Size(168, 21)
    Me.cboBitsDades.TabIndex = 1
    '
    'cboParitat
    '
    Me.cboParitat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboParitat.Items.AddRange(New Object() {"Ninguna", "Impar", "Par", "Marca", "Espai"})
    Me.cboParitat.Location = New System.Drawing.Point(112, 32)
    Me.cboParitat.Name = "cboParitat"
    Me.cboParitat.Size = New System.Drawing.Size(168, 21)
    Me.cboParitat.TabIndex = 0
    '
    'GroupBox3
    '
    Me.GroupBox3.Controls.Add(Me.chkRs485)
    Me.GroupBox3.Controls.Add(Me.chkCtsRts)
    Me.GroupBox3.Controls.Add(Me.chkXonXoff)
    Me.GroupBox3.Controls.Add(Me.chkDsrDtr)
    Me.GroupBox3.Controls.Add(Me.chkRTS)
    Me.GroupBox3.Controls.Add(Me.chkDTR)
    Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.GroupBox3.Location = New System.Drawing.Point(16, 296)
    Me.GroupBox3.Name = "GroupBox3"
    Me.GroupBox3.Size = New System.Drawing.Size(304, 128)
    Me.GroupBox3.TabIndex = 2
    Me.GroupBox3.TabStop = False
    Me.GroupBox3.Text = "Control de fluxe"
    '
    'chkRs485
    '
    Me.chkRs485.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkRs485.Location = New System.Drawing.Point(136, 96)
    Me.chkRs485.Name = "chkRs485"
    Me.chkRs485.Size = New System.Drawing.Size(144, 16)
    Me.chkRs485.TabIndex = 5
    Me.chkRs485.Text = "Activar control RS-485"
    '
    'chkCtsRts
    '
    Me.chkCtsRts.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkCtsRts.Location = New System.Drawing.Point(136, 72)
    Me.chkCtsRts.Name = "chkCtsRts"
    Me.chkCtsRts.Size = New System.Drawing.Size(152, 16)
    Me.chkCtsRts.TabIndex = 4
    Me.chkCtsRts.Text = "Activar control CTS/RTS"
    '
    'chkXonXoff
    '
    Me.chkXonXoff.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkXonXoff.Location = New System.Drawing.Point(136, 48)
    Me.chkXonXoff.Name = "chkXonXoff"
    Me.chkXonXoff.Size = New System.Drawing.Size(152, 16)
    Me.chkXonXoff.TabIndex = 3
    Me.chkXonXoff.Text = "Activar control Xon/Xoff"
    '
    'chkDsrDtr
    '
    Me.chkDsrDtr.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkDsrDtr.Location = New System.Drawing.Point(136, 24)
    Me.chkDsrDtr.Name = "chkDsrDtr"
    Me.chkDsrDtr.Size = New System.Drawing.Size(152, 16)
    Me.chkDsrDtr.TabIndex = 2
    Me.chkDsrDtr.Text = "Activar control DSR/DTR"
    '
    'chkRTS
    '
    Me.chkRTS.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkRTS.Location = New System.Drawing.Point(24, 48)
    Me.chkRTS.Name = "chkRTS"
    Me.chkRTS.Size = New System.Drawing.Size(88, 16)
    Me.chkRTS.TabIndex = 1
    Me.chkRTS.Text = "Activar RTS"
    '
    'chkDTR
    '
    Me.chkDTR.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.chkDTR.Location = New System.Drawing.Point(24, 24)
    Me.chkDTR.Name = "chkDTR"
    Me.chkDTR.Size = New System.Drawing.Size(88, 16)
    Me.chkDTR.TabIndex = 0
    Me.chkDTR.Text = "Activar DTR"
    '
    'cmdDefault
    '
    Me.cmdDefault.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdDefault.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.cmdDefault.Location = New System.Drawing.Point(16, 432)
    Me.cmdDefault.Name = "cmdDefault"
    Me.cmdDefault.Size = New System.Drawing.Size(75, 23)
    Me.cmdDefault.TabIndex = 5
    Me.cmdDefault.Text = "&Defecte"
    '
    'cmdOk
    '
    Me.cmdOk.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK
    Me.cmdOk.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.cmdOk.Location = New System.Drawing.Point(162, 432)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.Size = New System.Drawing.Size(75, 23)
    Me.cmdOk.TabIndex = 3
    Me.cmdOk.Text = "&Acceptar"
    '
    'cmdCancelar
    '
    Me.cmdCancelar.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCancelar.DialogResult = System.Windows.Forms.DialogResult.Cancel
    Me.cmdCancelar.FlatStyle = System.Windows.Forms.FlatStyle.System
    Me.cmdCancelar.Location = New System.Drawing.Point(242, 432)
    Me.cmdCancelar.Name = "cmdCancelar"
    Me.cmdCancelar.Size = New System.Drawing.Size(75, 23)
    Me.cmdCancelar.TabIndex = 4
    Me.cmdCancelar.Text = "&Cancelar"
    '
    'frmSerialDialog
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(330, 486)
    Me.Controls.Add(Me.cmdCancelar)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.cmdDefault)
    Me.Controls.Add(Me.GroupBox3)
    Me.Controls.Add(Me.GroupBox2)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.lblTitol)
    Me.Controls.Add(Me.StatusBar)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.HelpProvider.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.Topic)
    Me.Name = "frmSerialDialog"
    Me.HelpProvider.SetShowHelp(Me, True)
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Paràmetres conexió sèrie"
    CType(Me.sbpInfo, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpFormName, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpVersion, System.ComponentModel.ISupportInitialize).EndInit()
    CType(Me.sbpDummy, System.ComponentModel.ISupportInitialize).EndInit()
    Me.GroupBox1.ResumeLayout(False)
    Me.GroupBox1.PerformLayout()
    Me.GroupBox2.ResumeLayout(False)
    Me.GroupBox2.PerformLayout()
    Me.GroupBox3.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub frmSerialDialog_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'Me.Icon = AppData.AppIcon
    'Me.HelpProvider.HelpNamespace = AppData.HelpNameSpace
    'Me.HelpProvider.SetHelpKeyword(Me, "/pr_frmSerialDialog.htm")
    'lblTitol.BackColor = AppData.TitleBackColor
    'lblTitol.ForeColor = AppData.TitleForeColor

    'Me.HelpProvider.HelpNamespace = AppData.HelpNameSpace
    'Me.HelpProvider.SetHelpNavigator(Me, System.Windows.Forms.HelpNavigator.Topic)
    'Me.HelpProvider.SetHelpKeyword(Me, "/pr0Utils_SerialDialog.htm")

    sbpFormName.Text = "pr0Utils_SerialDialog"
    sbpVersion.Text = "1.0.1"

    If mSetDefaultValuesOnStartUp Then
      SetDefaults()
    End If

  End Sub

  Private Sub SetDefaults()

    cboPorts.SelectedItem = "COM1"

    cboVelocitat.SelectedIndex = 4
    cboParitat.SelectedIndex = 0
    cboBitsParada.SelectedIndex = 0
    cboBitsDades.SelectedIndex = 3

    chkDTR.Checked = True
    chkRTS.Checked = True

    chkDsrDtr.Checked = False
    chkXonXoff.Checked = False
    chkCtsRts.Checked = True
    chkRs485.Checked = False


  End Sub
  Private Sub cmdSortir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    Me.Close()
  End Sub

End Class

